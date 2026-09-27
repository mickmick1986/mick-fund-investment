#!/usr/bin/env python3
"""Publish the latest local fund guide to the configured GitHub repository.

The script copies the newly generated guide into deploy/index.html, stages only
known project deliverables, then commits and pushes if there are changes.
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import shutil
import subprocess
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SOURCE_HTML = ROOT / "金字塔丛林补仓指导图.html"
PAGES_HTML = ROOT / "deploy" / "index.html"
TRACKED_PATHS = [
    ".github/workflows/deploy.yml",
    ".github/workflows/live-estimates.yml",
    ".gitignore",
    "git_auto_push.py",
    "generate_guide_html.py",
    "enrich_fund_data_v2.py",
    "fill_excel_combined.py",
    "fetch_live_estimates.py",
    "publish_live_snapshot.py",
    "rsi_dual_track.py",
    "generate_rsi_dual_track_html.py",
    "generate_rsi_dual_track_conclusion_html.py",
    "rsi_dual_track_validation.json",
    "rsi_dual_track_history.json",
    "rsi_dual_track_cycle_state.json",
    "fund_data_enriched.json",
    "金字塔丛林补仓指导图.html",
    "deploy/index.html",
    "deploy/rsi_dual_track_validation.json",
    "deploy/rsi_dual_track_comparison.html",
    "deploy/rsi_dual_track_20day_conclusion.html",
]


def run_git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    # 本机 Git 使用 Windows Schannel 时，当前网络无法检查证书吊销状态。
    # 仅对子进程这一次推送关闭校验；不写入 Git 全局或本地配置。
    env["GIT_SSL_NO_VERIFY"] = "true"
    return subprocess.run(
        ["git", *args], cwd=ROOT, text=True, encoding="utf-8", errors="replace",
        capture_output=True, check=check, env=env,
    )


GCM_CANDIDATES = (
    Path(r"C:/Users/13697/.workbuddy/binaries/PortableGit/versions/1.2.0/mingw64/bin/git-credential-manager.exe"),
)


def get_github_token() -> str | None:
    """优先读环境变量，回退到 Git Credential Manager 读取已保存的 GitHub 令牌。"""
    token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    if token:
        return token
    candidates = list(GCM_CANDIDATES)
    found = shutil.which("git-credential-manager.exe")
    if found:
        candidates.append(Path(found))
    for gcm in candidates:
        if not gcm.is_file():
            continue
        try:
            out = subprocess.run(
                [str(gcm), "get"],
                input="protocol=https\nhost=github.com\n\n",
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=15,
            ).stdout
            for line in out.splitlines():
                if line.startswith("password="):
                    return line.split("=", 1)[1]
        except (OSError, subprocess.TimeoutExpired):
            continue
    return None


API_BASE = "https://api.github.com"
API_HEADERS_USER_AGENT = "JinzitaGuidePublisher/2.0"


def _api_request(token: str, method: str, path: str, payload: dict | None = None, timeout: int = 60) -> dict:
    """调用 GitHub REST API。path 形如 /repos/{owner}/{repo}/git/ref/heads/main。"""
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "User-Agent": API_HEADERS_USER_AGENT,
        "X-GitHub-Api-Version": "2022-11-28",
    }
    if data is not None:
        headers["Content-Type"] = "application/json"
    request = urllib.request.Request(f"{API_BASE}{path}", data=data, headers=headers, method=method)
    with urllib.request.urlopen(request, timeout=timeout) as response:
        body = response.read().decode("utf-8")
    return json.loads(body) if body else {}


def get_repo_slug() -> str | None:
    """从本地 git remote 解析 owner/repo。"""
    result = run_git("remote", "get-url", "origin", check=False)
    if result.returncode != 0:
        return None
    repo = result.stdout.strip()
    for prefix in ("https://github.com/", "git@github.com:"):
        if repo.startswith(prefix):
            repo = repo[len(prefix):]
            break
    repo = repo.removesuffix(".git").strip("/")
    return repo or None


def git_blob_sha(content: bytes) -> str:
    """计算 git blob 的 SHA-1（与远端 tree 中的 sha 可直接比较）。"""
    header = f"blob {len(content)}\0".encode("utf-8")
    return hashlib.sha1(header + content).hexdigest()


def fetch_remote_state(token: str, repo: str, ref: str = "main") -> dict:
    """一次取回远端 ref 指向的 commit、tree 及全量文件 blob sha 映射。"""
    ref_info = _api_request(token, "GET", f"/repos/{repo}/git/ref/heads/{ref}")
    commit_sha = ref_info["object"]["sha"]
    commit = _api_request(token, "GET", f"/repos/{repo}/git/commits/{commit_sha}")
    tree_sha = commit["tree"]["sha"]
    tree = _api_request(token, "GET", f"/repos/{repo}/git/trees/{tree_sha}?recursive=1")
    files: dict[str, str] = {}
    if not tree.get("truncated"):
        files = {
            item["path"]: item["sha"]
            for item in tree.get("tree", [])
            if item.get("type") == "blob"
        }
    return {"commit": commit_sha, "tree": tree_sha, "files": files}


def publish_via_git_data_api(token: str, paths: list[str], message: str) -> tuple[bool, str]:
    """用 Git Data API 原子提交全部发布产物，完全绕开本地 git 对象。

    相比 Contents API 单文件直推的优势：
      - 一次提交多个文件，避免"页面更新了但 RSI 数据没更新"的半成品状态；
      - 与远端 tree 做 blob SHA 比对，无变化的文件自动跳过，不产生空提交；
      - 不读取本地 .git，天然免疫本地对象损坏问题。
    """
    repo = get_repo_slug()
    if not repo:
        return False, "无法解析远端仓库地址"
    try:
        state = fetch_remote_state(token, repo)
        entries = []
        for path in paths:
            local = ROOT / path
            if not local.is_file():
                continue
            content = local.read_bytes()
            if state["files"].get(path) == git_blob_sha(content):
                continue
            blob = _api_request(
                token, "POST", f"/repos/{repo}/git/blobs",
                {"content": base64.b64encode(content).decode("ascii"), "encoding": "base64"},
            )
            entries.append({"path": path, "mode": "100644", "type": "blob", "sha": blob["sha"]})
        if not entries:
            return True, "远端已是最新，无需推送"
        new_tree = _api_request(
            token, "POST", f"/repos/{repo}/git/trees",
            {"base_tree": state["tree"], "tree": entries},
        )["sha"]
        new_commit = _api_request(
            token, "POST", f"/repos/{repo}/git/commits",
            {"message": message, "tree": new_tree, "parents": [state["commit"]]},
        )["sha"]
        _api_request(token, "PATCH", f"/repos/{repo}/git/refs/heads/main", {"sha": new_commit})
        return True, f"已提交 {new_commit[:8]}（{len(entries)} 个文件更新）"
    except urllib.error.HTTPError as error:
        detail = ""
        try:
            detail = error.read().decode("utf-8")[:200]
        except OSError:
            pass
        return False, f"Git Data API 失败：HTTP {error.code} {detail}"
    except (OSError, ValueError, KeyError, json.JSONDecodeError) as error:
        return False, f"Git Data API 失败：{error}"


def push_html_via_contents_api(token: str) -> tuple[bool, str]:
    """用 GitHub Contents API 直推 deploy/index.html，绕开本地 git 对象损坏。"""
    remote_result = run_git("remote", "get-url", "origin", check=False)
    if remote_result.returncode != 0:
        return False, "无法读取远端地址"
    repo = remote_result.stdout.strip()
    for prefix in ("https://github.com/", "git@github.com:"):
        if repo.startswith(prefix):
            repo = repo[len(prefix):]
            break
    repo = repo.removesuffix(".git")
    url = f"https://api.github.com/repos/{repo}/contents/deploy/index.html?ref=main"
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "User-Agent": "JinzitaGuidePublisher/1.0",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    try:
        req = urllib.request.Request(url, headers=headers, method="GET")
        with urllib.request.urlopen(req, timeout=30) as response:
            current = json.loads(response.read().decode("utf-8"))
        sha = current.get("sha")
        if not sha:
            return False, "Contents 响应缺少 SHA"
        local = PAGES_HTML.read_bytes()
        remote_content = base64.b64decode(current.get("content", "").replace("\\n", ""))
        if remote_content == local:
            return True, "远端已是最新，无需推送"
        payload = {
            "message": f"chore: update guide html {datetime.now():%Y-%m-%d %H:%M}",
            "content": base64.b64encode(local).decode("ascii"),
            "sha": sha,
            "branch": "main",
        }
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        put_headers = dict(headers)
        put_headers["Content-Type"] = "application/json"
        req = urllib.request.Request(url.split("?")[0], data=body, headers=put_headers, method="PUT")
        with urllib.request.urlopen(req, timeout=60) as response:
            result = json.loads(response.read().decode("utf-8"))
        commit_sha = result.get("commit", {}).get("sha", "?")
        return True, f"已直推（commit {commit_sha[:8]}）"
    except (OSError, ValueError, json.JSONDecodeError, urllib.error.HTTPError) as error:
        return False, f"Contents API 直推失败：{error}"


def _commit_local_history(message: str) -> bool:
    """本地 git 仅作历史记录：提交并尝试推送，任何失败都不影响公开页。"""
    try:
        existing_paths = [path for path in TRACKED_PATHS if (ROOT / path).exists()]
        run_git("add", "--", *existing_paths)
        staged = run_git("diff", "--cached", "--quiet", check=False)
        if staged.returncode == 0:
            return True  # 无变化 = 本地已是最新
        if staged.returncode != 1:
            print(f"[git] diff 异常：{staged.stderr}", file=sys.stderr)
            return False
        committed = run_git("commit", "-m", message, check=False)
        if committed.returncode != 0:
            print(f"[git] 提交失败：{committed.stderr}", file=sys.stderr)
            return False
        run_git("fetch", "origin", "main", check=False)
        rebased = run_git("rebase", "origin/main", check=False)
        if rebased.returncode != 0:
            # 本地与远端分叉（例如 API 通道已发布同一批内容）：以远端为准，
            # 否则这条记录通道会长期卡在非快进状态。内容已由 API 通道发布，不会丢失。
            run_git("rebase", "--abort", check=False)
            run_git("reset", "--mixed", "origin/main", check=False)
            print("[git] 本地记录与远端分叉，已对齐远端（公开页不受影响）", file=sys.stderr)
            return False
        pushed = run_git("push", "origin", "main", check=False)
        if pushed.returncode != 0:
            print(f"[git] 推送失败（不影响公开页）：{pushed.stderr}", file=sys.stderr)
            return False
        print(f"[git] 已推送：{message}")
        return True
    except Exception as error:  # noqa: BLE001 - 本地 git 任何异常都不影响公开页
        print(f"[git] 本地 git 记录失败（不影响公开页）：{error}", file=sys.stderr)
        return False


def main() -> int:
    if not (ROOT / ".git").is_dir():
        print("未初始化 Git 仓库：请先完成首次 GitHub 仓库连接。", file=sys.stderr)
        return 2
    if not SOURCE_HTML.is_file():
        print(f"缺少待发布文件：{SOURCE_HTML.name}", file=sys.stderr)
        return 2

    # 独立双轨验证始终在正式页面发布前刷新：只读取确认层/Excel，绝不改写正式决策或止盈锚点。
    for script_name in (
        "rsi_dual_track.py",
        "generate_rsi_dual_track_html.py",
        "generate_rsi_dual_track_conclusion_html.py",
    ):
        generated = subprocess.run(
            [sys.executable, str(ROOT / script_name)], cwd=ROOT, text=True,
            encoding="utf-8", errors="replace", capture_output=True,
        )
        if generated.returncode != 0:
            print(f"双轨验证生成失败({script_name})：{generated.stdout}{generated.stderr}", file=sys.stderr)
            return generated.returncode

    PAGES_HTML.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE_HTML, PAGES_HTML)

    message = f"chore: 更新基金数据 {datetime.now():%Y-%m-%d %H:%M}"

    # 主通道：Git Data API 原子提交全部发布产物（主页面 + RSI 对比页 + RSI 数据）。
    # 2026-09-27 改造：原先只推 deploy/index.html，导致 RSI 页面/数据依赖本地 git，
    # 而本地 git 对象反复损坏 → 远端 RSI 对比页长期停留在旧版本。
    token = get_github_token()
    api_ok = False
    if token:
        api_ok, api_msg = publish_via_git_data_api(token, TRACKED_PATHS, message)
        print(f"[Git Data API] {api_msg}")
    else:
        print("[Git Data API] 未取到 GitHub 令牌，跳过（转本地 git 通道）", file=sys.stderr)

    # 二级回退：原子提交失败时，至少保住公开主页面。
    if not api_ok and token:
        contents_ok, contents_msg = push_html_via_contents_api(token)
        print(f"[Contents API 回退] {contents_msg}")
        api_ok = contents_ok

    # 历史记录通道：本地 git，失败不阻断（发布已由 API 通道保证）。
    git_ok = _commit_local_history(message)

    if api_ok or git_ok:
        return 0
    print("公开页更新失败：Git Data API、Contents API 回退与本地 git 均未成功。", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
