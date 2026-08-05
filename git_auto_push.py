#!/usr/bin/env python3
"""Publish the latest local fund guide to the configured GitHub repository.

The script copies the newly generated guide into deploy/index.html, stages only
known project deliverables, then commits and pushes if there are changes.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
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
    "fund_data_enriched.json",
    "金字塔丛林补仓指导图.html",
    "deploy/index.html",
    "deploy/live_estimates.json",
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


def main() -> int:
    if not (ROOT / ".git").is_dir():
        print("未初始化 Git 仓库：请先完成首次 GitHub 仓库连接。", file=sys.stderr)
        return 2
    if not SOURCE_HTML.is_file():
        print(f"缺少待发布文件：{SOURCE_HTML.name}", file=sys.stderr)
        return 2

    PAGES_HTML.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE_HTML, PAGES_HTML)
    run_git("add", "--", *TRACKED_PATHS)

    staged = run_git("diff", "--cached", "--quiet", check=False)
    if staged.returncode == 0:
        print("数据无变化，跳过 Git 提交与推送。")
        return 0
    if staged.returncode != 1:
        print(staged.stderr, file=sys.stderr)
        return staged.returncode

    message = f"chore: 更新基金数据 {datetime.now():%Y-%m-%d %H:%M}"
    committed = run_git("commit", "-m", message, check=False)
    if committed.returncode != 0:
        print(committed.stdout + committed.stderr, file=sys.stderr)
        return committed.returncode

    # 先拉取远程更新，避免非快进推送被拒
    fetched = run_git("fetch", "origin", "main", check=False)
    if fetched.returncode != 0:
        print(fetched.stdout + fetched.stderr, file=sys.stderr)
        # fetch 失败不阻断，继续尝试推送

    # 如果远程有新提交，先 rebase 到远程最新
    rebased = run_git("rebase", "origin/main", check=False)
    if rebased.returncode != 0:
        print("rebase 失败，尝试中止并继续：" + rebased.stdout + rebased.stderr, file=sys.stderr)
        run_git("rebase", "--abort", check=False)

    pushed = run_git("push", "origin", "main", check=False)
    if pushed.returncode != 0:
        print(pushed.stdout + pushed.stderr, file=sys.stderr)
        return pushed.returncode

    print(f"已推送：{message}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
