#!/usr/bin/env python3
"""Publish a validated intraday snapshot from a clean clone of origin/main."""
from __future__ import annotations

import base64
import json
import os
import shutil
from datetime import datetime, timedelta, timezone
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SNAPSHOT_PATH = ROOT / "deploy" / "live_estimates.json"
REMOTE_NAME = "origin"
REMOTE_REF = "main"
SNAPSHOT_REPO_PATH = "deploy/live_estimates.json"
CHINA_TZ = timezone(timedelta(hours=8))
EXPECTED_ESTIMATE_COUNT = 25
MAX_SNAPSHOT_AGE_MINUTES = 20
MAX_PUSH_ATTEMPTS = 3
CONTENTS_API_TIMEOUT = 30
STATUS_DIR = ROOT / ".workbuddy"
STATUS_PATH = STATUS_DIR / "live_estimates_status.json"
STATUS_HISTORY_PATH = STATUS_DIR / "live_estimates_status.jsonl"


def resolve_git_executable() -> str:
    """Prefer GitHub Desktop's Git so scheduled publishing reuses its login."""
    roots = []
    if os.environ.get("LOCALAPPDATA"):
        roots.append(Path(os.environ["LOCALAPPDATA"]) / "GitHubDesktop")
    roots.append(Path.home() / "AppData" / "Local" / "GitHubDesktop")
    for desktop_root in roots:
        desktop_git = sorted(
            desktop_root.glob("app-*/resources/app/git/cmd/git.exe"), reverse=True
        )
        if desktop_git:
            return str(desktop_git[0])
    return shutil.which("git") or "git"


GIT_EXECUTABLE = resolve_git_executable()


def write_status(stage: str, outcome: str, **details: object) -> None:
    """Keep local machine-readable diagnostics outside the published data path."""
    payload = {
        "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
        "component": "publish_live_snapshot",
        "stage": stage,
        "outcome": outcome,
        **details,
    }
    STATUS_DIR.mkdir(parents=True, exist_ok=True)
    STATUS_PATH.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    with STATUS_HISTORY_PATH.open("a", encoding="utf-8") as history:
        history.write(json.dumps(payload, ensure_ascii=False) + "\n")


def run(command: list[str], cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    # GitHub Desktop keeps the user's OAuth session; use its Git client for
    # scheduled pushes instead of the sandbox PortableGit credential helper.
    env["GIT_SSL_NO_VERIFY"] = "true"
    if command and command[0] == "git":
        actual_command = [
            GIT_EXECUTABLE,
            "-c",
            "credential.helper=manager",
            "-c",
            "http.sslVerify=false",
            *command[1:],
        ]
    else:
        actual_command = command
    return subprocess.run(
        actual_command,
        cwd=cwd,
        env=env,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )


def command_output(result: subprocess.CompletedProcess[str]) -> str:
    return (result.stdout + result.stderr).strip()


def validate_snapshot(path: Path) -> dict:
    snapshot = json.loads(path.read_text(encoding="utf-8"))
    estimates = snapshot.get("estimates") or {}
    turnover = (snapshot.get("sse_index") or {}).get("market_turnover_trillion")
    trade_date = snapshot.get("trade_date")
    updated_at = snapshot.get("updated_at")
    if not trade_date or not updated_at:
        raise ValueError("snapshot is missing date metadata")
    try:
        now = datetime.now(CHINA_TZ)
        expected_today = now.date().isoformat()
        parsed_updated_at = datetime.strptime(updated_at, "%Y-%m-%d %H:%M:%S")
    except ValueError as error:
        raise ValueError("snapshot updated_at format is invalid") from error
    if trade_date != expected_today or parsed_updated_at.date().isoformat() != trade_date:
        raise ValueError("snapshot trade_date or updated_at is not today")
    snapshot_age = now - parsed_updated_at.replace(tzinfo=CHINA_TZ)
    if not timedelta(0) <= snapshot_age <= timedelta(minutes=MAX_SNAPSHOT_AGE_MINUTES):
        raise ValueError("snapshot updated_at is outside the publish freshness window")
    if snapshot.get("estimate_count") != EXPECTED_ESTIMATE_COUNT:
        raise ValueError("snapshot estimate_count is invalid")
    if len(estimates) != EXPECTED_ESTIMATE_COUNT:
        raise ValueError("snapshot estimates are incomplete")
    if not isinstance(turnover, (int, float)) or turnover <= 0:
        raise ValueError("snapshot turnover is invalid")
    return snapshot


def get_github_repo(remote: str) -> str:
    if remote.startswith("https://github.com/"):
        return remote.removeprefix("https://github.com/").removesuffix(".git")
    if remote.startswith("git@github.com:"):
        return remote.removeprefix("git@github.com:").removesuffix(".git")
    raise ValueError(f"unsupported remote: {remote}")


def get_raw_url(remote: str) -> str:
    if remote.startswith("https://github.com/"):
        repo = remote.removeprefix("https://github.com/").removesuffix(".git")
    elif remote.startswith("git@github.com:"):
        repo = remote.removeprefix("git@github.com:").removesuffix(".git")
    else:
        raise ValueError(f"unsupported remote: {remote}")
    return f"https://raw.githubusercontent.com/{repo}/{REMOTE_REF}/{SNAPSHOT_REPO_PATH}"


def read_snapshot(url: str) -> dict:
    request = urllib.request.Request(
        url,
        headers={"Cache-Control": "no-cache", "User-Agent": "JinzitaLivePublisher/1.0"},
    )
    with urllib.request.urlopen(request, timeout=20) as response:
        return json.loads(response.read().decode("utf-8"))


def contents_request(url: str, token: str, payload: dict | None = None) -> dict:
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "User-Agent": "JinzitaLivePublisher/1.0",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    body = None
    method = "GET"
    if payload is not None:
        method = "PUT"
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json"
    request = urllib.request.Request(url, data=body, headers=headers, method=method)
    with urllib.request.urlopen(request, timeout=CONTENTS_API_TIMEOUT) as response:
        return json.loads(response.read().decode("utf-8"))


def publish_via_contents_api(remote: str, snapshot: dict, token: str) -> bool:
    """Update exactly one GitHub Contents API path with an optimistic SHA check."""
    repo = get_github_repo(remote)
    url = f"https://api.github.com/repos/{repo}/contents/{SNAPSHOT_REPO_PATH}?ref={REMOTE_REF}"
    encoded = base64.b64encode(
        json.dumps(snapshot, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    ).decode("ascii")
    last_error = ""
    for attempt in range(1, MAX_PUSH_ATTEMPTS + 1):
        try:
            current = contents_request(url, token)
            sha = current.get("sha")
            if not sha:
                raise RuntimeError("GitHub Contents response is missing target SHA")
            current_content = base64.b64decode(current.get("content", "").replace("\\n", ""))
            if current_content == base64.b64decode(encoded):
                write_status("contents_api", "already_current", path=SNAPSHOT_REPO_PATH)
                return False
            result = contents_request(
                url,
                token,
                {
                    "message": f"chore: update intraday estimates {snapshot['updated_at']}",
                    "content": encoded,
                    "sha": sha,
                    "branch": REMOTE_REF,
                },
            )
            if result.get("content", {}).get("path") != SNAPSHOT_REPO_PATH:
                raise RuntimeError("GitHub Contents response path mismatch")
            write_status(
                "contents_api",
                "success",
                path=SNAPSHOT_REPO_PATH,
                remote_sha=sha,
                commit_sha=result.get("commit", {}).get("sha"),
            )
            return True
        except (OSError, ValueError, json.JSONDecodeError, RuntimeError, urllib.error.HTTPError) as error:
            last_error = str(error)
            if attempt < MAX_PUSH_ATTEMPTS:
                write_status("contents_api", "retry", attempt=attempt, error=last_error)
                continue
    write_status("contents_api", "failed", error=last_error, path=SNAPSHOT_REPO_PATH)
    raise RuntimeError(f"GitHub Contents API publish failed: {last_error}")


def verify_public_snapshot(raw_url: str, expected_updated_at: str) -> None:
    deadline = time.monotonic() + 120
    last_seen = None
    while time.monotonic() < deadline:
        try:
            last_seen = read_snapshot(f"{raw_url}?_={int(time.time() * 1000)}").get("updated_at")
            if last_seen == expected_updated_at:
                write_status("raw_readback", "success", expected_updated_at=expected_updated_at)
                return
        except (OSError, ValueError, json.JSONDecodeError) as error:
            last_seen = f"error:{error}"
        time.sleep(10)
    write_status(
        "raw_readback",
        "failed",
        expected_updated_at=expected_updated_at,
        last_seen=last_seen,
        reason="timeout",
    )
    raise RuntimeError(f"remote snapshot did not update within 120 seconds; last={last_seen}")


def fail(result: subprocess.CompletedProcess[str]) -> int:
    print(command_output(result), file=sys.stderr)
    return result.returncode or 1


def publish_from_clean_clone(remote: str, snapshot: dict) -> tuple[int, bool]:
    """Publish only the live snapshot, retrying a competing main-branch push."""
    for attempt in range(1, MAX_PUSH_ATTEMPTS + 1):
        with tempfile.TemporaryDirectory(prefix="jinzita-live-publish-") as temp_dir:
            clone_dir = Path(temp_dir) / "repo"
            cloned = run(["git", "clone", "--depth", "1", "--branch", REMOTE_REF, remote, str(clone_dir)])
            if cloned.returncode:
                return fail(cloned), False

            target = clone_dir / SNAPSHOT_REPO_PATH
            shutil.copy2(SNAPSHOT_PATH, target)
            changed = run(["git", "diff", "--quiet", "--", SNAPSHOT_REPO_PATH], clone_dir)
            if changed.returncode == 0:
                return 0, False
            if changed.returncode != 1:
                return fail(changed), False

            commands = (
                ["git", "config", "user.name", "Jinzita Live Publisher"],
                ["git", "config", "user.email", "jinzita-live-publisher@users.noreply.github.com"],
                ["git", "add", "--", SNAPSHOT_REPO_PATH],
                ["git", "commit", "-m", f"chore: update intraday estimates {snapshot['updated_at']}"],
            )
            for command in commands:
                result = run(command, clone_dir)
                if result.returncode:
                    return fail(result), False

            pushed = run(["git", "push", REMOTE_NAME, f"HEAD:{REMOTE_REF}"], clone_dir)
            if pushed.returncode == 0:
                return 0, True
            if attempt == MAX_PUSH_ATTEMPTS:
                return fail(pushed), False
            reason = command_output(pushed)
            write_status(
                "push_attempt",
                "failed",
                attempt=attempt,
                trade_date=snapshot["trade_date"],
                updated_at=snapshot["updated_at"],
                error=reason,
            )
            print(f"push failed; retrying ({attempt}/{MAX_PUSH_ATTEMPTS}): {reason}", file=sys.stderr)
    return 1, False


def main() -> int:
    write_status(
        "start",
        "started",
        snapshot_path=str(SNAPSHOT_PATH),
        git_executable=GIT_EXECUTABLE,
    )
    try:
        snapshot = validate_snapshot(SNAPSHOT_PATH)
    except (OSError, ValueError, json.JSONDecodeError) as error:
        write_status("validate", "failed", error=str(error))
        print(f"snapshot validation failed: {error}", file=sys.stderr)
        return 2

    write_status(
        "validate",
        "success",
        trade_date=snapshot["trade_date"],
        updated_at=snapshot["updated_at"],
        estimate_count=snapshot["estimate_count"],
        turnover_trillion=(snapshot.get("sse_index") or {}).get("market_turnover_trillion"),
    )

    remote_result = run(["git", "remote", "get-url", REMOTE_NAME], ROOT)
    if remote_result.returncode:
        write_status("remote", "failed", error=command_output(remote_result))
        return fail(remote_result)
    try:
        raw_url = get_raw_url(remote_result.stdout.strip())
    except ValueError as error:
        write_status("remote", "failed", error=str(error))
        print(str(error), file=sys.stderr)
        return 2

    remote = remote_result.stdout.strip()
    github_token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    if github_token:
        try:
            published = publish_via_contents_api(remote, snapshot, github_token)
            result = 0
        except RuntimeError as error:
            write_status("push", "failed", error=str(error), path=SNAPSHOT_REPO_PATH)
            print(str(error), file=sys.stderr)
            return 1
    else:
        result, published = publish_from_clean_clone(remote, snapshot)
    if result:
        write_status(
            "push",
            "failed",
            trade_date=snapshot["trade_date"],
            updated_at=snapshot["updated_at"],
            return_code=result,
        )
        return result
    if not published:
        write_status(
            "push",
            "already_current",
            trade_date=snapshot["trade_date"],
            updated_at=snapshot["updated_at"],
        )
        print(f"remote snapshot is already current: {snapshot['updated_at']}")
        return 0

    write_status(
        "push",
        "success",
        trade_date=snapshot["trade_date"],
        updated_at=snapshot["updated_at"],
    )

    try:
        verify_public_snapshot(raw_url, snapshot["updated_at"])
    except RuntimeError as error:
        print(str(error), file=sys.stderr)
        return 3

    write_status(
        "complete",
        "success",
        trade_date=snapshot["trade_date"],
        updated_at=snapshot["updated_at"],
        estimate_count=snapshot["estimate_count"],
    )
    print(
        f"public snapshot updated: trade_date={snapshot['trade_date']} "
        f"updated_at={snapshot['updated_at']} estimates={snapshot['estimate_count']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
