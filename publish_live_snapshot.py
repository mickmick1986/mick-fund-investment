#!/usr/bin/env python3
"""Publish a validated intraday snapshot from a clean clone of origin/main."""
from __future__ import annotations

import json
import os
import shutil
from datetime import datetime
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SNAPSHOT_PATH = ROOT / "deploy" / "live_estimates.json"
REMOTE_NAME = "origin"
REMOTE_REF = "main"
SNAPSHOT_REPO_PATH = "deploy/live_estimates.json"
EXPECTED_ESTIMATE_COUNT = 25
MAX_PUSH_ATTEMPTS = 3
STATUS_DIR = ROOT / ".workbuddy"
STATUS_PATH = STATUS_DIR / "live_estimates_status.json"
STATUS_HISTORY_PATH = STATUS_DIR / "live_estimates_status.jsonl"


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
    env["GIT_SSL_NO_VERIFY"] = "true"
    return subprocess.run(
        command,
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
    if not snapshot.get("trade_date") or not snapshot.get("updated_at"):
        raise ValueError("snapshot is missing date metadata")
    if snapshot.get("estimate_count") != EXPECTED_ESTIMATE_COUNT:
        raise ValueError("snapshot estimate_count is invalid")
    if len(estimates) != EXPECTED_ESTIMATE_COUNT:
        raise ValueError("snapshot estimates are incomplete")
    if not isinstance(turnover, (int, float)):
        raise ValueError("snapshot turnover is missing")
    return snapshot


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
            print(f"push raced with another publisher; retrying ({attempt}/{MAX_PUSH_ATTEMPTS})", file=sys.stderr)
    return 1, False


def main() -> int:
    write_status("start", "started", snapshot_path=str(SNAPSHOT_PATH))
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

    result, published = publish_from_clean_clone(remote_result.stdout.strip(), snapshot)
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
