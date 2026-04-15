"""
Wait until this app's FastAPI server responds on /api/health, then open the web UI.

Used by start_server.bat so the browser never opens before uvicorn is ready,
and we only treat a successful health JSON as "this server".
"""
from __future__ import annotations

import argparse
import json
import sys
import time
import urllib.error
import urllib.request
import webbrowser


def _fetch_health(url: str, timeout: float) -> dict | None:
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            if r.status != 200:
                return None
            raw = r.read(4096).decode("utf-8", errors="replace")
    except (urllib.error.URLError, OSError, ValueError):
        return None
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return None
    if not isinstance(data, dict):
        return None
    if data.get("status") != "ok":
        return None
    if not isinstance(data.get("version"), str):
        return None
    return data


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("host", nargs="?", default="127.0.0.1", help="Bind host (use 127.0.0.1, not localhost)")
    p.add_argument("port", nargs="?", type=int, default=8000)
    p.add_argument("--max-wait", type=float, default=90.0)
    p.add_argument("--interval", type=float, default=0.4)
    p.add_argument("--health-timeout", type=float, default=2.0)
    args = p.parse_args()

    base = f"http://{args.host}:{args.port}"
    health_url = f"{base}/api/health"
    deadline = time.monotonic() + args.max_wait

    while time.monotonic() < deadline:
        if _fetch_health(health_url, args.health_timeout):
            opened = webbrowser.open(f"{base}/")
            if opened:
                print(f"[OK] 伺服器已就緒，已開啟瀏覽器: {base}/", flush=True)
            else:
                print(f"[OK] 伺服器已就緒，請手動開啟: {base}/", flush=True)
            return 0
        time.sleep(args.interval)

    print(
        f"[錯誤] 在 {args.max_wait:.0f} 秒內無法連上本專案的 /api/health: {health_url}",
        file=sys.stderr,
        flush=True,
    )
    print("[錯誤] 請確認 uvicorn 已啟動、連接埠正確，或先執行 stop_server.bat 後再試。", file=sys.stderr, flush=True)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
