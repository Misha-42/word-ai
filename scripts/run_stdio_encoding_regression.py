from __future__ import annotations

import io
import json
import sys
from typing import Any

import word_ai_mcp.server as server_module
from word_ai_mcp.server import run_stdio

QUERY = "Стерлитамак"


class RecordingServer:
    """Stands in for WordAiMcpServer and records the request as decoded."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        self.requests: list[Any] = []

    def handle(self, req: Any) -> Any:
        self.requests.append(req)
        query = req.get("params", {}).get("arguments", {}).get("query")
        return {"jsonrpc": "2.0", "id": req.get("id"), "result": {"echo": query}}


def _locale_stream(raw: io.BytesIO) -> io.TextIOWrapper:
    """Emulate a Windows ANSI console: the locale codepage, not UTF-8."""
    return io.TextIOWrapper(raw, encoding="cp1251")


def main() -> int:
    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "docx_search_text",
            "arguments": {"docx_path": "C:\\docs\\sample.docx", "query": QUERY},
        },
    }
    # MCP stdio is UTF-8: this is what every client puts on the wire.
    payload = (json.dumps(request, ensure_ascii=False) + "\n").encode("utf-8")

    stdout_raw = io.BytesIO()
    captured = RecordingServer()
    real_stdin, real_stdout = sys.stdin, sys.stdout
    real_server_cls = server_module.WordAiMcpServer
    server_module.WordAiMcpServer = lambda **kwargs: captured  # type: ignore[assignment]
    sys.stdin = _locale_stream(io.BytesIO(payload))
    sys.stdout = _locale_stream(stdout_raw)
    try:
        run_stdio(root=".", allow_write=True)
        sys.stdout.flush()
        response_bytes = stdout_raw.getvalue()
    finally:
        sys.stdin, sys.stdout = real_stdin, real_stdout
        server_module.WordAiMcpServer = real_server_cls  # type: ignore[assignment]

    assert captured.requests, "server received no request"
    received = captured.requests[0]["params"]["arguments"]["query"]
    assert received == QUERY, f"request decoded as {received!r}, expected {QUERY!r}"

    response = json.loads(response_bytes.decode("utf-8"))
    assert response["result"]["echo"] == QUERY, response

    print(json.dumps({"ok": True, "covered": ["stdio request decode", "stdio response encode"]}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
