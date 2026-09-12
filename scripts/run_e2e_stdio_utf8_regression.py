"""Regression: end-to-end stdio write path must survive real UTF-8 payloads.

Two transport bugs were fixed at the unit level:

- stdio decoded/encoded in the process locale (cp1251 on Russian Windows)
  instead of UTF-8, so non-ASCII arguments were silently corrupted on write
  and non-ASCII responses crashed the frozen binary. Covered mock-level by
  ``run_stdio_encoding_regression.py``.
- ``patchset`` arriving as a JSON string (clients that do not expand the
  tool-schema ``$ref``) was rejected with "patchset must be an object".
  Covered unit-level by ``run_validation_isolation_regression.py``.

Neither exercises a real server process writing a real document. This script
closes that gap: it spawns ``python -m word_ai_mcp.server`` over stdio, sends
a PatchSet **as a JSON string** (bug 2 end-to-end) whose replacement text
contains Cyrillic and cp1251-unrepresentable characters (``Σ → ✓`` - the
class that used to raise UnicodeEncodeError), applies it to a copy of
``examples/sample_contract.docx``, and verifies:

1. assess accepts the string payload (ok: true, no risks);
2. apply succeeds and the paragraph text in the output is byte-exact -
   no mojibake, no loss;
3. ``docx_search_text`` with a Cyrillic query finds the new text (the
   original bug-1 repro returned count: 0);
4. validate reports ok with changed_parts == ["word/document.xml"];
5. the source file is untouched (sha256 stable).

Run:  python scripts/run_e2e_stdio_utf8_regression.py   (from the repo root)
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

# The covered-case labels contain non-cp1251 characters (Σ→✓); on a Russian
# Windows console printing them would raise UnicodeEncodeError in this script
# itself - the very bug class under test. Same remedy as the server: UTF-8 out.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from word_ai_mcp.ooxml import list_paragraphs  # noqa: E402

SAMPLE = ROOT / "examples" / "sample_contract.docx"
# Cyrillic + characters outside cp1251: the payload class that used to corrupt
# (mojibake written with ok: true) or crash (UnicodeEncodeError) the transport.
NEW_TEXT = "Предмет: сквозная проверка Σ→✓ кодировки stdio."
SEARCH_QUERY = "сквозная проверка"


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class StdioClient:
    """Minimal newline-delimited JSON-RPC client speaking strict UTF-8."""

    def __init__(self, root: Path) -> None:
        # The target of this regression is the transport, not the write engine.
        # Default to the Python engine so the test does not require a .NET SDK
        # (engine parity is covered by run_dotnet_regression.py); an explicit
        # WORD_AI_ENGINE/WORD_AI_DOTNET_EXE in the caller env is respected.
        env = dict(os.environ)
        env.setdefault("WORD_AI_ENGINE", "python")
        self.proc = subprocess.Popen(
            [sys.executable, "-m", "word_ai_mcp.server", "--root", str(root)],
            cwd=str(ROOT),
            env=env,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        self.next_id = 0

    def _send(self, obj: dict) -> None:
        self.proc.stdin.write((json.dumps(obj, ensure_ascii=False) + "\n").encode("utf-8"))
        self.proc.stdin.flush()

    def _recv(self) -> dict:
        line = self.proc.stdout.readline()
        if not line:
            stderr = self.proc.stderr.read().decode("utf-8", "replace")
            raise RuntimeError(f"server closed stdout; stderr: {stderr[:2000]}")
        return json.loads(line.decode("utf-8"))  # strict UTF-8: a locale byte must fail here

    def call(self, tool: str, arguments: dict) -> dict:
        self.next_id += 1
        self._send({"jsonrpc": "2.0", "id": self.next_id, "method": "tools/call",
                    "params": {"name": tool, "arguments": arguments}})
        resp = self._recv()
        if "error" in resp:
            raise AssertionError(f"{tool} returned error: {resp['error']}")
        content = resp["result"].get("content", [])
        text = content[0].get("text", "") if content else ""
        return json.loads(text) if text.strip().startswith(("{", "[")) else {"raw": text}

    def close(self) -> None:
        self.proc.stdin.close()
        self.proc.terminate()


def main() -> int:
    covered: list[str] = []
    with TemporaryDirectory(prefix="word-ai-e2e-utf8-") as tmp_name:
        tmp = Path(tmp_name)
        source = tmp / "source.docx"
        output = tmp / "output.docx"
        shutil.copyfile(SAMPLE, source)
        source_sha = sha256_file(source)

        paragraph = next(
            p for p in list_paragraphs(source)["paragraphs"]
            if (p["text_preview"] or "").strip()
        )
        patchset = {
            "schema_version": "2.0",
            "strict": True,
            "source_sha256": source_sha,
            "reason": "e2e stdio UTF-8 regression",
            "guard": {"require_preconditions": True, "allow_overwrite": False},
            "operations": [{
                "op": "replace_paragraph_text",
                "paragraph_index": paragraph["paragraph_index"],
                "expected_old_sha256": paragraph["text_sha256"],
                "text": NEW_TEXT,
                "preserve_style": True,
            }],
        }

        client = StdioClient(tmp)
        try:
            # 1. string patchset (what $ref-blind clients send) is accepted
            assessment = client.call("docx_assess_patchset", {
                "docx_path": str(source),
                "patchset": json.dumps(patchset, ensure_ascii=False),
            })
            assert assessment["ok"] is True, f"assess rejected string patchset: {assessment}"
            assert not assessment.get("risks"), f"unexpected risks: {assessment['risks']}"
            covered.append("string patchset accepted end-to-end (assess)")

            # 2. apply writes the exact UTF-8 text
            client.call("docx_apply_patchset", {
                "docx_path": str(source),
                "patchset": patchset,
                "output_path": str(output),
            })
            assert output.is_file(), "apply produced no output file"
            written = client.call("docx_read_paragraph", {
                "docx_path": str(output),
                "paragraph_index": paragraph["paragraph_index"],
            })
            assert written["text"] == NEW_TEXT, (
                f"text corrupted through the transport: {written['text']!r} != {NEW_TEXT!r}"
            )
            covered.append("write path preserves Cyrillic and non-cp1251 characters (Σ→✓)")

            # 3. the original bug-1 repro: a Cyrillic search query must find the text
            search = client.call("docx_search_text", {
                "docx_path": str(output),
                "query": SEARCH_QUERY,
            })
            assert search["count"] >= 1, f"Cyrillic search found nothing: {search}"
            covered.append("Cyrillic docx_search_text query matches (was count: 0)")

            # 4. in-place contract
            validation = client.call("docx_validate", {
                "source_docx": str(source),
                "target_docx": str(output),
                "touched_paragraph_indices": [paragraph["paragraph_index"]],
            })
            assert validation["ok"] is True, f"validate failed: {validation.get('issues')}"
            assert validation["metrics"]["changed_parts"] == ["word/document.xml"], (
                f"more than document.xml changed: {validation['metrics']['changed_parts']}"
            )
            covered.append("validate ok, changed_parts == [word/document.xml]")

            # 5. source untouched
            assert sha256_file(source) == source_sha, "source file was modified"
            covered.append("source sha256 stable")
        finally:
            client.close()

    print(f"ok: {len(covered)} cases, end-to-end stdio UTF-8 write path holds")
    for label in covered:
        print(f"  - {label}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
