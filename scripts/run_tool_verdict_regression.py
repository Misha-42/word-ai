"""Regression: a tool's verdict must be where the caller looks for it.

Two smaller inconsistencies of the same shape as the assessment bug - the tool
knew the answer and did not put it where the caller reads.

`docx_dry_run_patchset` carried its verdict only inside `safety_assessment.ok`
and `validation.ok`, while `docx_assess_patchset`, `docx_validate` and
`docx_compare_structure` each answer with a top-level `ok`. A caller checking
`result["ok"]`, as it does for the others, met a KeyError and read a successful
dry run as a failure.

A tool reads its required arguments as `args["docx_path"]`, so a caller who
guesses the parameter name - `path`, say - was answered with `KeyError:
'docx_path'` and a Python traceback, which says nothing about what the tool
accepts.
"""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from tempfile import TemporaryDirectory

from word_ai_mcp.ooxml import dry_run_patchset, dry_run_verdict, list_paragraphs
from word_ai_mcp.server import WordAiMcpServer

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "examples" / "sample_contract.docx"


def main() -> int:
    server = WordAiMcpServer(root=".")
    covered: list[str] = []

    # 1. The verdict is the conjunction of the two sub-verdicts.
    assert dry_run_verdict({"safety_assessment": {"ok": True}, "validation": {"ok": True}})["ok"] is True
    assert dry_run_verdict({"safety_assessment": {"ok": True}, "validation": {"ok": False}})["ok"] is False
    assert dry_run_verdict({"safety_assessment": {"ok": False}, "validation": {"ok": True}})["ok"] is False
    covered.append("dry_run_verdict combines both sub-verdicts")

    # 2. An absent sub-verdict is not a pass: an unproven verdict is the one
    #    answer this tool must never give.
    assert dry_run_verdict({"safety_assessment": {"ok": True}})["ok"] is False
    assert dry_run_verdict({})["ok"] is False
    covered.append("a missing sub-verdict is not treated as a pass")

    # 3. A real dry run now reports it at the top level, without disturbing the
    #    fields that were already there.
    with TemporaryDirectory(prefix="word-ai-verdict-") as tmp:
        sample = Path(tmp) / "sample.docx"
        shutil.copyfile(SAMPLE, sample)
        paragraph = next(p for p in list_paragraphs(sample)["paragraphs"] if (p["text_preview"] or "").strip())
        result = dry_run_patchset(sample, {"schema_version": "2.0", "operations": [
            {"op": "replace_paragraph_text", "paragraph_index": paragraph["paragraph_index"],
             "text": "edited", "expected_old_text": paragraph["text_preview"]}]}, keep_output=False)
        assert result["ok"] is True, result
        assert result["dry_run"] is True and result["safety_assessment"]["ok"] is True, result
        covered.append("a real dry run reports a top-level ok")

    # 4. A missing required argument is named, not left as a bare KeyError.
    #    Asserted through `handle`, which is where the caller actually arrives -
    #    a test of the helper alone would not notice the helper going unwired.
    response = server.handle({
        "jsonrpc": "2.0", "id": 1, "method": "tools/call",
        "params": {"name": "docx_search_text", "arguments": {"path": "some.docx", "query": "text"}},
    })
    message = (response.get("error") or {}).get("message", "")
    assert "'docx_path'" in message, response
    assert "query" in message, "the tool's own accepted arguments should be named"
    covered.append("a missing argument is named with the tool's own schema")

    # 5. A KeyError from a tool's own data lookup is not the caller's mistake, so
    #    it must keep its traceback rather than be reported as a bad argument.
    assert server._missing_argument("docx_search_text", KeyError("some_internal_key")) is None
    covered.append("an internal KeyError is not reported as a bad argument")

    print(json.dumps({"ok": True, "covered": covered}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
