"""Regression: a `paraId` anchor must be executable, not just declarable.

The .NET engine cannot resolve a `w14:paraId` write target - it fails with
`target_resolution_failed: Target paragraph not found` for an id the Python
engine resolves without complaint, for a paragraph the caller located by index,
and for an id the engine itself reports back in `touched.para_ids`. `paraId` is
a documented PatchSet anchor and `SKILL.md` ranks it above `paragraph_index`, so
the documented path was the one that could not be used.

The .NET engine is shipped prebuilt and cannot be rebuilt here, so the facade
translates the anchor into the paragraph index that engine does accept. These
assertions pin that translation: it must resolve the id against the source
document, let the id win when an operation carries both axes, cover paragraphs
inside tables, and leave an id it cannot find alone so the engine - not the
facade - reports the missing target.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path

from word_ai_mcp.ooxml import resolve_paraid_targets

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"

DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:document xmlns:w="{W_NS}" xmlns:w14="{W14_NS}"><w:body>'
    "<w:p><w:r><w:t>plain</w:t></w:r></w:p>"
    '<w:p w14:paraId="AAAAAAAA"><w:r><w:t>second</w:t></w:r></w:p>'
    "<w:tbl><w:tr><w:tc>"
    '<w:p w14:paraId="CCCCCCCC"><w:r><w:t>in table</w:t></w:r></w:p>'
    "</w:tc></w:tr></w:tbl>"
    '<w:p w14:paraId="BBBBBBBB"><w:r><w:t>fourth</w:t></w:r></w:p>'
    "</w:body></w:document>"
).encode("utf-8")


def _write_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", DOCUMENT_XML)


def main() -> int:
    covered: list[tuple[str, bool]] = []

    def covers(what: str, condition: bool) -> None:
        if not condition:
            raise AssertionError(what)
        covered.append((what, True))

    with tempfile.TemporaryDirectory(prefix="word-ai-paraid-") as tmp:
        docx = Path(tmp) / "source.docx"
        _write_docx(docx)

        def resolve(operations: list) -> dict:
            return resolve_paraid_targets(
                docx, {"schema_version": "2.0", "reason": "test", "operations": operations}
            )

        # 1. A paraId anchor becomes the index of that very paragraph.
        out = resolve([{"op": "replace_paragraph_text", "paraId": "BBBBBBBB"}])
        op = out["operations"][0]
        covers("resolves paraId to the paragraph's global index", op.get("paragraph_index") == 4)
        covers("drops the anchor the engine cannot resolve", "paraId" not in op)
        covers("keeps the operation and its other keys", op["op"] == "replace_paragraph_text")
        covers("keeps the patchset's other keys", out["reason"] == "test")

        # 2. Paragraphs inside a table share the //w:body//w:p axis.
        out = resolve([{"op": "replace_paragraph_text", "paraId": "CCCCCCCC"}])
        covers("resolves a paragraph inside a table", out["operations"][0]["paragraph_index"] == 3)

        # 3. The id wins when an operation carries both axes - which is what the
        #    Python engine already does, so the engines cannot disagree.
        out = resolve([{"op": "replace_paragraph_text", "paraId": "BBBBBBBB", "paragraph_index": 99}])
        covers("the paraId wins over a stale index", out["operations"][0]["paragraph_index"] == 4)

        # 4. An id that is not in the document is left for the engine to report,
        #    rather than silently redirected to some other paragraph.
        op = resolve([{"op": "replace_paragraph_text", "paraId": "DEADBEEF"}])["operations"][0]
        covers("leaves an unknown paraId untouched", op.get("paraId") == "DEADBEEF")
        covers("an unknown paraId gains no index", "paragraph_index" not in op)

        # 5. Operations that never needed translating are returned as they were.
        patchset = {"schema_version": "2.0", "operations": [{"op": "replace_paragraph_text", "paragraph_index": 2}]}
        covers("returns an index-addressed patchset unchanged", resolve_paraid_targets(docx, patchset) is patchset)
        covers("tolerates a patchset without operations", resolve_paraid_targets(docx, {"schema_version": "2.0"}) == {"schema_version": "2.0"})
        covers("tolerates a non-list operations value", resolve_paraid_targets(docx, {"operations": "x"}) == {"operations": "x"})

        # 6. A mixed patchset is translated per operation, and the caller's
        #    object is not mutated.
        operations = [
            {"op": "replace_paragraph_text", "paraId": "AAAAAAAA", "expected_old_text": "second"},
            {"op": "replace_paragraph_text", "paragraph_index": 1},
        ]
        out = resolve(operations)
        covers("translates only the anchored operation", out["operations"][0]["paragraph_index"] == 2)
        covers("preconditions survive the translation", out["operations"][0]["expected_old_text"] == "second")
        covers("leaves the index-addressed sibling alone", out["operations"][1] == operations[1])
        covers("does not mutate the caller's operations", operations[0].get("paraId") == "AAAAAAAA")

    print(json.dumps({"ok": True, "covered": [what for what, _ in covered]}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
