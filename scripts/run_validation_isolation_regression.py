"""Regression: an isolation claim must be complete, or validation must refuse.

Three defects are covered here.

1. Touched scopes were caller-supplied and replaced the audit-derived ones, so a
   hand-written half of the claim turned a correct edit into protected_* errors.
   The isolation checks match paragraphs by w14:paraId and body blocks by
   paragraph index, so a single-axis claim leaves the other check comparing
   against an empty set.

2. Supplying nothing and having no audit skipped the deep checks entirely while
   still reporting ok: true - a green report that proved nothing.

3. _canonical_hash used inclusive c14n, so a part that merely redeclares a
   namespace changed the hash of every element beneath it. The .NET backend
   hoists drawing namespaces, so validating a .NET-edited document with the
   Python backend reported every paragraph as changed.
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

from lxml import etree

from word_ai_mcp.ooxml import _canonical_hash
from word_ai_mcp.patchset import normalize_patchset
from word_ai_mcp.server import WordAiMcpServer

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

PARA_ID = "1810201C"
PARA_INDEX = 236


def _write_audit(target: Path) -> None:
    """A minimal sibling audit, as docx_apply_patchset leaves behind."""
    target.with_suffix(".audit.json").write_text(
        json.dumps(
            {
                "safety_assessment": {
                    "touched": {
                        "content_control_tags": [],
                        "para_ids": [PARA_ID],
                        "paragraph_indices": [PARA_INDEX],
                        "table_indices": [],
                        "table_cells": [],
                    }
                },
                "applied": [{"op": "replace_paragraph_text", "paragraph_index": PARA_INDEX}],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def _element(xml: bytes) -> etree._Element:
    return etree.fromstring(xml)


def main() -> int:
    server = WordAiMcpServer(root=".")
    covered: list[str] = []

    with tempfile.TemporaryDirectory(prefix="word-ai-isolation-") as tmp:
        target = Path(tmp) / "target.docx"
        target.write_bytes(b"placeholder")

        # 1. No claim and no audit: refuse rather than report an unchecked pass.
        try:
            server._validation_kwargs({"source_docx": "s.docx", "target_docx": str(target)}, str(target))
        except ValueError as exc:
            assert "isolation" in str(exc).lower(), exc
            covered.append("refuses when the isolation claim is unavailable")
        else:
            raise AssertionError("validation without an isolation claim should have been refused")

        # 2. Same, but the caller explicitly accepts an inconclusive result.
        kwargs, notes = server._validation_kwargs(
            {"source_docx": "s.docx", "target_docx": str(target), "allow_skipped_isolation_checks": True},
            str(target),
        )
        assert not any(kwargs[key] for key in kwargs if key.startswith("touched")), kwargs
        assert any(note["code"] == "isolation_checks_skipped" for note in notes), notes
        annotated = server._annotate_isolation({"ok": True, "issues": []}, kwargs, notes)
        assert annotated["isolation_checks"] == "skipped", annotated
        covered.append("reports isolation_checks=skipped instead of a silent pass")

        # 3. A single-axis claim must not replace the audit, which carries both axes.
        _write_audit(target)
        kwargs, notes = server._validation_kwargs(
            {"source_docx": "s.docx", "target_docx": str(target), "touched_para_ids": [PARA_ID]},
            str(target),
        )
        assert kwargs["touched_para_ids"] == [PARA_ID], kwargs
        assert kwargs["touched_paragraph_indices"] == [PARA_INDEX], kwargs
        assert any(note["code"] == "touched_merged_with_audit" for note in notes), notes
        covered.append("merges an explicit claim with the audit instead of replacing it")

        # 4. With no audit, a single-axis claim is diagnosed rather than left to
        #    surface as an unexplained protected_* failure.
        target.with_suffix(".audit.json").unlink()
        kwargs, notes = server._validation_kwargs(
            {"source_docx": "s.docx", "target_docx": str(target), "touched_para_ids": [PARA_ID]},
            str(target),
        )
        assert any(note["code"] == "touched_sets_incomplete" for note in notes), notes
        covered.append("diagnoses a single-axis isolation claim")

    # 5. Equivalent markup hashes equally however it was serialized.
    plain = _element(('<w:p xmlns:w="%s"><w:r><w:t>hi</w:t></w:r></w:p>' % W_NS).encode())
    hoisted = _element(
        ('<w:p xmlns:w="%s" xmlns:a="%s"><w:r><w:t>hi</w:t></w:r></w:p>' % (W_NS, A_NS)).encode()
    )
    changed = _element(('<w:p xmlns:w="%s"><w:r><w:t>bye</w:t></w:r></w:p>' % W_NS).encode())
    assert _canonical_hash(plain) == _canonical_hash(hoisted), (
        "an unused namespace declaration in scope must not change the hash"
    )
    assert _canonical_hash(plain) != _canonical_hash(changed), "a real text change must change the hash"
    covered.append("hashes equivalent markup equally (exclusive c14n)")

    # 6. A patchset arrives as a JSON string from clients that do not expand $ref.
    payload = {"schema_version": "2.0", "operations": [{"op": "replace_paragraph_text"}]}
    assert normalize_patchset(json.dumps(payload))["schema_version"] == "2.0"
    try:
        normalize_patchset("{not json")
    except ValueError as exc:
        assert "JSON object string" in str(exc), exc
    else:
        raise AssertionError("malformed patchset text should have been refused")
    try:
        normalize_patchset(123)
    except ValueError as exc:
        assert "must be an object" in str(exc), exc
    else:
        raise AssertionError("a non-object patchset should have been refused")
    covered.append("accepts a JSON-string patchset and still refuses junk")

    print(json.dumps({"ok": True, "covered": covered}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
