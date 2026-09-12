"""Regression: assessment must never bless a PatchSet the write path refuses.

`assess_patchset` resolved every target and then checked only that a precondition
was *present* (`is None`), never that it held. So it reported `ok: true, risks:
[]` for an `expected_old_sha256` of all zeroes, and the mismatch surfaced only
later as a dry-run or apply failure - the one answer assessment exists to give,
answered wrongly.

The reported case was a table: `docx_list_tables` offers both `text_sha256` and
`xml_sha256` per table, while every `expected_old_sha256` is compared against the
target's text, so the natural pick from the listing could never satisfy the
precondition. Assessment now compares through `_assert_expected_text` - the very
call the write path makes - so the two cannot drift apart again.

`apply_patchset` refuses whenever `assess_patchset` is not ok, so on this matrix
the two verdicts must be identical in both directions. The dangerous direction,
and the reason this is a regression rather than a nicety, is the other one:
assessment reporting `ok: true` for an edit the write path then refuses.
"""

from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from word_ai_mcp.ooxml import (
    apply_patchset,
    assess_patchset,
    list_paragraphs,
    list_tables,
    read_table,
)
from word_ai_mcp.server import WordAiMcpServer

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "examples" / "sample_contract.docx"
ZERO_SHA = "0" * 64

# label, operations, whether assessment should pass, the error code it should give
Case = tuple[str, list[dict], bool, str | None]


def main() -> int:
    server = WordAiMcpServer(root=".")
    paragraphs = list_paragraphs(SAMPLE)["paragraphs"]
    paragraph = next(p for p in paragraphs if (p["text_preview"] or "").strip())
    index, text, text_sha = paragraph["paragraph_index"], paragraph["text_preview"], paragraph["text_sha256"]

    table = list_tables(SAMPLE)["tables"][0]
    table_index, columns = table["table_index"], table["column_counts"][0]
    cell_text = read_table(SAMPLE, table_index)["rows"][0][0]["text"]
    appended_row = ["a"] * columns

    cases: list[Case] = [
        ("a correct expected_old_text is accepted",
         [{"op": "replace_paragraph_text", "paragraph_index": index, "text": "edited", "expected_old_text": text}], True, None),
        ("a correct expected_old_sha256 is accepted",
         [{"op": "replace_paragraph_text", "paragraph_index": index, "text": "edited", "expected_old_sha256": text_sha}], True, None),
        ("a sha256 of all zeroes is refused",
         [{"op": "replace_paragraph_text", "paragraph_index": index, "text": "edited", "expected_old_sha256": ZERO_SHA}], False, "precondition_mismatch"),
        ("a stale expected_old_text is refused",
         [{"op": "replace_paragraph_text", "paragraph_index": index, "text": "edited", "expected_old_text": "something else"}], False, "precondition_mismatch"),
        ("a stale anchor paragraph precondition is refused",
         [{"op": "insert_paragraph_after", "paragraph_index": index, "text": "inserted", "expected_old_text": "not the anchor"}], False, "precondition_mismatch"),
        ("a table's xml_sha256 cannot satisfy expected_old_sha256",
         [{"op": "append_table_row", "table_index": table_index, "values": appended_row, "expected_old_sha256": table["xml_sha256"]}], False, "precondition_mismatch"),
        ("a table's text_sha256 does satisfy it",
         [{"op": "append_table_row", "table_index": table_index, "values": appended_row, "expected_old_sha256": table["text_sha256"]}], True, None),
        ("a stale table cell precondition is refused",
         [{"op": "replace_table_cell_text", "table_index": table_index, "row": 1, "col": 1, "text": "edited", "expected_old_text": "not the cell text"}], False, "precondition_mismatch"),
        ("a correct table cell precondition is accepted",
         [{"op": "replace_table_cell_text", "table_index": table_index, "row": 1, "col": 1, "text": "edited", "expected_old_text": cell_text}], True, None),
        ("a missing precondition on a high-risk op is refused",
         [{"op": "replace_paragraph_text", "paragraph_index": index, "text": "edited"}], False, "missing_precondition"),
    ]

    covered: list[str] = []
    with TemporaryDirectory(prefix="word-ai-precondition-") as tmp:
        for number, (label, operations, expected_ok, expected_code) in enumerate(cases):
            patchset = {"schema_version": "2.0", "operations": operations}
            assessment = assess_patchset(SAMPLE, patchset)
            errors = [risk["code"] for risk in assessment["risks"] if risk["severity"] == "error"]

            assert assessment["ok"] is expected_ok, f"{label}: assessment said {assessment}, errors={errors}"
            if expected_code:
                assert expected_code in errors, f"{label}: expected {expected_code}, got {errors}"

            # Both directions. A case where apply fails for some other reason -
            # an output path that already exists, a failed post-write validation
            # - would be a legitimate exception, and none of these takes one.
            applied, refusal = attempt(SAMPLE, patchset, tmp, number)
            assert applied is expected_ok, f"{label}: assessment said {expected_ok}, the write path said {applied}"

            # The refusal wording is deliberately not compared here. Both
            # `apply_patchset` and `dry_run_patchset` run `assess_patchset` before
            # touching the document, so a mismatch is always caught there and the
            # `_assert_expected_text` calls inside the write path are unreachable
            # for this - comparing the two messages would compare assessment with
            # itself. Those calls still matter to the .NET backend's own writer,
            # which checks independently.
            if expected_code and expected_code != "precondition_mismatch":
                assert refusal, f"{label}: expected a refusal to report"
            covered.append(label)

    # The .NET backend answers `ok: true` for every one of these - it resolves
    # each target but never compares the precondition against it - so the facade
    # merges the verdict in. Exercised here through a stand-in for that backend,
    # which is all the facade needs and which keeps this script free of .NET.
    def dotnet_stand_in(_path: str, _patchset: dict, **_kw) -> dict:
        return {"ok": True, "risks": [], "engine": "dotnet"}

    good = {"schema_version": "2.0", "operations": [
        {"op": "replace_paragraph_text", "paragraph_index": index, "text": "edited", "expected_old_sha256": text_sha}]}
    bad = {"schema_version": "2.0", "operations": [
        {"op": "replace_paragraph_text", "paragraph_index": index, "text": "edited", "expected_old_sha256": ZERO_SHA}]}

    unchanged = server._with_precondition_verdict(str(SAMPLE), {"patchset": good}, dotnet_stand_in(str(SAMPLE), good))
    assert unchanged["ok"] is True and not unchanged["risks"], unchanged

    merged = server._with_precondition_verdict(str(SAMPLE), {"patchset": bad}, dotnet_stand_in(str(SAMPLE), bad))
    assert merged["ok"] is False, merged
    assert [risk["code"] for risk in merged["risks"]] == ["precondition_mismatch"], merged
    assert merged["engine"] == "dotnet", "the backend's own fields must survive the merge"

    print(f"ok: {len(covered)} cases, assessment and apply agree; the .NET path is corrected too")
    for label in covered:
        print(f"  - {label}")
    return 0


def attempt(source: Path, patchset: dict, tmp: str, number: int) -> tuple[bool, str]:
    """Run the write path on the patchset assessment just judged."""
    try:
        apply_patchset(source, patchset, Path(tmp) / f"case-{number}.docx")
    except Exception as exc:
        return False, str(exc)
    return True, ""


if __name__ == "__main__":
    raise SystemExit(main())
