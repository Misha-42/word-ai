from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Any, Callable

from . import __version__
from .ooxml import (
    apply_patchset,
    assess_patchset,
    backup_docx,
    build_document_map,
    compare_structure,
    diff_text,
    dry_run_patchset,
    export_plain_text,
    export_table_csv,
    extract_comments,
    extract_plain_text,
    get_content_control_text,
    get_outline,
    health_check,
    inspect_docx,
    list_anchors,
    list_bookmarks,
    list_comments,
    list_content_controls,
    list_docx_parts,
    list_fields,
    list_headings,
    list_headers_footers,
    list_hyperlinks,
    list_images,
    list_notes,
    list_numbering,
    list_paragraphs,
    list_revisions,
    list_sections,
    list_styles,
    list_tables,
    package_manifest,
    part_hashes,
    plan_patchset,
    read_anchor,
    read_heading_section,
    read_paragraph,
    read_table,
    read_table_cell,
    restore_backup,
    rollback_docx,
    search_text,
    structural_fingerprint,
    table_to_csv,
    validate_structure,
    write_sidecar_index,
)
from .openxml_engine import (
    dotnet_apply_patchset,
    dotnet_assess_patchset,
    dotnet_dry_run_patchset,
    dotnet_status,
    dotnet_validate,
    mark_python_result,
    select_engine,
)
from .patchset import normalize_patchset
from .session_store import (
    enqueue_command,
    get_command,
    list_sessions,
    session_summary,
    wait_for_command,
)

JSON = dict[str, Any]


def _text_payload(obj: Any) -> list[dict[str, str]]:
    text = obj if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False, indent=2)
    return [{"type": "text", "text": text}]


def _schema(props: dict[str, Any], required: list[str] | None = None) -> JSON:
    return {"type": "object", "properties": props, "required": required or [], "additionalProperties": False}


def _split_path_values(values: list[str] | None) -> list[str]:
    out: list[str] = []
    for value in values or []:
        for item in str(value).split(os.pathsep):
            item = item.strip()
            if item:
                out.append(item)
    return out


def _normalize_root(value: str | Path) -> Path:
    return Path(os.path.expandvars(str(value))).expanduser().resolve()


class WordAiMcpServer:
    def __init__(self, root: str | None = None, allow_write: bool = True, allowed_roots: list[str] | None = None):
        self.root = _normalize_root(root or os.getcwd())
        self.allow_write = allow_write
        configured_roots = _split_path_values(allowed_roots) + _split_path_values([os.environ.get("WORD_AI_ALLOWED_ROOTS", "")])
        roots = [self.root, *[_normalize_root(p) for p in configured_roots]]
        self.allowed_roots: list[Path] = []
        for allowed in roots:
            if allowed not in self.allowed_roots:
                self.allowed_roots.append(allowed)
        self.tools: dict[str, Callable[[JSON], Any]] = {
            # Package / health / map
            "docx_health_check": self.tool_docx_health_check,
            "docx_inspect": self.tool_docx_inspect,
            "docx_package_manifest": self.tool_docx_package_manifest,
            "docx_list_parts": self.tool_docx_list_parts,
            "docx_part_hashes": self.tool_docx_part_hashes,
            "docx_structural_fingerprint": self.tool_docx_structural_fingerprint,
            "docx_map": self.tool_docx_map,
            # Navigation / retrieval
            "docx_get_outline": self.tool_docx_get_outline,
            "docx_outline": self.tool_docx_get_outline,
            "docx_list_headings": self.tool_docx_list_headings,
            "docx_read_heading_section": self.tool_docx_read_heading_section,
            "docx_list_anchors": self.tool_docx_list_anchors,
            "docx_read_anchor": self.tool_docx_read_anchor,
            "docx_search_text": self.tool_docx_search_text,
            "docx_list_paragraphs": self.tool_docx_list_paragraphs,
            "docx_read_paragraph": self.tool_docx_read_paragraph,
            "docx_list_content_controls": self.tool_docx_list_content_controls,
            "docx_read_content_control": self.tool_docx_read_content_control,
            "docx_extract_plain_text": self.tool_docx_extract_plain_text,
            "docx_export_plain_text": self.tool_docx_export_plain_text,
            # Tables / styles / complex Word objects
            "docx_list_tables": self.tool_docx_list_tables,
            "docx_read_table": self.tool_docx_read_table,
            "docx_read_table_cell": self.tool_docx_read_table_cell,
            "docx_export_table_csv": self.tool_docx_export_table_csv,
            "docx_table_to_csv": self.tool_docx_table_to_csv,
            "docx_list_styles": self.tool_docx_list_styles,
            "docx_list_numbering": self.tool_docx_list_numbering,
            "docx_list_sections": self.tool_docx_list_sections,
            "docx_list_fields": self.tool_docx_list_fields,
            "docx_list_images": self.tool_docx_list_images,
            "docx_list_hyperlinks": self.tool_docx_list_hyperlinks,
            "docx_list_comments": self.tool_docx_list_comments,
            "docx_extract_comments": self.tool_docx_extract_comments,
            "docx_list_bookmarks": self.tool_docx_list_bookmarks,
            "docx_list_headers_footers": self.tool_docx_list_headers_footers,
            "docx_list_notes": self.tool_docx_list_notes,
            "docx_list_revisions": self.tool_docx_list_revisions,
            # Patch / write / validation lifecycle
            "docx_assess_patchset": self.tool_docx_assess_patchset,
            "docx_plan_patchset": self.tool_docx_plan_patchset,
            "docx_preflight_patchset": self.tool_docx_preflight_patchset,
            "docx_dry_run_patchset": self.tool_docx_dry_run_patchset,
            "docx_write_index": self.tool_docx_write_index,
            "docx_backup": self.tool_docx_backup,
            "docx_restore_backup": self.tool_docx_restore_backup,
            "docx_rollback": self.tool_docx_rollback,
            "docx_apply_patchset": self.tool_docx_apply_patchset,
            "docx_validate": self.tool_docx_validate,
            "docx_compare_structure": self.tool_docx_compare_structure,
            "docx_text_diff": self.tool_docx_text_diff,
            # Live Word / Office.js session tools
            "word_session_list": self.tool_word_session_list,
            "word_session_snapshot": self.tool_word_session_snapshot,
            "word_session_refresh": self.tool_word_session_refresh,
            "word_session_read_content_control": self.tool_word_session_read_content_control,
            "word_session_preview_patchset": self.tool_word_session_preview_patchset,
            "word_session_apply_patchset": self.tool_word_session_apply_patchset,
            "word_session_wrap_selection": self.tool_word_session_wrap_selection,
            "word_session_rollback": self.tool_word_session_rollback,
            "word_session_command_status": self.tool_word_session_command_status,
            # Optional OfficeCLI read-only / low-risk auxiliary evidence.
            "officecli_view_html": self.tool_officecli_view_html,
            "officecli_view_screenshot": self.tool_officecli_view_screenshot,
            "officecli_view_issues": self.tool_officecli_view_issues,
            "officecli_query": self.tool_officecli_query,
            "officecli_validate": self.tool_officecli_validate,
        }

    def _resolve_path(self, p: str) -> str:
        path = Path(p)
        if not path.is_absolute():
            path = self.root / path
        path = path.resolve()
        for allowed_root in self.allowed_roots:
            try:
                path.relative_to(allowed_root)
                return str(path)
            except ValueError:
                continue
        allowed = ", ".join(str(x) for x in self.allowed_roots)
        raise PermissionError(f"Path is outside allowed roots: {path}. Allowed roots: {allowed}")

    def _maybe_resolve(self, p: str | None) -> str | None:
        return self._resolve_path(p) if p else None

    def _ensure_write(self) -> None:
        if not self.allow_write:
            raise PermissionError("Server is running in read-only mode")

    def offline_engine_status(self) -> JSON:
        requested = os.environ.get("WORD_AI_ENGINE", "auto")
        try:
            selected, detail = select_engine(requested)
        except Exception as exc:
            selected = "unavailable"
            detail = {"error": str(exc)}
        return {
            "requested": requested,
            "selected": selected,
            "dotnet": dotnet_status(),
            "detail": detail,
        }

    def _offline_engine(self, args: JSON) -> tuple[str, JSON | None]:
        return select_engine(args.get("engine"))

    @staticmethod
    def _python_result(result: Any, detail: JSON | None = None) -> Any:
        return mark_python_result(result, detail)

    @staticmethod
    def _with_precondition_verdict(docx_path: str, args: JSON, result: JSON) -> JSON:
        """Add the precondition verdict the .NET assessment does not compute.

        The .NET backend resolves every target but never compares `expected_old_*`
        against it: it answers `ok: true, risks: []` for an `expected_old_sha256`
        of all zeroes, and for a table's `xml_sha256` - which `docx_list_tables`
        offers right beside the `text_sha256` a precondition actually needs -
        while its own write path goes on to refuse that very edit. Assessment
        blessing an edit the writer rejects is the worst answer this tool can
        give, and the .NET engine ships prebuilt, so the verdict from
        `assess_patchset` (the same call the write path makes) is merged in here.
        """
        try:
            reference = assess_patchset(docx_path, args["patchset"])
        except Exception:
            return result
        mismatches = [risk for risk in reference.get("risks", []) if risk.get("code") == "precondition_mismatch"]
        if not mismatches:
            return result
        risks = list(result.get("risks") or []) + mismatches
        return {**result, "risks": risks, "ok": False, "precondition_verdict": "mismatch"}


    def list_tools(self) -> list[JSON]:
        strp = {"type": "string"}
        intp = {"type": "integer", "minimum": 1}
        boolp = {"type": "boolean"}
        enginep = {"type": "string", "enum": ["auto", "dotnet", "python"], "default": "auto"}
        patchset = {"$ref": "https://word-ai.local/schemas/patchset.schema.json"}
        wait_props = {
            "wait": {"type": "boolean", "default": True},
            "timeout_seconds": {"type": "number", "default": 20},
        }
        validate_props = {
            "source_docx": strp,
            "target_docx": strp,
            "strict": {"type": "boolean", "default": True},
            "touched_content_control_tags": {"type": "array", "items": strp},
            "touched_para_ids": {"type": "array", "items": strp},
            "touched_paragraph_indices": {"type": "array", "items": {"type": "integer"}},
            "touched_table_indices": {"type": "array", "items": {"type": "integer"}},
            "touched_table_cells": {"type": "array", "items": {"oneOf": [{"type": "string"}, {"type": "array", "items": {"type": "integer"}, "minItems": 3, "maxItems": 3}, {"type": "object", "properties": {"table_index": {"type": "integer"}, "row": {"type": "integer"}, "col": {"type": "integer"}}, "required": ["table_index", "row", "col"], "additionalProperties": False}]}},
            "allow_paragraph_count_change": {"type": "boolean", "default": False},
            "allow_table_dimension_change": {"type": "boolean", "default": False},
            "allowed_added_content_control_tags": {"type": "array", "items": strp},
            "use_audit": {"type": "boolean", "default": True},
            "allow_skipped_isolation_checks": {"type": "boolean", "default": False},
        }
        specs = [
            ("docx_health_check", "Read-only. Conservative stability report: duplicate tags, complex content controls, fields, tracked changes, comments, table previews and recommended safety policy.", {"docx_path": strp, "max_items": {"type": "integer", "default": 50}}, ["docx_path"]),
            ("docx_inspect", "Read-only. Inspect counts, anchors and optional visible text.", {"docx_path": strp, "include_text": {"type": "boolean", "default": False}}, ["docx_path"]),
            ("docx_package_manifest", "Read-only. DOCX ZIP manifest with part sizes and optional SHA-256 hashes.", {"docx_path": strp, "include_hashes": {"type": "boolean", "default": True}}, ["docx_path"]),
            ("docx_list_parts", "Read-only. List all package parts with sizes and hashes.", {"docx_path": strp}, ["docx_path"]),
            ("docx_part_hashes", "Read-only. Compact DOCX part-hash map for before/after invariance checks.", {"docx_path": strp}, ["docx_path"]),
            ("docx_structural_fingerprint", "Read-only. Strong structure fingerprint: package parts, content-control hashes, paragraph hashes, table hashes, outline and health report.", {"docx_path": strp}, ["docx_path"]),
            ("docx_map", "Read-only. Codex-friendly map: health, headings, anchors, paragraph chunks, table summaries and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 500}, "include_text": {"type": "boolean", "default": False}}, ["docx_path"]),
            ("docx_get_outline", "Read-only. Heading outline with section paragraph ranges for large-document chunking.", {"docx_path": strp}, ["docx_path"]),
            ("docx_outline", "Read-only alias for docx_get_outline; kept for Codex prompt ergonomics.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_headings", "Read-only. Heading-only view with previews and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_read_heading_section", "Read-only. Read one heading section by heading anchor_id or exact heading text.", {"docx_path": strp, "heading_anchor_id": strp, "heading_text": strp, "max_chars": {"type": "integer", "default": 20000}}, ["docx_path"]),
            ("docx_list_anchors", "Read-only. List stable editing anchors: content controls, headings, bookmarks and paraId paragraphs.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_read_anchor", "Read-only. Read text behind an anchor_id from docx_list_anchors.", {"docx_path": strp, "anchor_id": strp, "max_chars": {"type": "integer", "default": 20000}}, ["docx_path", "anchor_id"]),
            ("docx_search_text", "Read-only. Search visible paragraph text; returns paraId, paragraph index, scope and hashes.", {"docx_path": strp, "query": strp, "regex": {"type": "boolean", "default": False}, "case_sensitive": {"type": "boolean", "default": False}, "max_results": {"type": "integer", "default": 50}, "context_chars": {"type": "integer", "default": 120}}, ["docx_path", "query"]),
            ("docx_list_paragraphs", "Read-only. Paragraph inventory with style, paraId, heading level, scope, preview, hash and complexity.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}, "include_empty": {"type": "boolean", "default": False}}, ["docx_path"]),
            ("docx_read_paragraph", "Read-only. Read one paragraph by paragraph_index or paraId and return hash/complexity.", {"docx_path": strp, "paragraph_index": intp, "paraId": strp}, ["docx_path"]),
            ("docx_list_content_controls", "Read-only. List content controls with tag, alias, id, preview, text hash and complexity.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_read_content_control", "Read-only. Read content-control text by w:tag. Preferred primitive for safe edits.", {"docx_path": strp, "tag": strp}, ["docx_path", "tag"]),
            ("docx_extract_plain_text", "Read-only. Extract visible body text for coarse review; never use to regenerate DOCX.", {"docx_path": strp}, ["docx_path"]),
            ("docx_export_plain_text", "Write-sidecar only. Export visible body text to .txt; does not modify DOCX.", {"docx_path": strp, "out_path": strp}, ["docx_path"]),
            ("docx_list_tables", "Read-only. List table dimensions, scope tags, complexity and preview rows.", {"docx_path": strp, "max_cell_chars": {"type": "integer", "default": 120}}, ["docx_path"]),
            ("docx_read_table", "Read-only. Read a table matrix by table_index and optional content-control scope_tag.", {"docx_path": strp, "table_index": intp, "scope_tag": strp, "max_chars_per_cell": {"type": "integer", "default": 2000}}, ["docx_path", "table_index"]),
            ("docx_read_table_cell", "Read-only. Read a single table cell and return hash for write preconditions.", {"docx_path": strp, "table_index": intp, "row": intp, "col": intp, "scope_tag": strp, "max_chars": {"type": "integer", "default": 20000}}, ["docx_path", "table_index", "row", "col"]),
            ("docx_export_table_csv", "Write-sidecar only. Export table to CSV; does not modify DOCX. Uses default sidecar naming when out_path is omitted.", {"docx_path": strp, "table_index": intp, "out_path": strp, "scope_tag": strp}, ["docx_path", "table_index"]),
            ("docx_table_to_csv", "Write-sidecar only. Export table to CSV using default naming when out_path is omitted.", {"docx_path": strp, "table_index": intp, "out_path": strp, "scope_tag": strp}, ["docx_path", "table_index"]),
            ("docx_list_styles", "Read-only. List Word styles from styles.xml for style-aware generation.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_numbering", "Read-only. Inspect numbering.xml numId/abstractNum mappings.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_sections", "Read-only. Inspect section properties, margins, page size and header/footer refs.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_fields", "Read-only. List field instructions such as TOC, REF, PAGEREF and SEQ.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_list_images", "Read-only. List drawing/pict objects with relationship IDs and alt metadata.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_hyperlinks", "Read-only. List hyperlinks with relationship targets, preview and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_list_comments", "Read-only. Extract comments.xml previews, author/date/id and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 500}}, ["docx_path"]),
            ("docx_extract_comments", "Read-only. Extract full comment metadata and text.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_bookmarks", "Read-only. List bookmarks and paths for cross-reference preservation.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_headers_footers", "Read-only. Inspect header/footer parts and visible text without modifying them.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 500}}, ["docx_path"]),
            ("docx_list_notes", "Read-only. Inspect footnotes/endnotes when present.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 500}}, ["docx_path"]),
            ("docx_list_revisions", "Read-only. List tracked-change nodes with previews and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_assess_patchset", "Read-only. Resolve PatchSet targets and report risks, touched objects and precondition gaps. Defaults to the .NET Open XML backend when available.", {"docx_path": strp, "patchset": patchset, "engine": enginep}, ["docx_path", "patchset"]),
            ("docx_plan_patchset", "Read-only. Alias-grade PatchSet planning assessment. Defaults to the .NET Open XML backend when available.", {"docx_path": strp, "patchset": patchset, "engine": enginep}, ["docx_path", "patchset"]),
            ("docx_preflight_patchset", "Write-temp. Dry-run PatchSet to a temporary copy, validate, and remove output unless keep_output=true. Defaults to the .NET Open XML backend when available.", {"docx_path": strp, "patchset": patchset, "keep_output": {"type": "boolean", "default": False}, "engine": enginep}, ["docx_path", "patchset"]),
            ("docx_dry_run_patchset", "Write-temp. Same as preflight; useful for explicit Codex dry-run stage. Defaults to the .NET Open XML backend when available.", {"docx_path": strp, "patchset": patchset, "keep_output": {"type": "boolean", "default": False}, "engine": enginep}, ["docx_path", "patchset"]),
            ("docx_write_index", "Write-sidecar only. Write .wordai index JSON; does not modify DOCX.", {"docx_path": strp, "out_path": strp}, ["docx_path"]),
            ("docx_backup", "Write-sidecar only. Create timestamped backup; does not modify original DOCX.", {"docx_path": strp, "backup_dir": strp}, ["docx_path"]),
            ("docx_restore_backup", "Write. Restore a backup to target_path.", {"backup_path": strp, "target_path": strp}, ["backup_path", "target_path"]),
            ("docx_rollback", "Write. Restore from backup and optionally back up replaced file first.", {"backup_path": strp, "restore_path": strp, "make_backup_of_current": {"type": "boolean", "default": True}}, ["backup_path", "restore_path"]),
            ("docx_apply_patchset", "Write. Apply constrained PatchSet to a new DOCX plus audit JSON; validation gates final commit. Defaults to the .NET Open XML backend when available.", {"docx_path": strp, "output_path": strp, "patchset": patchset, "engine": enginep}, ["docx_path", "patchset"]),
            ("docx_validate", "Read-only. Validate structural invariants. Touched scopes are loaded from the sibling '<target>.audit.json' written by docx_apply_patchset; do not hand-write touched_* unless you have no audit, and then always pass touched_para_ids and touched_paragraph_indices together. Defaults to the .NET Open XML backend when available.", {**validate_props, "engine": enginep}, ["source_docx", "target_docx"]),
            ("docx_compare_structure", "Read-only. Same validation report phrased as structural comparison. Touched scopes are loaded from the sibling '<target>.audit.json' written by docx_apply_patchset. Defaults to the .NET Open XML backend when available.", {**validate_props, "engine": enginep}, ["source_docx", "target_docx"]),
            ("docx_text_diff", "Read-only. Unified visible-text diff for human review after validation.", {"source_docx": strp, "target_docx": strp, "context": {"type": "integer", "default": 2}}, ["source_docx", "target_docx"]),
            ("word_session_list", "Read-only. List active Office.js taskpane sessions registered by an open Word document.", {"include_inactive": {"type": "boolean", "default": False}}, []),
            ("word_session_snapshot", "Read-only. Return the latest content-control snapshot for an open Word session. If session_id is omitted, uses the most recently active session.", {"session_id": strp}, []),
            ("word_session_refresh", "Read-only. Ask the Office.js taskpane to refresh its open-document content-control snapshot.", {"session_id": strp, **wait_props}, []),
            ("word_session_read_content_control", "Read-only. Ask the Office.js taskpane to read content-control text from the currently open Word document by tag.", {"session_id": strp, "tag": strp, **wait_props}, ["tag"]),
            ("word_session_preview_patchset", "Read-only. Ask the Office.js taskpane to validate and preview a PatchSet against the currently open Word document without modifying it.", {"session_id": strp, "patchset": patchset, **wait_props}, ["patchset"]),
            ("word_session_apply_patchset", "Write. Apply a supported content-control PatchSet to the currently open Word document through Office.js, with hash preconditions, audit, and rollback PatchSet.", {"session_id": strp, "patchset": patchset, **wait_props}, ["patchset"]),
            ("word_session_wrap_selection", "Write. Ask the Office.js taskpane to wrap the current Word selection in a content control with a stable tag/title.", {"session_id": strp, "tag": strp, "title": strp, **wait_props}, ["tag"]),
            ("word_session_rollback", "Write. Roll back a previous word_session_apply_patchset command by applying its generated rollback PatchSet to the open Word document.", {"session_id": strp, "command_id": strp, "rollback_patchset": patchset, **wait_props}, []),
            ("word_session_command_status", "Read-only. Return the current status/result/error for a queued Office.js session command.", {"command_id": strp}, ["command_id"]),
            ("officecli_view_html", "Optional OfficeCLI auxiliary read. Render a DOCX to a bounded HTML snapshot for visual evidence. Does not modify the DOCX.", {"docx_path": strp, "page": {"type": "integer", "minimum": 1}, "max_output_chars": {"type": "integer", "default": 50000}}, ["docx_path"]),
            ("officecli_view_screenshot", "Optional OfficeCLI auxiliary sidecar export. Render a DOCX screenshot/PNG to output_path. Does not modify the DOCX.", {"docx_path": strp, "output_path": strp, "page": {"type": "integer", "minimum": 1}, "screenshot_width": {"type": "integer", "minimum": 320}, "screenshot_height": {"type": "integer", "minimum": 320}, "max_output_chars": {"type": "integer", "default": 20000}}, ["docx_path", "output_path"]),
            ("officecli_view_issues", "Optional OfficeCLI auxiliary read. Run view issues with JSON output for formatting/content/structure evidence.", {"docx_path": strp, "issue_type": {"type": "string", "enum": ["format", "content", "structure"]}, "limit": {"type": "integer", "minimum": 1}, "max_output_chars": {"type": "integer", "default": 50000}}, ["docx_path"]),
            ("officecli_query", "Optional OfficeCLI auxiliary read. Run CSS-like query with --json. Use only for inspection, never mutation.", {"docx_path": strp, "selector": strp, "max_output_chars": {"type": "integer", "default": 50000}}, ["docx_path", "selector"]),
            ("officecli_validate", "Optional OfficeCLI auxiliary read. Validate the DOCX with OfficeCLI --json as extra evidence after Word AI validation.", {"docx_path": strp, "max_output_chars": {"type": "integer", "default": 50000}}, ["docx_path"]),
        ]
        return [{"name": n, "description": d, "inputSchema": _schema(p, r)} for n, d, p, r in specs]

    # Package / health / map.
    def tool_docx_health_check(self, args: JSON) -> Any:
        return health_check(self._resolve_path(args["docx_path"]), int(args.get("max_items", 50)))

    def tool_docx_inspect(self, args: JSON) -> Any:
        return inspect_docx(self._resolve_path(args["docx_path"]), bool(args.get("include_text", False)))

    def tool_docx_package_manifest(self, args: JSON) -> Any:
        return package_manifest(self._resolve_path(args["docx_path"]), bool(args.get("include_hashes", True)))

    def tool_docx_list_parts(self, args: JSON) -> Any:
        return list_docx_parts(self._resolve_path(args["docx_path"]))

    def tool_docx_part_hashes(self, args: JSON) -> Any:
        return part_hashes(self._resolve_path(args["docx_path"]))

    def tool_docx_structural_fingerprint(self, args: JSON) -> Any:
        return structural_fingerprint(self._resolve_path(args["docx_path"]))

    def tool_docx_map(self, args: JSON) -> Any:
        return build_document_map(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 500)), bool(args.get("include_text", False)))

    # Navigation / retrieval.
    def tool_docx_get_outline(self, args: JSON) -> Any:
        return get_outline(self._resolve_path(args["docx_path"]))

    def tool_docx_list_headings(self, args: JSON) -> Any:
        return list_headings(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    def tool_docx_read_heading_section(self, args: JSON) -> Any:
        return read_heading_section(self._resolve_path(args["docx_path"]), args.get("heading_anchor_id"), args.get("heading_text"), int(args.get("max_chars", 20000)))

    def tool_docx_list_anchors(self, args: JSON) -> Any:
        anchors = list_anchors(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))
        return {"anchors": [a.to_dict() for a in anchors], "anchor_count": len(anchors)}

    def tool_docx_read_anchor(self, args: JSON) -> Any:
        return read_anchor(self._resolve_path(args["docx_path"]), args["anchor_id"], int(args.get("max_chars", 20000)))

    def tool_docx_search_text(self, args: JSON) -> Any:
        return search_text(self._resolve_path(args["docx_path"]), args["query"], bool(args.get("regex", False)), bool(args.get("case_sensitive", False)), int(args.get("max_results", 50)), int(args.get("context_chars", 120)))

    def tool_docx_list_paragraphs(self, args: JSON) -> Any:
        return list_paragraphs(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)), bool(args.get("include_empty", False)))

    def tool_docx_read_paragraph(self, args: JSON) -> Any:
        return read_paragraph(self._resolve_path(args["docx_path"]), args.get("paragraph_index"), args.get("paraId"))

    def tool_docx_list_content_controls(self, args: JSON) -> Any:
        return list_content_controls(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    def tool_docx_read_content_control(self, args: JSON) -> Any:
        return get_content_control_text(self._resolve_path(args["docx_path"]), args["tag"])

    def tool_docx_extract_plain_text(self, args: JSON) -> Any:
        return extract_plain_text(self._resolve_path(args["docx_path"]))

    def tool_docx_export_plain_text(self, args: JSON) -> Any:
        self._ensure_write()
        return export_plain_text(self._resolve_path(args["docx_path"]), self._maybe_resolve(args.get("out_path")))

    # Tables / styles / complex Word objects.
    def tool_docx_list_tables(self, args: JSON) -> Any:
        return list_tables(self._resolve_path(args["docx_path"]), int(args.get("max_cell_chars", 120)))

    def tool_docx_read_table(self, args: JSON) -> Any:
        data = read_table(self._resolve_path(args["docx_path"]), int(args["table_index"]), args.get("scope_tag"), int(args.get("max_chars_per_cell", 2000)))
        if isinstance(data, dict) and "rows" in data:
            data.setdefault("row_count", len(data["rows"]))
            data.setdefault("column_counts", [len(r) for r in data["rows"]])
        return data

    def tool_docx_read_table_cell(self, args: JSON) -> Any:
        return read_table_cell(self._resolve_path(args["docx_path"]), int(args["table_index"]), int(args["row"]), int(args["col"]), args.get("scope_tag"), int(args.get("max_chars", 20000)))

    def tool_docx_export_table_csv(self, args: JSON) -> Any:
        self._ensure_write()
        return table_to_csv(self._resolve_path(args["docx_path"]), int(args["table_index"]), self._maybe_resolve(args.get("out_path")), args.get("scope_tag"))

    def tool_docx_table_to_csv(self, args: JSON) -> Any:
        self._ensure_write()
        return table_to_csv(self._resolve_path(args["docx_path"]), int(args["table_index"]), self._maybe_resolve(args.get("out_path")), args.get("scope_tag"))

    def tool_docx_list_styles(self, args: JSON) -> Any:
        return list_styles(self._resolve_path(args["docx_path"]))

    def tool_docx_list_numbering(self, args: JSON) -> Any:
        return list_numbering(self._resolve_path(args["docx_path"]))

    def tool_docx_list_sections(self, args: JSON) -> Any:
        return list_sections(self._resolve_path(args["docx_path"]))

    def tool_docx_list_fields(self, args: JSON) -> Any:
        return list_fields(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    def tool_docx_list_images(self, args: JSON) -> Any:
        return list_images(self._resolve_path(args["docx_path"]))

    def tool_docx_list_hyperlinks(self, args: JSON) -> Any:
        return list_hyperlinks(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    def tool_docx_list_comments(self, args: JSON) -> Any:
        return list_comments(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 500)))

    def tool_docx_extract_comments(self, args: JSON) -> Any:
        return extract_comments(self._resolve_path(args["docx_path"]))

    def tool_docx_list_bookmarks(self, args: JSON) -> Any:
        return list_bookmarks(self._resolve_path(args["docx_path"]))

    def tool_docx_list_headers_footers(self, args: JSON) -> Any:
        return list_headers_footers(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 500)))

    def tool_docx_list_notes(self, args: JSON) -> Any:
        return list_notes(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 500)))

    def tool_docx_list_revisions(self, args: JSON) -> Any:
        return list_revisions(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    # Patch / write / validation lifecycle.
    def tool_docx_assess_patchset(self, args: JSON) -> Any:
        docx_path = self._resolve_path(args["docx_path"])
        engine, detail = self._offline_engine(args)
        if engine == "dotnet":
            return self._with_precondition_verdict(docx_path, args, dotnet_assess_patchset(docx_path, args["patchset"]))
        return self._python_result(assess_patchset(docx_path, args["patchset"]), detail)

    def tool_docx_plan_patchset(self, args: JSON) -> Any:
        docx_path = self._resolve_path(args["docx_path"])
        engine, detail = self._offline_engine(args)
        if engine == "dotnet":
            return self._with_precondition_verdict(docx_path, args, dotnet_assess_patchset(docx_path, args["patchset"]))
        return self._python_result(plan_patchset(docx_path, args["patchset"]), detail)

    def tool_docx_preflight_patchset(self, args: JSON) -> Any:
        self._ensure_write()
        docx_path = self._resolve_path(args["docx_path"])
        keep_output = bool(args.get("keep_output", False))
        engine, detail = self._offline_engine(args)
        if engine == "dotnet":
            return dotnet_dry_run_patchset(docx_path, args["patchset"], keep_output)
        return self._python_result(dry_run_patchset(docx_path, args["patchset"], keep_output), detail)

    def tool_docx_dry_run_patchset(self, args: JSON) -> Any:
        self._ensure_write()
        docx_path = self._resolve_path(args["docx_path"])
        keep_output = bool(args.get("keep_output", False))
        engine, detail = self._offline_engine(args)
        if engine == "dotnet":
            return dotnet_dry_run_patchset(docx_path, args["patchset"], keep_output)
        return self._python_result(dry_run_patchset(docx_path, args["patchset"], keep_output), detail)

    def tool_docx_write_index(self, args: JSON) -> Any:
        self._ensure_write()
        return {"index_path": write_sidecar_index(self._resolve_path(args["docx_path"]), self._maybe_resolve(args.get("out_path")))}

    def tool_docx_backup(self, args: JSON) -> Any:
        self._ensure_write()
        return {"backup_path": backup_docx(self._resolve_path(args["docx_path"]), self._maybe_resolve(args.get("backup_dir")))}

    def tool_docx_restore_backup(self, args: JSON) -> Any:
        self._ensure_write()
        return restore_backup(self._resolve_path(args["backup_path"]), self._resolve_path(args["target_path"]))

    def tool_docx_rollback(self, args: JSON) -> Any:
        self._ensure_write()
        return rollback_docx(self._resolve_path(args["backup_path"]), self._resolve_path(args["restore_path"]), bool(args.get("make_backup_of_current", True)))

    def tool_docx_apply_patchset(self, args: JSON) -> Any:
        self._ensure_write()
        docx_path = self._resolve_path(args["docx_path"])
        output_path = self._maybe_resolve(args.get("output_path"))
        engine, detail = self._offline_engine(args)
        if engine == "dotnet":
            return dotnet_apply_patchset(docx_path, args["patchset"], output_path)
        return self._python_result(apply_patchset(docx_path, args["patchset"], output_path), detail)

    _ISOLATION_REQUIRED = (
        "docx_validate cannot prove that an edit stayed isolated without knowing what the edit "
        "was supposed to touch: with no touched scopes supplied the deep protected_* checks are "
        "skipped and a passing report would prove nothing. Do one of: "
        "(a) validate a target produced by docx_apply_patchset and keep its sibling "
        "'<target>.audit.json' in place - the touched scopes are then loaded automatically; "
        "(b) pass touched_* explicitly, always supplying touched_para_ids AND "
        "touched_paragraph_indices together, because the paraId and body-block checks read "
        "different axes; or (c) pass allow_skipped_isolation_checks=true to run structural "
        "package checks only and accept an inconclusive result."
    )

    def _validation_kwargs(self, args: JSON, target_docx: str | None = None) -> tuple[dict[str, Any], list[dict[str, str]]]:
        """Build validation kwargs and collect diagnostics about the isolation claim.

        The touched scopes describe what the caller believes it edited; the validator
        compares that claim against the two documents. Two things used to make a clean
        edit look broken:

        * An explicit touched_* argument suppressed the audit lookup entirely, so a caller
          who passed a hand-written half of the claim silently replaced a correct one.
          Explicit values are therefore merged into the audit-derived set, never used to
          replace it.
        * The isolation checks use two axes - paragraphs are matched by w14:paraId, body
          blocks by paragraph index - and supplying only one of them leaves the other
          check comparing against an empty set. Supplying only one is now reported.

        The audit is also the only way the deep checks can run at all, so validating with
        neither an audit nor an explicit claim is refused instead of silently returning a
        green report that checked nothing.
        """

        def ints(values: Any) -> list[int]:
            out: list[int] = []
            for value in values or []:
                try:
                    out.append(int(value))
                except (TypeError, ValueError):
                    continue
            return out

        explicit = {
            "content_control_tags": [str(x) for x in (args.get("touched_content_control_tags") or [])],
            "para_ids": [str(x) for x in (args.get("touched_para_ids") or [])],
            "paragraph_indices": ints(args.get("touched_paragraph_indices")),
            "table_indices": ints(args.get("touched_table_indices")),
            "table_cells": [str(x) for x in (args.get("touched_table_cells") or [])],
        }
        explicit_supplied = sorted(k for k, v in explicit.items() if v)
        allowed_added_tags = [str(x) for x in (args.get("allowed_added_content_control_tags") or [])]
        notes: list[dict[str, str]] = []

        audit_touched: dict[str, list[Any]] | None = None
        audit_path = Path(target_docx).with_suffix(".audit.json") if target_docx else None
        if audit_path and args.get("use_audit", True) and audit_path.exists():
            try:
                audit = json.loads(audit_path.read_text(encoding="utf-8"))
                touched = ((audit.get("safety_assessment") or {}).get("touched") or {})
                audit_touched = {
                    "content_control_tags": [str(x) for x in (touched.get("content_control_tags") or [])],
                    "para_ids": [str(x) for x in (touched.get("paraIds") or touched.get("para_ids") or [])],
                    "paragraph_indices": ints(touched.get("paragraph_indices")),
                    "table_indices": ints(touched.get("table_indices")),
                    "table_cells": [str(x) for x in (touched.get("table_cells") or [])],
                }
                for op in audit.get("applied") or []:
                    for key in ("paragraph_index", "anchor_paragraph_index", "target_paragraph_index"):
                        if op.get(key) is not None:
                            audit_touched["paragraph_indices"].append(int(op[key]))
                    if op.get("table_index") is not None and op.get("row") is not None and op.get("col") is not None:
                        audit_touched["table_cells"].append(f"{op['table_index']}:{op['row']}:{op['col']}")
                    if op.get("op") == "wrap_paragraph_with_content_control" and op.get("tag"):
                        allowed_added_tags.append(op["tag"])
            except Exception:
                audit_touched = None

        merged = {
            key: sorted(set(explicit[key]) | set(audit_touched[key] if audit_touched else []))
            for key in explicit
        }

        if audit_touched is not None and explicit_supplied:
            notes.append({
                "severity": "warning",
                "code": "touched_merged_with_audit",
                "message": (
                    f"Explicit touched_* ({', '.join(explicit_supplied)}) was merged with the audit "
                    f"at {audit_path.name} instead of replacing it, because the isolation checks need "
                    "both the paraId and the paragraph-index axes."
                ),
            })
        elif explicit_supplied and audit_touched is None:
            only_para_ids = bool(explicit["para_ids"]) and not explicit["paragraph_indices"]
            only_indices = bool(explicit["paragraph_indices"]) and not explicit["para_ids"]
            if only_para_ids or only_indices:
                missing = "touched_paragraph_indices" if only_para_ids else "touched_para_ids"
                notes.append({
                    "severity": "warning",
                    "code": "touched_sets_incomplete",
                    "message": (
                        f"Only one axis of the isolation claim was supplied and no audit was found; "
                        f"{missing} is empty. Paragraphs are matched by paraId while body blocks are "
                        "matched by paragraph index, so a single-axis claim makes the other check "
                        "report protected_* errors on a correct edit."
                    ),
                })

        if not any(merged.values()):
            if not bool(args.get("allow_skipped_isolation_checks", False)):
                raise ValueError(self._ISOLATION_REQUIRED)
            notes.append({
                "severity": "warning",
                "code": "isolation_checks_skipped",
                "message": "Isolation checks were skipped on request: no touched scopes supplied, so only structural package checks ran and the result is inconclusive.",
            })

        kwargs = {
            "touched_content_control_tags": merged["content_control_tags"],
            "touched_para_ids": merged["para_ids"],
            "touched_paragraph_indices": merged["paragraph_indices"],
            "touched_table_indices": merged["table_indices"],
            "touched_table_cells": merged["table_cells"],
            "allow_paragraph_count_change": bool(args.get("allow_paragraph_count_change", False)),
            "allow_table_dimension_change": bool(args.get("allow_table_dimension_change", False)),
            "allowed_added_content_control_tags": sorted(set(allowed_added_tags)),
        }
        return kwargs, notes

    @staticmethod
    def _annotate_isolation(result: Any, kwargs: JSON, notes: list[dict[str, str]]) -> Any:
        """Publish how the isolation claim was obtained and attach its diagnostics."""
        if not isinstance(result, dict):
            return result
        result["isolation_checks"] = "performed" if any(
            kwargs.get(key) for key in (
                "touched_content_control_tags",
                "touched_para_ids",
                "touched_paragraph_indices",
                "touched_table_indices",
                "touched_table_cells",
            )
        ) else "skipped"
        result["touched_scopes"] = {
            "content_control_tags": kwargs["touched_content_control_tags"],
            "para_ids": kwargs["touched_para_ids"],
            "paragraph_indices": kwargs["touched_paragraph_indices"],
            "table_indices": kwargs["touched_table_indices"],
            "table_cells": kwargs["touched_table_cells"],
        }
        if notes:
            issues = result.setdefault("issues", [])
            existing = {(i.get("code"), i.get("message")) for i in issues if isinstance(i, dict)}
            for note in notes:
                if (note["code"], note["message"]) not in existing:
                    issues.append(note)
        return result

    def tool_docx_validate(self, args: JSON) -> Any:
        source = self._resolve_path(args["source_docx"])
        target = self._resolve_path(args["target_docx"])
        kwargs, notes = self._validation_kwargs(args, target)
        strict = bool(args.get("strict", True))
        engine, detail = self._offline_engine(args)
        if engine == "dotnet":
            return self._annotate_isolation(dotnet_validate(source, target, strict, kwargs), kwargs, notes)
        return self._annotate_isolation(self._python_result(validate_structure(source, target, strict, **kwargs).to_dict(), detail), kwargs, notes)

    def tool_docx_compare_structure(self, args: JSON) -> Any:
        source = self._resolve_path(args["source_docx"])
        target = self._resolve_path(args["target_docx"])
        kwargs, notes = self._validation_kwargs(args, target)
        strict = bool(args.get("strict", True))
        engine, detail = self._offline_engine(args)
        if engine == "dotnet":
            return self._annotate_isolation(dotnet_validate(source, target, strict, kwargs), kwargs, notes)
        return self._annotate_isolation(self._python_result(validate_structure(source, target, strict, **kwargs).to_dict(), detail), kwargs, notes)

    def tool_docx_text_diff(self, args: JSON) -> Any:
        return diff_text(self._resolve_path(args["source_docx"]), self._resolve_path(args["target_docx"]), int(args.get("context", 2)))

    # Optional OfficeCLI auxiliary backend. These wrappers intentionally expose
    # only read-only / low-risk commands and never accept raw OfficeCLI args.
    def _officecli_path(self) -> str | None:
        configured = os.environ.get("WORD_AI_OFFICECLI")
        if configured:
            return configured
        return shutil.which("officecli")

    def _run_officecli(self, argv: list[str], max_output_chars: int = 50000, timeout_seconds: int = 90) -> Any:
        exe = self._officecli_path()
        if not exe:
            return {
                "ok": False,
                "available": False,
                "error": "officecli not found. Install OfficeCLI only if you want optional render/query/validate evidence.",
                "allowed_commands": [
                    "view html",
                    "view screenshot",
                    "view issues",
                    "query --json",
                    "validate --json",
                ],
            }
        try:
            proc = subprocess.run(
                [exe, *argv],
                cwd=str(self.root),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                errors="replace",
                timeout=timeout_seconds,
            )
        except subprocess.TimeoutExpired as exc:
            return {"ok": False, "available": True, "error": f"officecli timed out after {timeout_seconds}s", "stdout": exc.stdout, "stderr": exc.stderr}

        stdout = proc.stdout or ""
        stderr = proc.stderr or ""
        stdout_truncated = len(stdout) > max_output_chars
        stderr_truncated = len(stderr) > max_output_chars
        stdout_out = stdout[:max_output_chars]
        stderr_out = stderr[:max_output_chars]
        parsed = None
        if stdout_out.strip():
            try:
                parsed = json.loads(stdout_out)
            except json.JSONDecodeError:
                parsed = None
        return {
            "ok": proc.returncode == 0,
            "available": True,
            "command": ["officecli", *argv],
            "returncode": proc.returncode,
            "stdout": stdout_out,
            "stderr": stderr_out,
            "stdout_truncated": stdout_truncated,
            "stderr_truncated": stderr_truncated,
            "json": parsed,
        }

    def tool_officecli_view_html(self, args: JSON) -> Any:
        docx_path = self._resolve_path(args["docx_path"])
        cmd = ["view", docx_path, "html"]
        if args.get("page"):
            cmd.extend(["--page", str(int(args["page"]))])
        return self._run_officecli(cmd, int(args.get("max_output_chars", 50000)))

    def tool_officecli_view_screenshot(self, args: JSON) -> Any:
        self._ensure_write()
        docx_path = self._resolve_path(args["docx_path"])
        output_path = self._resolve_path(args["output_path"])
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        cmd = ["view", docx_path, "screenshot", "-o", output_path]
        if args.get("page"):
            cmd.extend(["--page", str(int(args["page"]))])
        if args.get("screenshot_width"):
            cmd.extend(["--screenshot-width", str(int(args["screenshot_width"]))])
        if args.get("screenshot_height"):
            cmd.extend(["--screenshot-height", str(int(args["screenshot_height"]))])
        result = self._run_officecli(cmd, int(args.get("max_output_chars", 20000)), timeout_seconds=120)
        result["output_path"] = output_path
        result["output_exists"] = Path(output_path).exists()
        return result

    def tool_officecli_view_issues(self, args: JSON) -> Any:
        docx_path = self._resolve_path(args["docx_path"])
        cmd = ["view", docx_path, "issues", "--json"]
        if args.get("issue_type"):
            cmd.extend(["--type", str(args["issue_type"])])
        if args.get("limit"):
            cmd.extend(["--limit", str(int(args["limit"]))])
        return self._run_officecli(cmd, int(args.get("max_output_chars", 50000)))

    def tool_officecli_query(self, args: JSON) -> Any:
        docx_path = self._resolve_path(args["docx_path"])
        return self._run_officecli(["query", docx_path, str(args["selector"]), "--json"], int(args.get("max_output_chars", 50000)))

    def tool_officecli_validate(self, args: JSON) -> Any:
        docx_path = self._resolve_path(args["docx_path"])
        return self._run_officecli(["validate", docx_path, "--json"], int(args.get("max_output_chars", 50000)))

    # Live Word / Office.js session tools.
    def _active_session_id(self, args: JSON) -> str:
        if args.get("session_id"):
            return str(args["session_id"])
        active = list_sessions(self.root, include_inactive=False)
        if not active:
            raise ValueError("No active Word session found. Open the Office.js taskpane, connect the bridge, and keep Word open.")
        active.sort(key=lambda item: str(item.get("updated_at") or ""), reverse=True)
        return str(active[0]["session_id"])

    def _enqueue_word_session_command(self, args: JSON, command_type: str, payload: JSON) -> Any:
        session_id = self._active_session_id(args)
        command = enqueue_command(self.root, session_id, command_type, payload)
        if not bool(args.get("wait", True)):
            return command
        timeout = float(args.get("timeout_seconds", 20))
        return wait_for_command(self.root, command["command_id"], timeout_seconds=timeout)

    def tool_word_session_list(self, args: JSON) -> Any:
        sessions = list_sessions(self.root, bool(args.get("include_inactive", False)))
        return {"sessions": sessions, "count": len(sessions)}

    def tool_word_session_snapshot(self, args: JSON) -> Any:
        return session_summary(self.root, self._active_session_id(args))

    def tool_word_session_refresh(self, args: JSON) -> Any:
        return self._enqueue_word_session_command(args, "list_content_controls", {})

    def tool_word_session_read_content_control(self, args: JSON) -> Any:
        return self._enqueue_word_session_command(args, "read_content_control", {"tag": args["tag"]})

    def tool_word_session_preview_patchset(self, args: JSON) -> Any:
        return self._enqueue_word_session_command(args, "preview_patchset", {"patchset": normalize_patchset(args["patchset"])})

    def tool_word_session_apply_patchset(self, args: JSON) -> Any:
        self._ensure_write()
        return self._enqueue_word_session_command(args, "apply_patchset", {"patchset": normalize_patchset(args["patchset"])})

    def tool_word_session_wrap_selection(self, args: JSON) -> Any:
        self._ensure_write()
        return self._enqueue_word_session_command(args, "wrap_selection", {"tag": args["tag"], "title": args.get("title")})

    def tool_word_session_rollback(self, args: JSON) -> Any:
        self._ensure_write()
        rollback_patchset = args.get("rollback_patchset")
        source_command_id = args.get("command_id")
        if rollback_patchset is None:
            if not source_command_id:
                raise ValueError("word_session_rollback requires command_id or rollback_patchset")
            source = get_command(self.root, str(source_command_id))
            rollback_patchset = ((source.get("result") or {}).get("audit") or {}).get("rollback_patchset") or (source.get("result") or {}).get("rollback_patchset")
        if not rollback_patchset:
            raise ValueError("No rollback_patchset found for the requested command")
        return self._enqueue_word_session_command(
            args,
            "apply_patchset",
            {"patchset": normalize_patchset(rollback_patchset), "rollback_of_command_id": source_command_id},
        )

    def tool_word_session_command_status(self, args: JSON) -> Any:
        return get_command(self.root, args["command_id"])

    def _missing_argument(self, name: str, exc: KeyError) -> str | None:
        """Name the argument a tool was missing, or None if it was not one.

        Tools read their required arguments as `args["docx_path"]`, so a caller
        who guesses the parameter name - `path`, say - is answered with
        `KeyError: 'docx_path'` and a Python traceback, which says nothing about
        what the tool actually accepts. Returns None when the missing key is not
        one of the tool's declared required arguments: a KeyError from inside a
        tool's own data lookup is not the caller's mistake and must keep its
        original traceback.
        """
        missing = exc.args[0] if exc.args else None
        schema = next((tool.get("inputSchema") or {} for tool in self.list_tools() if tool["name"] == name), {})
        required = list(schema.get("required") or [])
        if missing not in required:
            return None
        accepts = ", ".join(required) or ", ".join(sorted(schema.get("properties") or {}))
        return f"{name}: missing required argument {missing!r}; this tool requires: {accepts}"

    def handle(self, request: JSON) -> JSON | None:
        method = request.get("method")
        req_id = request.get("id")
        try:
            if method == "initialize":
                result = {"protocolVersion": request.get("params", {}).get("protocolVersion", "2025-11-25"), "serverInfo": {"name": "word-ai-mcp", "version": __version__}, "capabilities": {"tools": {"listChanged": False}, "resources": {}, "prompts": {}}}
                return {"jsonrpc": "2.0", "id": req_id, "result": result}
            if method == "notifications/initialized":
                return None
            if method == "tools/list":
                return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": self.list_tools()}}
            if method == "tools/call":
                params = request.get("params") or {}
                name = params.get("name")
                arguments = params.get("arguments") or {}
                if name not in self.tools:
                    raise ValueError(f"Unknown tool: {name}")
                try:
                    result = self.tools[name](arguments)
                except KeyError as exc:
                    hint = self._missing_argument(name, exc)
                    if hint is None:
                        raise
                    raise ValueError(hint) from exc
                return {"jsonrpc": "2.0", "id": req_id, "result": {"content": _text_payload(result), "isError": False}}
            if method == "resources/list":
                return {"jsonrpc": "2.0", "id": req_id, "result": {"resources": []}}
            if method == "prompts/list":
                return {"jsonrpc": "2.0", "id": req_id, "result": {"prompts": []}}
            raise ValueError(f"Unsupported method: {method}")
        except Exception as exc:
            return {"jsonrpc": "2.0", "id": req_id, "error": {"code": -32000, "message": str(exc), "data": {"traceback": traceback.format_exc(limit=8)}}}


def configure_stdio_utf8() -> None:
    """Bind the stdio transport to UTF-8 regardless of the process locale.

    MCP stdio is UTF-8, but Python binds sys.stdin/sys.stdout to the locale
    encoding. On Windows that is the ANSI codepage (cp1251 on a Russian host),
    so a UTF-8 request decodes as cp1251 into mojibake: non-ASCII arguments are
    silently corrupted, and read tools return empty results instead of failing.
    """
    for stream in (sys.stdin, sys.stdout):
        if stream is None:
            continue
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8")
        except (ValueError, OSError):
            # Detached or non-reconfigurable stream; leave it untouched.
            pass


def run_stdio(root: str | None, allow_write: bool, allowed_roots: list[str] | None = None) -> None:
    configure_stdio_utf8()
    server = WordAiMcpServer(root=root, allow_write=allow_write, allowed_roots=allowed_roots)
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
            resp = server.handle(req)
        except Exception as exc:
            resp = {"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": str(exc)}}
        if resp is not None:
            sys.stdout.write(json.dumps(resp, ensure_ascii=False) + "\n")
            sys.stdout.flush()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Word AI MCP stdio server")
    parser.add_argument("--root", default=os.getcwd(), help="Primary workspace root. Relative paths resolve here.")
    parser.add_argument("--allow-root", action="append", default=[], help="Additional absolute directory allowed for DOCX input/output. Repeatable. WORD_AI_ALLOWED_ROOTS also works.")
    parser.add_argument("--read-only", action="store_true", help="Disable write tools")
    args = parser.parse_args(argv)
    run_stdio(args.root, allow_write=not args.read_only, allowed_roots=args.allow_root)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
