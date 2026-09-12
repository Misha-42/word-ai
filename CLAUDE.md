# Word AI — MCP сервер для Word DOCX

**Репозиторий:** https://github.com/flyfish-dev/word-ai
**MCP имя:** `io.github.flyfish-dev/word-ai`
**Локальный путь:** `C:\Users\МУС\Documents\GitHub\word-ai`

## Установка и запуск

```powershell
# Активация venv и запуск сервера
& "C:\Users\МУС\Documents\GitHub\word-ai\.venv\Scripts\activate.ps1"
word-ai-mcp --root C:\Users\МУС\Documents\GitHub\word-ai --allow-root C:\Users\МУС\Documents
```

Сервер уже добавлен в Claude Code через `claude mcp add`.

## Быстрые ссылки

- [AGENTS.md](AGENTS.md) — правила для агентов (русская копия ниже)
- [server.json](server.json) — метаданные MCP сервера
- [docs/](docs/) — документация
- [skills/word-ai/SKILL.md](skills/word-ai/SKILL.md) — скилл для агентов

## Принципы работы

- Не перестраивать DOCX целиком
- Не конвертировать DOCX → Markdown/HTML → DOCX
- Не перезаписывать исходный файл без резервной копии
- Писать только через PatchSet с precondition (source_sha256, expected_old_sha256)
- Перед записью: assess → dry-run; после: validate → diff → audit

## Стандартный флоу

1. `docx_health_check`
2. `docx_map` / `docx_list_anchors` / `docx_list_content_controls`
3. Чтение цели (`docx_read_content_control` / `docx_read_anchor`)
4. Сборка PatchSet с `source_sha256` и `expected_old_sha256`
5. `docx_assess_patchset`
6. `docx_dry_run_patchset`
7. `docx_backup`
8. `docx_apply_patchset`
9. `docx_validate` / `docx_compare_structure` / `docx_text_diff`