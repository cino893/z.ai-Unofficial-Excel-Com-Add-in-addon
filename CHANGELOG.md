# Changelog

## v5 / 5.0.0

### COM add-in (.NET / Excel-DNA)

- Product version is **5.0.0** (assembly, About text, add-in log).
- JSON argument helpers (`Int` / `Dbl` / `Bool`) accept numbers, numeric strings, and booleans so tool-calling floats like `1.0` no longer crash.
- `write_cell` writes invariant-culture numbers as numeric Excel values instead of text.
- New tool: `insert_columns` (single COM `Insert` with `xlToRight`).
- 31 Excel tools total, including `move_pivot_table` and `modify_pivot_table` (not `move_table`).
- Tool loop remains 45 rounds with continue / cancel / loop detection.

### Legacy VBA (feature parity)

- `legacy/` is a maintained VBA build with the same 31 Excel tool names, JSON fields, and conversation capabilities as COM.
- New modules: `modExcelSkillsOps.bas`, `modExcelSkillsPivot.bas`; i18n JSON loaded from `legacy/i18n/` (same 8 languages as COM).
- Conversation: 45 tool rounds, Continue after max rounds or cancel, Stop to cancel, history trim, exact + name-level loop detection, Excel screen-updating suppressed during the tool loop.
- Default model is `glm-4.7-flash`; model picker lists the COM catalog; API keys are sanitized to printable ASCII; HTTP 401/429 and Z.AI codes 1113/1261/1301/1302 map to i18n errors.
- Async XMLHTTP with `DoEvents` so Stop can abort in-flight requests (sync fallback retained).
- Build with `cscript legacy\build.vbs` (copies `i18n/` next to the `.xlam`).

### CI and release

- GitHub Actions `ci.yml`: offline Python parity tests on Ubuntu; `dotnet publish` of the packed `.xll` on Windows.
- GitHub Actions `release.yml`: tag `v*` or `workflow_dispatch` creates a GitHub Release (`Com Release V5` for tag `v5`) and attaches packed XLLs. Uses `GITHUB_TOKEN` only.

### Security

- Removed a hardcoded Z.AI API key from `test_api.py`. Live tests now require the `ZAI_API_KEY` environment variable and skip (exit 0) when it is unset.
- Anyone who forked or cloned a copy that contained that key should **rotate their Z.AI API keys**. Do not commit keys; `.env` is gitignored.

### Tests and docs

- Offline tests: `tests/test_tool_parity.py`, `tests/test_i18n.py`, `tests/test_vwb.py`.
- README / CONTRIBUTING / SECURITY updated for v5, 31 tools, 45 rounds, and maintained VBA parity.
