# Z.AI Excel Add-in

[🇵🇱 Czytaj po polsku](README.pl.md)

Add-in for Microsoft Excel that lets you talk to **z.ai** (Zhipu AI) directly from a worksheet.

**Version 2.0** — rewritten as a .NET COM Add-in with a modern WPF UI (built with Excel-DNA, shipped as `.xll`).

![Showreel](show-reel.gif)

## Project versions

- **v2.0 (.NET COM Add-in)** — main, actively developed project in `src/ZaiExcelAddin` (solution `dodatek-z-ai-opus.sln`).
- **v1.0 (VBA .xlam)** — legacy project in `legacy`; rebuild with `cscript build.vbs`.

## Download

- [Latest Excel-DNA (.xll) package for the COM Add-in](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/releases/latest/download/ZaiExcelAddin-AddIn64-packed.xll)
- [All releases](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/releases)

## ✨ Highlights in v2.0

- **Custom Task Pane** — chat lives on the right side of Excel
- **Modern WPF UI** — chat bubbles, gradients, animated typing dots, logo
- **8 languages** — PL, EN, DE, FR, ES, UK, ZH, JA (auto-detected from Windows)
- **15 AI tools** — incl. `list_charts` and `delete_chart` (chart loop bug fixed)
- **Loop detection** — stops repeating the same tool forever
- **Ribbon tab** — dedicated Z.AI buttons

## Capabilities

| Tool | Description |
|------|-------------|
| `read_cell` / `write_cell` | Read/write a single cell |
| `read_range` / `write_range` | Read/write a 2D range |
| `get_sheet_info` | Sheet dimensions and headers |
| `get_workbook_info` | Workbook sheets and path |
| `format_range` | Fonts, colors, borders, alignment, merge |
| `insert_formula` | Insert Excel formulas |
| `sort_range` | Sort data |
| `add_sheet` | Add sheet |
| `delete_rows` / `insert_rows` | Delete/insert rows |
| `create_chart` | Create charts (column, bar, line, pie, scatter, area) |
| `delete_chart` | Delete a chart |
| `list_charts` | List charts on a sheet |

## Requirements

- Microsoft Excel 2016+ (Windows, 64-bit recommended)
- .NET 8.0 Runtime ([download](https://dotnet.microsoft.com/download/dotnet/8.0))
- API key from [z.ai](https://open.z.ai/) (free registration)

## Build (COM Add-in v2.0)

Requires .NET SDK 8.0+:

```powershell
cd src\ZaiExcelAddin
dotnet build -c Release
```

Output: `src\ZaiExcelAddin\bin\Release\net8.0-windows\publish\ZaiExcelAddin-AddIn64-packed.xll`

## Install (COM Add-in v2.0)

1. Open Excel
2. **File** → **Options** → **Add-ins**
3. At bottom: **Manage** → **Excel Add-ins** → **Go**
4. **Browse** and select `ZaiExcelAddin-AddIn64-packed.xll`
5. Confirm

The **Z.AI** tab appears on the ribbon.

## Usage

### Login
Click **Z.AI** → **Login** → paste your z.ai API key.

### Chat with AI
Click **💬 Chat** on the Z.AI ribbon tab — the right-side pane opens.

Example prompts:
- "Read data from A1:D10"
- "Add a SUM formula to E1"
- "Bold header row with blue background"
- "Create a pie chart from A1:B5"
- "Sort by column C descending"

### Change language
**Z.AI** → **Language** → type code: `pl`, `en`, `de`, `fr`, `es`, `uk`, `zh`, `ja`

## Project structure

```
z.ai-Unofficial-Excel-Com-Add-in-addon/
├── src/ZaiExcelAddin/           # .NET COM Add-in (v2.0, Excel-DNA)
│   ├── ZaiExcelAddin.csproj     # C# project
│   ├── AddIn.cs                 # Entry point (IExcelAddIn)
│   ├── RibbonController.cs      # Ribbon + Custom Task Pane
│   ├── Models/
│   ├── Services/                # Auth, Conversation, Excel skills, I18n, API
│   └── UI/                      # WPF chat panel + host
├── legacy/                      # VBA (v1.0)
│   ├── *.bas, *.frm
│   └── build.vbs                # Builds .xlam
├── show-reel.gif
└── README*.md                   # EN + PL docs
```

## v2.0 architecture (COM Add-in)

```
┌──────────────┐    HTTP/JSON     ┌─────────────────┐
│   z.ai API   │◄───────────────►│  ZaiApiService   │
│   (GLM-4+)   │                 └────────┬────────┘
└──────────────┘                          │
                                ┌─────────▼────────┐
                                │ ConversationSvc   │ ← tool-calling loop
                                │ (max 15 rounds,   │   + loop detection
                                │  dedup detection)  │
                                └─────────┬────────┘
                                          │
                     ┌────────────────────┼────────────────────┐
                     │                    │                    │
              ┌──────▼──────┐    ┌───────▼──────┐    ┌───────▼──────┐
              │  ChatPanel   │    │ ExcelSkillSvc │    │  I18nService  │
              │  (WPF CTP)   │    │ (15 tools)    │    │  (8 langs)    │
              └─────────────┘    └──────────────┘    └──────────────┘
```

## Legacy version (VBA)

Legacy VBA (.xlam) remains available — run `cscript build.vbs` in `legacy` to build.

## License

Open-source project. Uses the z.ai API — you need an account and API key.
