# Contributing to Z.AI Excel Add-in

Thank you for your interest in contributing to the Z.AI Excel Add-in! This document provides guidelines for contributing to this project.

## Code of Conduct

This project adheres to a Code of Conduct that all contributors are expected to follow. Please read [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md) before contributing.

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check existing issues to avoid duplicates. When creating a bug report, include:

- **Clear title and description** of the issue
- **Steps to reproduce** the problem
- **Expected behavior** vs actual behavior
- **Excel version** and Windows version
- **.NET version** (should be .NET 8.0 Desktop Runtime)
- **Add-in version** (check in Z.AI → About)
- **Error messages or logs** (find logs via Z.AI → Logs)

### Suggesting Enhancements

Enhancement suggestions are welcome! Please provide:

- **Clear use case** for the enhancement
- **Description** of the proposed functionality
- **Examples** of how it would work
- **Why this would be useful** to other users

### Pull Requests

1. **Fork the repository** and create your branch from `main`
2. **Make your changes** following the coding guidelines below
3. **Test your changes** thoroughly
4. **Update documentation** if needed
5. **Commit with clear messages** describing what and why
6. **Submit a pull request**

## Development Setup

### Prerequisites

- Windows 10/11
- Visual Studio 2022 or VS Code with C# extension
- .NET SDK 8.0 or later
- Microsoft Excel 2016+ (64-bit recommended)

### Building from Source

```powershell
cd src
dotnet build -c Release
```

The output will be in `src\bin\Release\net8.0-windows\publish\ZaiExcelAddin-AddIn64-packed.xll`

### Project Structure

- `src/` - Main .NET 8 COM Add-in source code
- `src/Services/` - Core services (API, Excel skills, i18n, auth)
- `src/UI/` - WPF user interface components
- `src/Models/` - Data models
- `legacy/` - Deprecated VBA version (v1.0)

## Coding Guidelines

### C# Style

- Use **C# 12 features** when appropriate
- Follow **Microsoft C# coding conventions**
- Use **nullable reference types** (`<Nullable>enable</Nullable>`)
- Prefer **async/await** for I/O operations
- Use **meaningful variable names**

### Excel COM Interop

- Always use `Marshal.ReleaseComObject()` for COM objects
- Wrap Excel operations in try-catch blocks
- Use `Application.ScreenUpdating = false` for bulk operations
- Check for null references before accessing COM objects

### Adding New Excel Skills

When adding new Excel automation skills:

1. Add method to `ExcelSkillService.cs` (or relevant partial class)
2. Include XML documentation comments
3. Return structured results (success/error messages)
4. Handle COM exceptions gracefully
5. Test with various Excel scenarios

### Localization

- Add new UI strings to `src/i18n/*.json` files
- Support all 8 languages (en, pl, de, fr, es, uk, zh, ja)
- Use `I18nService` for retrieving translated strings

## Testing

- **Manual testing** is currently the primary method
- Test with different Excel versions
- Test with various workbook formats (.xlsx, .xlsm, .xlsb)
- Verify AI responses are appropriate
- Check for COM object leaks (use Task Manager)

## Git Commit Messages

- Use present tense ("Add feature" not "Added feature")
- Keep first line under 72 characters
- Reference issues: "Fix #123"
- Be descriptive about what and why, not just what

Examples:
```
Add support for conditional formatting rules

Implements three types of conditional formatting:
- Highlight cells rules
- Color scales
- Data bars

Fixes #45
```

## Questions?

Feel free to open an issue with your question or reach out through GitHub Discussions.

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
