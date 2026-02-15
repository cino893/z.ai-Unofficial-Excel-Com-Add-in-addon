# Security Policy

## Supported Versions

We release patches for security vulnerabilities for the following versions:

| Version | Supported          |
| ------- | ------------------ |
| 2.x.x   | :white_check_mark: |
| 1.x.x   | :x: (Legacy VBA)   |

## Reporting a Vulnerability

We take the security of Z.AI Excel Add-in seriously. If you believe you have found a security vulnerability, please report it to us as described below.

### How to Report

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, please report them via:

1. **GitHub Security Advisories**: Use the [Security tab](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/security/advisories/new) to privately report a vulnerability
2. **Create a private security advisory** with details about the vulnerability

### What to Include

Please include as much of the following information as possible:

- Type of vulnerability (e.g., buffer overflow, SQL injection, cross-site scripting, etc.)
- Full paths of source file(s) related to the vulnerability
- Location of the affected source code (tag/branch/commit or direct URL)
- Any special configuration required to reproduce the issue
- Step-by-step instructions to reproduce the issue
- Proof-of-concept or exploit code (if possible)
- Impact of the issue, including how an attacker might exploit it

### Response Timeline

- We will acknowledge receipt of your vulnerability report within 48 hours
- We will provide a more detailed response within 7 days, indicating next steps
- We will keep you informed of the progress towards resolving the issue
- Once the vulnerability is fixed, we will publicly disclose it (with credit to you, if desired)

## Security Best Practices for Users

### API Key Security

- **Never commit your API key** to version control
- **Store API keys securely** (the add-in stores keys in Windows Registry, encrypted by Windows DPAPI)
- **Use API keys with minimal required permissions**
- **Rotate API keys regularly**
- **Monitor API key usage** via the [Z.AI billing page](https://z.ai/manage-apikey/billing)

### Excel File Security

- **Be cautious with untrusted workbooks** - the add-in can read and modify Excel data
- **Review AI actions** before accepting them, especially:
  - Formula insertions that reference external data
  - Macro or VBA code generation
  - Data exports or copies
- **Use Excel's native security features**:
  - Password-protect sensitive workbooks
  - Enable Protected View for files from the internet
  - Keep Excel and Windows updated

### Network Security

- The add-in communicates with `https://z.ai/` over HTTPS
- No data is sent to third parties other than Z.AI
- Review the [Z.AI privacy policy](https://z.ai) for details on data handling

### Code Security

- **Verify downloads**: Download the `.xll` file only from official [GitHub Releases](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/releases)
- **Check signatures**: Verify the file hash matches the release notes
- **Build from source**: For maximum security, build the add-in from source code yourself

### System Security

- **Keep .NET updated**: Install the latest [.NET 8.0 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)
- **Keep Excel updated**: Use the latest version of Microsoft Excel with security patches
- **Keep Windows updated**: Ensure Windows security updates are installed
- **Use antivirus software**: Keep your antivirus definitions up to date

## Known Security Considerations

### COM Interop

This add-in uses COM Interop to control Excel, which requires:
- Running with the same privileges as Excel
- Access to the Windows Registry (for storing API keys)
- Ability to read/write all open Excel workbooks

### Third-Party Dependencies

- **ExcelDna** (1.9.0): Used for Excel COM add-in infrastructure
- **.NET 8.0**: Runtime and libraries from Microsoft

We regularly review and update dependencies to address security vulnerabilities.

## Security Updates

Security updates will be released as needed. Users are encouraged to:
- Watch the repository for releases
- Enable GitHub notifications for security advisories
- Update to the latest version promptly

## Contact

For security-related questions that are not vulnerabilities, please open a public issue with the label "security".

## Attribution

This security policy is based on industry best practices and GitHub's recommended security policy template.
