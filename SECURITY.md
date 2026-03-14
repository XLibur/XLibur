# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| latest  | :white_check_mark: |

Only the latest release receives security fixes.

## Reporting a Vulnerability

If you discover a security vulnerability in XLibur, please report it responsibly.

**Do not open a public GitHub issue for security vulnerabilities.**

Instead, please report vulnerabilities by emailing the maintainers or by using [GitHub's private vulnerability reporting](https://github.com/XLibur/XLibur/security/advisories/new).

When reporting, please include:

- A description of the vulnerability
- Steps to reproduce the issue
- The potential impact
- Any suggested fix (if available)

We will acknowledge receipt within 72 hours and aim to provide a fix or mitigation plan within 30 days, depending on severity.

## Scope

XLibur processes `.xlsx` and `.xlsm` files, which are ZIP-based XML packages. Security concerns include but are not limited to:

- XML External Entity (XXE) attacks via crafted OpenXML content
- Zip bomb / decompression bomb attacks
- Path traversal via malicious package part names
- Denial of service via excessively large or deeply nested structures
