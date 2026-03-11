# CLAUDE.md

## Project Overview

ClosedXML is a .NET library for reading, manipulating, and writing Excel 2007+ (.xlsx, .xlsm) files. It provides an intuitive interface over the OpenXML API, enabling Excel file creation without the Excel application. Licensed under MIT.

- **Repository:** https://github.com/ClosedXML/ClosedXML
- **Docs:** https://closedxml.readthedocs.io/
- **Version:** 0.105.0

## Project Structure

- **ClosedXML** - Core library
- **ClosedXML.Tests** - NUnit test suite
- **ClosedXML.Examples** - Example applications
- **ClosedXML.Sandbox** - Development sandbox

Solution uses `.slnx` format (modern MSBuild Solution Extension).

## Build

- **Target Frameworks:** net8.0, net10.0
- **Nullable Reference Types:** Enabled
- **Warnings as Errors:** Enabled (TreatWarningsAsErrors=true)
- **CI:** GitHub Actions with .NET 8.0.x and 10.0.x SDKs

## Testing

- **Framework:** NUnit 4.x with NUnit3TestAdapter
- **Culture:** Tests default to en-US via SetCultureAttribute

## Key Dependencies

- **DocumentFormat.OpenXml** 3.4.1 - Core OpenXML implementation
- **ExcelNumberFormat** 1.1.0 - Excel number formatting
- **SixLabors.Fonts** 1.0.1 - Font handling
- **ClosedXML.Parser** 2.0.0 - Parser utilities
- **RBush.Signed** 4.0.0 - Spatial indexing

## Shell Commands

- Do not use compound commands (e.g., `&&`, `||`, `;`) in Bash tool calls. Run each command as a separate Bash tool invocation.
- Never use compound commands with bash or git. Each command must be its own separate Bash tool call.
- Never use `cd <folder> && git <params>` style commands. Use absolute paths or set the working directory separately.

## Dependencies

- Do NOT upgrade SixLabors.Fonts. Its newer versions have a conflicting license.
