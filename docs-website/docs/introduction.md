---
id: introduction
title: Introduction
sidebar_label: Introduction
sidebar_position: 1
slug: /
description: XLibur is a .NET 8+ library for reading, manipulating, and writing Excel 2007+ (.xlsx, .xlsm) files.
---

import useBaseUrl from '@docusaurus/useBaseUrl';

# XLibur

<img src={useBaseUrl('/img/logo.png')} alt="XLibur logo" width="512" />

[![Build and Test](https://github.com/XLibur/XLibur/actions/workflows/build-and-test.yml/badge.svg)](https://github.com/XLibur/XLibur/actions/workflows/build-and-test.yml)
[![NuGet](https://img.shields.io/nuget/v/XLibur.svg)](https://www.nuget.org/packages/XLibur)
[![NuGet Downloads](https://img.shields.io/nuget/dt/XLibur.svg)](https://www.nuget.org/packages/XLibur)
[![SonarCloud Quality Gate](https://sonarcloud.io/api/project_badges/measure?project=XLibur_XLibur&metric=alert_status)](https://sonarcloud.io/dashboard?id=XLibur_XLibur)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/XLibur/XLibur/blob/main/LICENSE)

## About

XLibur is a .NET 8+ library for reading, manipulating, and writing Excel 2007+
(.xlsx, .xlsm) files. It provides an intuitive interface over the underlying
[OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

XLibur was forked from [ClosedXML v0.105.0](https://github.com/ClosedXML/ClosedXML/)
(May 2025), created to ship patches and improvements that didn't land upstream.
Namespaces are prefixed with `XLibur` to avoid conflicts with ClosedXML if both
are referenced in the same project.

## Usage

XLibur lets you create and manipulate Excel files without Excel installed — a common
use case is generating reports on a web server.

```csharp
using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Sample Sheet");
    worksheet.Cell("A1").Value = "Hello World!";
    worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
    workbook.SaveAs("HelloWorld.xlsx");
}
```

Head to [Getting Started](./getting-started.md) for installation and a tour of the
common read/write operations.

## Migration from ClosedXML

The public API surface is largely unchanged from ClosedXML 0.105 — migrating is mostly a namespace
rename, with font engine configuration as the one packaging difference. See
[Migration from ClosedXML](./migration.md).
