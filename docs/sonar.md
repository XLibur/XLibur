# SonarQube issues — XLibur_XLibur

- Server: https://sonarcloud.io
- Filters: impactSeverities=`MEDIUM`, issueStatuses=`OPEN,CONFIRMED`
- Total issues: 150

Each item below is one issue to fix. The path and line locate the code; the rule and message describe the problem.

---

## Triage status

Every issue below carries a `Status:` line. 68 are worth fixing, 82 are accepted or
false positives. The verdict is per rule — no rule splits across both outcomes.

### Verified against current `main`, not taken on trust

The export predates the nine commits merged at `dfd5f49b`, so its line numbers had
drifted (`DynamicArray.cs:734` is now different code). Rather than guess, the findings
were reproduced locally:

- **Roslyn rules (113):** the project build is clean — these rules are off by default
  and only fire under SonarCloud's analyzer configuration. Re-enabling exactly the
  reported rule IDs via `.editorconfig` and building the solution reproduces
  **112 of 113**, matching every rule's count. The one gone is a `CA2249` in
  `SheetReference.cs`, fixed by the commits above.
- **Sonar rules (37):** reproduced by referencing `SonarAnalyzer.CSharp` from
  `Directory.Build.props`. All rules reproduce except `S107`, which needs a
  `SonarLint.xml` to supply its parameter threshold.

Both measurement setups were reverted; neither is committed, for the reason given in
`68487738` — adding the analyzers to every build would promote hundreds of unrelated
rules to errors under `TreatWarningsAsErrors`.

The line numbers in the entries below are the exported ones. Current locations were
taken from the local build.

### Fix — 68

| Rule | Count | Action |
|---|---:|---|
| `CA1859` | 19 | Tighten to the concrete type where the member is not public API |
| `CA1510` | 9 | `ArgumentNullException.ThrowIfNull` |
| `S3358` | 5 | Lift the inner ternary into a named local |
| `S125` | 2 | Delete the commented-out code — the other 2 are prose, not code |
| `S6562` | 4 | Pass `DateTimeKind` explicitly in sample data |
| `CA2249` | 4 | `string.Contains` |
| `CA1806` | 3 | Make the tolerated-parse fallback explicit |
| `CA1829` | 3 | Use the `Count` property |
| `S1172` | 2 | Drop the unused parameter |
| `S3928` + `CA2208` | 2 + 2 | Real defect — see below |
| `CA1826` | 2 | Index the collection directly |
| `SYSLIB1045` | 2 | `[GeneratedRegex]` partial method |
| `S1066`, `S1118`, `S4581`, `CA1513`, `CA1845`, `CA1846` | 1 each | Mechanical |

Three entries were triaged as fixes and reversed once the code was read — one `CA1822`
and two `S125`. Each says so in its own entry. The final split is **65 fix / 85 ignore**.

**The one real defect.** `XLPictures.ValidateMembersAndComputeBounds(List<XLPicture> members)`
throws `ArgumentException` with the literal `paramName` `"pictures"`, which is not a
parameter of that method — so the exception names something the caller cannot map.
`S3928` and `CA2208` are the same two sites reported by both engines.

### Ignore — 82

| Rule | Count | Why |
|---|---:|---|
| `CA1861` | 58 | The inline array **is** the test data and belongs beside its assertion. Hoisting 58 literals into static fields to save one allocation per test run is a net readability loss. Scoped off for test projects in `.editorconfig` |
| `S107` | 11 | OOXML readers and writers carry one parameter per attribute; a parameter object renames the problem rather than solving it. Suppressed per method, matching the `S3776` convention from `68487738` |
| `TUnit0057` | 4 | Informational. Adding an unused `AssemblyHookContext` parameter to satisfy it is strictly worse |
| `S2234` | 3 | **False positive** |
| `TUnitAssertions0016` | 3 | **False positive** |
| `S127` | 2 | The loop variable is advanced deliberately while scanning a token run |
| `S1871` | 1 | The two arms are the two XLOOKUP match modes; merging the conditions hides that |

**Why the two false positives are false.**

`S2234` says arguments to `Invoke` are out of order in `SignatureAdapter`. They are not.
`Func<>` names its own generic parameters `arg1..argN`, while the adapter names its locals
`arg0..argN-1`; the rule matches them by name and sees a shift. Positionally the call is
correct, and the calc-engine tests covering these functions pass.

The suppression added for it in `1fcb9b82` did not close the findings, and the reason is
worth recording: `SignatureAdapter.cs` already carried a run of narrow
`#pragma warning disable`/`restore S2234` pairs, and the first `restore` cancels the
file-level `disable` placed above the class. Everything past that point is unsuppressed
unless it carries its own pair. Three forwarding calls did not, so they kept reporting.
They now do, matching the pairs already around the neighbouring adapters — a new adapter
below that line needs the same pair around its call.

`TUnitAssertions0016` says `.IsEqualTo(...)` on a collection compares by reference.
`Area` is a `readonly struct : IEquatable<Area>, IEnumerable<Point>` with `Equals` and
`GetHashCode` — it trips the rule only because it also enumerates points. Switching to
`IsEquivalentTo` would compare point sequences instead of the area's identity: a weaker
assertion, and the wrong one.

### Delivery

Stacked pull requests, each branching from the one before:

| # | Branch | Contents |
|---:|---|---|
| 1 | `docs/sonar-triage` | This triage |
| 2 | `fix/sonar-correctness` | `CA2208`/`S3928`, `CA1806`, `S1172`, `S4581` |
| 3 | `refactor/sonar-modern-apis` | `CA1510`, `CA1513`, `CA2249` |
| 4 | `perf/sonar-micro` | `CA1829`, `CA1826`, `CA1845`, `CA1846`, `CA1822`, `CA1859` (library) |
| 5 | `refactor/sonar-test-and-sample` | `CA1859` (tests), `SYSLIB1045`, `S1118`, `S6562`, `S125`, `S1066`, `S3358` |
| 6 | `chore/sonar-accepted` | Suppressions with reasons for the 82 accepted |

---

## `XLibur.Benchmarks/OpenXmlWorkbookBenchmarks.cs`

### 1. Method has 8 parameters, which is greater than the 7 authorized.
- Location: `XLibur.Benchmarks/OpenXmlWorkbookBenchmarks.cs:302`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ0ECRxpoApNP_XKBs0I
- Fix guidance: see rule `csharpsquid:S107` below.

## `XLibur.Benchmarks/Program.cs`

### 2. Remove this commented out code.
- Location: `XLibur.Benchmarks/Program.cs:14`
- **Status:** IGNORE - FALSE POSITIVE on inspection: the flagged lines are prose describing the benchmark modes, not commented-out code.
- **Status:** FIX - delete the commented-out code.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-jJ36vmmYrCxinzLl-
- Fix guidance: see rule `csharpsquid:S125` below.

## `XLibur.Fonts.SixLabors.Examples/FontEngine/UsingSixLaborsFontsV2.cs`

### 3. Add a 'protected' constructor or the 'static' keyword to the class declaration.
- Location: `XLibur.Fonts.SixLabors.Examples/FontEngine/UsingSixLaborsFontsV2.cs:10`
- Rule: `csharpsquid:S1118` — Utility classes should not have public constructors
- **Status:** FIX - make the example class static.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ08wR7cWmK3uKXKmtMm
- Fix guidance: see rule `csharpsquid:S1118` below.

## `XLibur.Fonts.SixLabors.Tests/TestDefaults.cs`

### 4. Hook method can accept a AssemblyHookContext parameter for additional context information
- Location: `XLibur.Fonts.SixLabors.Tests/TestDefaults.cs:14`
- Rule: `external_roslyn:TUnit0057` — roslyn:TUnit0057
- **Status:** IGNORE - informational; adding an unused AssemblyHookContext parameter to satisfy it is strictly worse.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RuliXpRjQDkdDXe
- Fix guidance: see rule `external_roslyn:TUnit0057` below.

## `XLibur.Fonts.SkiaSharp.Tests/TestDefaults.cs`

### 5. Hook method can accept a AssemblyHookContext parameter for additional context information
- Location: `XLibur.Fonts.SkiaSharp.Tests/TestDefaults.cs:15`
- Rule: `external_roslyn:TUnit0057` — roslyn:TUnit0057
- **Status:** IGNORE - informational; adding an unused AssemblyHookContext parameter to satisfy it is strictly worse.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RvjiXpRjQDkdDXf
- Fix guidance: see rule `external_roslyn:TUnit0057` below.

## `XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs`

### 6. Change type of field '_streamFonts' from 'System.Collections.Generic.IReadOnlyDictionary<string, SkiaSharp.SKTypeface>' to 'System.Collections.Generic.Dictionary<string, SkiaSharp.SKTypeface>' for improved performance
- Location: `XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs:20`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7fwn6rRNC8MhBOEGiV
- Fix guidance: see rule `external_roslyn:CA1859` below.

### 7. Change type of parameter 'fonts' from 'System.Collections.Generic.IDictionary<string, SkiaSharp.SKTypeface>' to 'System.Collections.Generic.Dictionary<string, SkiaSharp.SKTypeface>' for improved performance
- Location: `XLibur.Fonts.SkiaSharp/SkiaSharpFontEngine.cs:235`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7fwn6rRNC8MhBOEGiW
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur.Report.Benchmarks/ReportData.cs`

### 8. Provide the "DateTimeKind" when creating this object.
- Location: `XLibur.Report.Benchmarks/ReportData.cs:52`
- Rule: `csharpsquid:S6562` — Always set the "DateTimeKind" when creating new "DateTime" instances
- **Status:** FIX - pass DateTimeKind explicitly in sample data.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t87zyrxBNGQ_o419
- Fix guidance: see rule `csharpsquid:S6562` below.

## `XLibur.Report.DynamicLinq/DynamicLinqExpressionEngine.cs`

### 9. Use 'ArgumentNullException.ThrowIfNull' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report.DynamicLinq/DynamicLinqExpressionEngine.cs:115`
- Rule: `external_roslyn:CA1510` — roslyn:CA1510
- **Status:** FIX - ArgumentNullException.ThrowIfNull.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8-HyrxBNGQ_o42E
- Fix guidance: see rule `external_roslyn:CA1510` below.

### 10. Use 'ArgumentNullException.ThrowIfNull' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report.DynamicLinq/DynamicLinqExpressionEngine.cs:120`
- Rule: `external_roslyn:CA1510` — roslyn:CA1510
- **Status:** FIX - ArgumentNullException.ThrowIfNull.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8-HyrxBNGQ_o42F
- Fix guidance: see rule `external_roslyn:CA1510` below.

### 11. Use 'ArgumentNullException.ThrowIfNull' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report.DynamicLinq/DynamicLinqExpressionEngine.cs:133`
- Rule: `external_roslyn:CA1510` — roslyn:CA1510
- **Status:** FIX - ArgumentNullException.ThrowIfNull.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8-HyrxBNGQ_o42G
- Fix guidance: see rule `external_roslyn:CA1510` below.

### 12. Use 'ArgumentNullException.ThrowIfNull' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report.DynamicLinq/DynamicLinqExpressionEngine.cs:138`
- Rule: `external_roslyn:CA1510` — roslyn:CA1510
- **Status:** FIX - ArgumentNullException.ThrowIfNull.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8-HyrxBNGQ_o42H
- Fix guidance: see rule `external_roslyn:CA1510` below.

## `XLibur.Report.Examples/AnnualSalesReport.cs`

### 13. Provide the "DateTimeKind" when creating this object.
- Location: `XLibur.Report.Examples/AnnualSalesReport.cs:147`
- Rule: `csharpsquid:S6562` — Always set the "DateTimeKind" when creating new "DateTime" instances
- **Status:** FIX - pass DateTimeKind explicitly in sample data.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8wXyrxBNGQ_o41O
- Fix guidance: see rule `csharpsquid:S6562` below.

## `XLibur.Report.Examples/EverythingAtOnce.cs`

### 14. Provide the "DateTimeKind" when creating this object.
- Location: `XLibur.Report.Examples/EverythingAtOnce.cs:216`
- Rule: `csharpsquid:S6562` — Always set the "DateTimeKind" when creating new "DateTime" instances
- **Status:** FIX - pass DateTimeKind explicitly in sample data.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8t8yrxBNGQ_o41F
- Fix guidance: see rule `csharpsquid:S6562` below.

## `XLibur.Report.Examples/Program.cs`

### 15. Do not update the stop condition variable 'i' in the body of the for loop.
- Location: `XLibur.Report.Examples/Program.cs:101`
- Rule: `csharpsquid:S127` — "for" loop stop conditions should be invariant
- **Status:** IGNORE - accepted: the loop variable is advanced deliberately while scanning a token run.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8uoyrxBNGQ_o41G
- Fix guidance: see rule `csharpsquid:S127` below.

## `XLibur.Report.Examples/SalesData.cs`

### 16. Provide the "DateTimeKind" when creating this object.
- Location: `XLibur.Report.Examples/SalesData.cs:64`
- Rule: `csharpsquid:S6562` — Always set the "DateTimeKind" when creating new "DateTime" instances
- **Status:** FIX - pass DateTimeKind explicitly in sample data.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8v5yrxBNGQ_o41N
- Fix guidance: see rule `csharpsquid:S6562` below.

## `XLibur.Report.Tests/ExampleSmokeTests.cs`

### 17. Change return type of method 'RunAll' from 'System.Collections.Generic.IReadOnlyList<XLibur.Report.Examples.ExampleRun>' to 'System.Collections.Generic.List<XLibur.Report.Examples.ExampleRun>' for improved performance
- Location: `XLibur.Report.Tests/ExampleSmokeTests.cs:37`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t845yrxBNGQ_o415
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur.Report.Tests/Functions/ExcelFunctionBridgeTests.cs`

### 18. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Functions/ExcelFunctionBridgeTests.cs:41`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t851yrxBNGQ_o417
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 19. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Functions/ExcelFunctionBridgeTests.cs:63`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t851yrxBNGQ_o418
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Report.Tests/Infrastructure/GoldenFile.cs`

### 20. Change return type of method 'LoadTemplate' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Infrastructure/GoldenFile.cs:52`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t82nyrxBNGQ_o410
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur.Report.Tests/Infrastructure/WorkbookComparerTests.cs`

### 21. Change return type of method 'Sheet' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Infrastructure/WorkbookComparerTests.cs:10`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t82AyrxBNGQ_o41z
- Fix guidance: see rule `external_roslyn:CA1859` below.

### 22. Do not use Enumerable methods on indexable collections. Instead use the collection directly.
- Location: `XLibur.Report.Tests/Infrastructure/WorkbookComparerTests.cs:266`
- Rule: `external_roslyn:CA1826` — roslyn:CA1826
- **Status:** FIX - index the collection directly.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t82AyrxBNGQ_o41y
- Fix guidance: see rule `external_roslyn:CA1826` below.

## `XLibur.Report.Tests/Ranges/HorizontalRangeTests.cs`

### 23. Change return type of method 'Template' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Ranges/HorizontalRangeTests.cs:24`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t80yyrxBNGQ_o41w
- Fix guidance: see rule `external_roslyn:CA1859` below.

### 24. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Ranges/HorizontalRangeTests.cs:69`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t80yyrxBNGQ_o41q
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 25. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Ranges/HorizontalRangeTests.cs:80`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t80yyrxBNGQ_o41r
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 26. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Ranges/HorizontalRangeTests.cs:81`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t80yyrxBNGQ_o41s
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 27. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Ranges/HorizontalRangeTests.cs:170`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t80yyrxBNGQ_o41t
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 28. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Ranges/HorizontalRangeTests.cs:209`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t80yyrxBNGQ_o41u
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 29. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Ranges/HorizontalRangeTests.cs:240`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t80yyrxBNGQ_o41v
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Report.Tests/Ranges/RangeExpansionTests.cs`

### 30. Change return type of method 'TemplateWithItemsRange' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Ranges/RangeExpansionTests.cs:26`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t80UyrxBNGQ_o41p
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur.Report.Tests/Rewriting/ChartRewritingTests.cs`

### 31. Change return type of method 'Template' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Rewriting/ChartRewritingTests.cs:29`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t83HyrxBNGQ_o411
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur.Report.Tests/Rewriting/PicturePlacementTests.cs`

### 32. Change return type of method 'Template' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Rewriting/PicturePlacementTests.cs:30`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t84ZyrxBNGQ_o414
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur.Report.Tests/Rewriting/PivotRewritingTests.cs`

### 33. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Rewriting/PivotRewritingTests.cs:91`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t837yrxBNGQ_o413
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Report.Tests/Rewriting/SheetReferenceTests.cs`

### 34. Parse calls TryParse but does not explicitly check whether the conversion succeeded. Either use the return value in a conditional statement or verify that the call site expects that the out argument will be set to the default value when the conversion fails.
- Location: `XLibur.Report.Tests/Rewriting/SheetReferenceTests.cs:10`
- Rule: `external_roslyn:CA1806` — roslyn:CA1806
- **Status:** FIX - make the tolerated-parse fallback explicit rather than dropping the bool.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t83fyrxBNGQ_o412
- Fix guidance: see rule `external_roslyn:CA1806` below.

## `XLibur.Report.Tests/Tags/CultureBoundOrderingTests.cs`

### 35. Change return type of method 'Template' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Tags/CultureBoundOrderingTests.cs:59`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-6mdpjCO3a6BsbatnZ
- Fix guidance: see rule `external_roslyn:CA1859` below.

### 36. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/CultureBoundOrderingTests.cs:104`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-6mdpjCO3a6BsbatnV
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 37. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/CultureBoundOrderingTests.cs:115`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-6mdpjCO3a6BsbatnW
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 38. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/CultureBoundOrderingTests.cs:153`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-6mdpjCO3a6BsbatnX
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 39. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/CultureBoundOrderingTests.cs:164`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-6mdpjCO3a6BsbatnY
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Report.Tests/Tags/GroupTagTests.cs`

### 40. Change return type of method 'Template' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:26`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41n
- Fix guidance: see rule `external_roslyn:CA1859` below.

### 41. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:74`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41e
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 42. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:75`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41f
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 43. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:182`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41g
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 44. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:240`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41h
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 45. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:253`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41i
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 46. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:266`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41j
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 47. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:278`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41k
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 48. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:304`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41l
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 49. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/GroupTagTests.cs:389`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zUyrxBNGQ_o41m
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Report.Tests/Tags/IfTagTests.cs`

### 50. Change return type of method 'Template' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Tags/IfTagTests.cs:23`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8yXyrxBNGQ_o41a
- Fix guidance: see rule `external_roslyn:CA1859` below.

### 51. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/IfTagTests.cs:71`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8yXyrxBNGQ_o41W
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 52. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/IfTagTests.cs:84`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8yXyrxBNGQ_o41X
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 53. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/IfTagTests.cs:116`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8yXyrxBNGQ_o41Y
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 54. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/IfTagTests.cs:164`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8yXyrxBNGQ_o41Z
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Report.Tests/Tags/PivotTagTests.cs`

### 55. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/PivotTagTests.cs:110`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8y2yrxBNGQ_o41b
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 56. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/PivotTagTests.cs:154`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8y2yrxBNGQ_o41c
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 57. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/PivotTagTests.cs:205`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8y2yrxBNGQ_o41d
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Report.Tests/Tags/TagBehaviourTests.cs`

### 58. Change return type of method 'Template' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/Tags/TagBehaviourTests.cs:23`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8x0yrxBNGQ_o41V
- Fix guidance: see rule `external_roslyn:CA1859` below.

### 59. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/TagBehaviourTests.cs:65`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8x0yrxBNGQ_o41T
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 60. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/TagBehaviourTests.cs:76`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8x0yrxBNGQ_o41Q
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 61. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/TagBehaviourTests.cs:87`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8x0yrxBNGQ_o41R
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 62. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/TagBehaviourTests.cs:98`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8x0yrxBNGQ_o41U
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 63. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/TagBehaviourTests.cs:110`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8x0yrxBNGQ_o41S
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Report.Tests/Tags/TagParserTests.cs`

### 64. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Report.Tests/Tags/TagParserTests.cs:35`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8zzyrxBNGQ_o41o
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Report.Tests/TestInfrastructure.cs`

### 65. Hook method can accept a AssemblyHookContext parameter for additional context information
- Location: `XLibur.Report.Tests/TestInfrastructure.cs:16`
- Rule: `external_roslyn:TUnit0057` — roslyn:TUnit0057
- **Status:** IGNORE - informational; adding an unused AssemblyHookContext parameter to satisfy it is strictly worse.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t85TyrxBNGQ_o416
- Fix guidance: see rule `external_roslyn:TUnit0057` below.

## `XLibur.Report.Tests/XLTemplateTests.cs`

### 66. Change return type of method 'WorkbookWith' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report.Tests/XLTemplateTests.cs:12`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t81hyrxBNGQ_o41x
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur.Report/Expressions/ExpressionScope.cs`

### 67. Use 'ArgumentNullException.ThrowIfNull' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report/Expressions/ExpressionScope.cs:32`
- Rule: `external_roslyn:CA1510` — roslyn:CA1510
- **Status:** FIX - ArgumentNullException.ThrowIfNull.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8payrxBNGQ_o40u
- Fix guidance: see rule `external_roslyn:CA1510` below.

## `XLibur.Report/Expressions/ScribanExpressionEngine.cs`

### 68. Use 'ArgumentNullException.ThrowIfNull' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report/Expressions/ScribanExpressionEngine.cs:65`
- Rule: `external_roslyn:CA1510` — roslyn:CA1510
- **Status:** FIX - ArgumentNullException.ThrowIfNull.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8o_yrxBNGQ_o40r
- Fix guidance: see rule `external_roslyn:CA1510` below.

### 69. Use 'ArgumentNullException.ThrowIfNull' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report/Expressions/ScribanExpressionEngine.cs:95`
- Rule: `external_roslyn:CA1510` — roslyn:CA1510
- **Status:** FIX - ArgumentNullException.ThrowIfNull.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8o_yrxBNGQ_o40s
- Fix guidance: see rule `external_roslyn:CA1510` below.

### 70. Use 'ArgumentNullException.ThrowIfNull' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report/Expressions/ScribanExpressionEngine.cs:107`
- Rule: `external_roslyn:CA1510` — roslyn:CA1510
- **Status:** FIX - ArgumentNullException.ThrowIfNull.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8o_yrxBNGQ_o40t
- Fix guidance: see rule `external_roslyn:CA1510` below.

## `XLibur.Report/Ranges/RangeExpander.cs`

### 71. Method has 11 parameters, which is greater than the 7 authorized.
- Location: `XLibur.Report/Ranges/RangeExpander.cs:211`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8qiyrxBNGQ_o400
- Fix guidance: see rule `csharpsquid:S107` below.

### 72. Remove this unused method parameter 'area'.
- Location: `XLibur.Report/Ranges/RangeExpander.cs:478`
- Rule: `csharpsquid:S1172` — Unused method parameters should be removed
- **Status:** FIX - drop the unused parameter (private method, no API impact).
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8qiyrxBNGQ_o401
- Fix guidance: see rule `csharpsquid:S1172` below.

## `XLibur.Report/Rewriting/PivotBuilder.cs`

### 73. Change type of parameter 'workbook' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report/Rewriting/PivotBuilder.cs:142`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8s2yrxBNGQ_o407
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur.Report/Rewriting/SheetReference.cs`

### 74. Use 'string.Contains' instead of 'string.IndexOf' to improve readability
- Location: `XLibur.Report/Rewriting/SheetReference.cs:43`
- Rule: `external_roslyn:CA2249` — roslyn:CA2249
- **Status:** FIX - string.Contains.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8tTyrxBNGQ_o40-
- Fix guidance: see rule `external_roslyn:CA2249` below.

### 75. Use 'string.Contains' instead of 'string.IndexOf' to improve readability
- Location: `XLibur.Report/Rewriting/SheetReference.cs:43`
- Rule: `external_roslyn:CA2249` — roslyn:CA2249
- **Status:** FIX - string.Contains.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8tTyrxBNGQ_o40_
- Fix guidance: see rule `external_roslyn:CA2249` below.

### 76. Use 'string.Contains' instead of 'string.IndexOf' to improve readability
- Location: `XLibur.Report/Rewriting/SheetReference.cs:158`
- Rule: `external_roslyn:CA2249` — roslyn:CA2249
- **Status:** FIX - string.Contains.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8tTyrxBNGQ_o41A
- Fix guidance: see rule `external_roslyn:CA2249` below.

## `XLibur.Report/Tags/OptionTag.cs`

### 77. Constructor has 12 parameters, which is greater than the 7 authorized.
- Location: `XLibur.Report/Tags/OptionTag.cs:74`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8nyyrxBNGQ_o40o
- Fix guidance: see rule `csharpsquid:S107` below.

## `XLibur.Report/Tags/PivotTag.cs`

### 78. Change type of parameter 'workbook' from 'XLibur.Excel.IXLWorkbook' to 'XLibur.Excel.XLWorkbook' for improved performance
- Location: `XLibur.Report/Tags/PivotTag.cs:181`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8oWyrxBNGQ_o40q
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur.Report/XLTemplate.cs`

### 79. Use 'ArgumentNullException.ThrowIfNull' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report/XLTemplate.cs:120`
- Rule: `external_roslyn:CA1510` — roslyn:CA1510
- **Status:** FIX - ArgumentNullException.ThrowIfNull.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8p9yrxBNGQ_o40x
- Fix guidance: see rule `external_roslyn:CA1510` below.

### 80. Use 'ObjectDisposedException.ThrowIf' instead of explicitly throwing a new exception instance
- Location: `XLibur.Report/XLTemplate.cs:206`
- Rule: `external_roslyn:CA1513` — roslyn:CA1513
- **Status:** FIX - ObjectDisposedException.ThrowIf.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-1t8p9yrxBNGQ_o40y
- Fix guidance: see rule `external_roslyn:CA1513` below.

## `XLibur.Tests/Excel/AutoFilters/AutoFilterRoundTripTests.cs`

### 81. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/AutoFilters/AutoFilterRoundTripTests.cs:112`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-7EIFXy0K1yczgewo2
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/Charts/ChartAnchorTests.cs`

### 82. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Charts/ChartAnchorTests.cs:172`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RaBiXpRjQDkdDXU
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/Charts/ChartGroupKindReaderTests.cs`

### 83. Use the "Count" property instead of Enumerable.Count()
- Location: `XLibur.Tests/Excel/Charts/ChartGroupKindReaderTests.cs:77`
- Rule: `external_roslyn:CA1829` — roslyn:CA1829
- **Status:** FIX - use the Count property.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RZdiXpRjQDkdDXT
- Fix guidance: see rule `external_roslyn:CA1829` below.

## `XLibur.Tests/Excel/Charts/ChartLegendAndAxisTests.cs`

### 84. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Charts/ChartLegendAndAxisTests.cs:278`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RbOiXpRjQDkdDXW
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/Charts/ChartPrimaryAxisReaderTests.cs`

### 85. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Charts/ChartPrimaryAxisReaderTests.cs:60`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-cbdSZsyU_O2b4Rqso
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 86. Use the "Count" property instead of Enumerable.Count()
- Location: `XLibur.Tests/Excel/Charts/ChartPrimaryAxisReaderTests.cs:92`
- Rule: `external_roslyn:CA1829` — roslyn:CA1829
- **Status:** FIX - use the Count property.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RakiXpRjQDkdDXV
- Fix guidance: see rule `external_roslyn:CA1829` below.

## `XLibur.Tests/Excel/Charts/ChartSeriesFormattingTests.cs`

### 87. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Charts/ChartSeriesFormattingTests.cs:298`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RYwiXpRjQDkdDXR
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 88. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Charts/ChartSeriesFormattingTests.cs:307`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RYwiXpRjQDkdDXS
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/Columns/ColumnsUsedLinqTests.cs`

### 89. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Columns/ColumnsUsedLinqTests.cs:26`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RdyiXpRjQDkdDXX
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 90. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Columns/ColumnsUsedLinqTests.cs:47`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RdyiXpRjQDkdDXY
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 91. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Columns/ColumnsUsedLinqTests.cs:69`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RdyiXpRjQDkdDXZ
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 92. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Columns/ColumnsUsedLinqTests.cs:97`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RdyiXpRjQDkdDXa
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/Comments/ThreadedCommentReviewTests.cs`

### 93. Use 'GeneratedRegexAttribute' to generate the regular expression implementation at compile-time.
- Location: `XLibur.Tests/Excel/Comments/ThreadedCommentReviewTests.cs:127`
- Rule: `external_roslyn:SYSLIB1045` — roslyn:SYSLIB1045
- **Status:** FIX - [GeneratedRegex] partial method.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-goiFXyIH_nJ2fTUr3

## `XLibur.Tests/Excel/ConditionalFormats/ConditionalFormatRangeShiftTests.cs`

### 94. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/ConditionalFormats/ConditionalFormatRangeShiftTests.cs:92`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RVdiXpRjQDkdDXP
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/Coordinates/AreaTests.cs`

### 95. `.IsEqualTo(...)` on a collection compares by reference - use `.IsEquivalentTo(...)` to compare contents
- Location: `XLibur.Tests/Excel/Coordinates/AreaTests.cs:63`
- Rule: `external_roslyn:TUnitAssertions0016` — roslyn:TUnitAssertions0016
- **Status:** IGNORE - FALSE POSITIVE: Area is a readonly struct with IEquatable value equality; it only trips the rule because it also implements IEnumerable<Point>. IsEquivalentTo would be a weaker and wrong assertion.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RSViXpRjQDkdDXO
- Fix guidance: see rule `external_roslyn:TUnitAssertions0016` below.

## `XLibur.Tests/Excel/DataValidations/DataValidationDropOnInsertTests.cs`

### 96. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/DataValidations/DataValidationDropOnInsertTests.cs:31`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RQliXpRjQDkdDXM
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 97. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/DataValidations/DataValidationDropOnInsertTests.cs:51`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RQliXpRjQDkdDXN
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/DataValidations/EmptyDataValidationTests.cs`

### 98. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/DataValidations/EmptyDataValidationTests.cs:54`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-2qOGK20ESPB3hI3Sh
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 99. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/DataValidations/EmptyDataValidationTests.cs:69`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-2qOGK20ESPB3hI3Si
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 100. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/DataValidations/EmptyDataValidationTests.cs:91`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-2qOGK20ESPB3hI3Sj
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/Drawings/GroupedPictureTests.cs`

### 101. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Drawings/GroupedPictureTests.cs:384`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7aPV-N5NXmL4P0hn96
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 102. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Drawings/GroupedPictureTests.cs:508`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RX1iXpRjQDkdDXQ
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/NamedRanges/DefinedNameStructuredReferenceTests.cs`

### 103. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/NamedRanges/DefinedNameStructuredReferenceTests.cs:128`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-7MjT1F3F_S2XkP5Lb
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 104. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/NamedRanges/DefinedNameStructuredReferenceTests.cs:145`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-7MjT1F3F_S2XkP5Lc
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/PivotTables/PivotFilterCriteriaTests.cs`

### 105. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/PivotTables/PivotFilterCriteriaTests.cs:42`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-7EIOey0K1yczgewo3
- Fix guidance: see rule `external_roslyn:CA1861` below.

### 106. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/PivotTables/PivotFilterCriteriaTests.cs:121`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-7TUcF-CkDYWncqV3G
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/Streaming/StreamingWriteTests.cs`

### 107. Prefer 'static readonly' fields over constant array arguments if the called method is called repeatedly and is not mutating the passed array
- Location: `XLibur.Tests/Excel/Streaming/StreamingWriteTests.cs:309`
- Rule: `external_roslyn:CA1861` — roslyn:CA1861
- **Status:** IGNORE - accepted: the inline array IS the test data and belongs next to its assertion. Hoisting 58 literals into static fields to save an allocation per test is a net readability loss. Scoped off in .editorconfig for test projects.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-jJ3urmmYrCxinzLl4
- Fix guidance: see rule `external_roslyn:CA1861` below.

## `XLibur.Tests/Excel/Styles/XLColorAutomaticTests.cs`

### 108. Use 'GeneratedRegexAttribute' to generate the regular expression implementation at compile-time.
- Location: `XLibur.Tests/Excel/Styles/XLColorAutomaticTests.cs:102`
- Rule: `external_roslyn:SYSLIB1045` — roslyn:SYSLIB1045
- **Status:** FIX - [GeneratedRegex] partial method.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-cn9gBsyU_O2b4SZMc

## `XLibur.Tests/Extensions/ReferenceAreaExtensionsTests.cs`

### 109. `.IsEqualTo(...)` on a collection compares by reference - use `.IsEquivalentTo(...)` to compare contents
- Location: `XLibur.Tests/Extensions/ReferenceAreaExtensionsTests.cs:18`
- Rule: `external_roslyn:TUnitAssertions0016` — roslyn:TUnitAssertions0016
- **Status:** IGNORE - FALSE POSITIVE: Area is a readonly struct with IEquatable value equality; it only trips the rule because it also implements IEnumerable<Point>. IsEquivalentTo would be a weaker and wrong assertion.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5Rh8iXpRjQDkdDXb
- Fix guidance: see rule `external_roslyn:TUnitAssertions0016` below.

### 110. `.IsEqualTo(...)` on a collection compares by reference - use `.IsEquivalentTo(...)` to compare contents
- Location: `XLibur.Tests/Extensions/ReferenceAreaExtensionsTests.cs:25`
- Rule: `external_roslyn:TUnitAssertions0016` — roslyn:TUnitAssertions0016
- **Status:** IGNORE - FALSE POSITIVE: Area is a readonly struct with IEquatable value equality; it only trips the rule because it also implements IEnumerable<Point>. IsEquivalentTo would be a weaker and wrong assertion.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5Rh8iXpRjQDkdDXc
- Fix guidance: see rule `external_roslyn:TUnitAssertions0016` below.

## `XLibur.Tests/TestInfrastructure.cs`

### 111. Hook method can accept a AssemblyHookContext parameter for additional context information
- Location: `XLibur.Tests/TestInfrastructure.cs:29`
- Rule: `external_roslyn:TUnit0057` — roslyn:TUnit0057
- **Status:** IGNORE - informational; adding an unused AssemblyHookContext parameter to satisfy it is strictly worse.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-c5RiliXpRjQDkdDXd
- Fix guidance: see rule `external_roslyn:TUnit0057` below.

## `XLibur/Excel/CalcEngine/Functions/Distributions.cs`

### 112. Extract this nested ternary operation into an independent statement.
- Location: `XLibur/Excel/CalcEngine/Functions/Distributions.cs:599`
- Rule: `csharpsquid:S3358` — Ternary operators should not be nested
- **Status:** FIX - lift the inner ternary into a named local.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-glKuZhe4vAc1NYg2W
- Fix guidance: see rule `csharpsquid:S3358` below.

## `XLibur/Excel/CalcEngine/Functions/DynamicArray.cs`

### 113. Either merge this branch with the identical one on line 729 or change one of the implementations.
- Location: `XLibur/Excel/CalcEngine/Functions/DynamicArray.cs:734`
- Rule: `csharpsquid:S1871` — Two branches in a conditional structure should not have exactly the same implementation
- **Status:** IGNORE - accepted: the two arms are the two XLOOKUP match modes; merging the conditions hides that.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-POmZkxC1_kEbSU1hk
- Fix guidance: see rule `csharpsquid:S1871` below.

### 114. Change type of parameter 'rowOrder' from 'System.Collections.Generic.IReadOnlyList<int>' to 'System.Collections.Generic.List<int>' for improved performance
- Location: `XLibur/Excel/CalcEngine/Functions/DynamicArray.cs:771`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-POmZkxC1_kEbSU1hl
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur/Excel/CalcEngine/Functions/Financial.cs`

### 115. Method has 8 parameters, which is greater than the 7 authorized.
- Location: `XLibur/Excel/CalcEngine/Functions/Financial.cs:374`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-eopULnUgmSfWQwUdz
- Fix guidance: see rule `csharpsquid:S107` below.

### 116. Remove this unused method parameter 'numberOfPayments'.
- Location: `XLibur/Excel/CalcEngine/Functions/Financial.cs:545`
- Rule: `csharpsquid:S1172` — Unused method parameters should be removed
- **Status:** FIX - drop the unused parameter (private method, no API impact).
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-eopULnUgmSfWQwUd0
- Fix guidance: see rule `csharpsquid:S1172` below.

## `XLibur/Excel/CalcEngine/Functions/Regression.cs`

### 117. Merge this if statement with the enclosing one.
- Location: `XLibur/Excel/CalcEngine/Functions/Regression.cs:896`
- Rule: `csharpsquid:S1066` — Mergeable "if" statements should be combined
- **Status:** FIX - merge the guard into the enclosing if.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-guvS2opxAhajzJgSM
- Fix guidance: see rule `csharpsquid:S1066` below.

## `XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs`

### 118. Parameters to 'Invoke' have the same names but not the same order as the method arguments.
- Location: `XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs:1019`
- Rule: `csharpsquid:S2234` — Arguments should be passed in the same order as the method parameters
- **Status:** IGNORE - FALSE POSITIVE: Func<> names its generics arg1..argN, so locals arg0..argN-1 look transposed. The call is positionally correct.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-glKwihe4vAc1NYg2Y
- Fix guidance: see rule `csharpsquid:S2234` below.

### 119. Parameters to 'Invoke' have the same names but not the same order as the method arguments.
- Location: `XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs:1063`
- Rule: `csharpsquid:S2234` — Arguments should be passed in the same order as the method parameters
- **Status:** IGNORE - FALSE POSITIVE: Func<> names its generics arg1..argN, so locals arg0..argN-1 look transposed. The call is positionally correct.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-glKwihe4vAc1NYg2Z
- Fix guidance: see rule `csharpsquid:S2234` below.

### 120. Parameters to 'Invoke' have the same names but not the same order as the method arguments.
- Location: `XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs:1094`
- Rule: `csharpsquid:S2234` — Arguments should be passed in the same order as the method parameters
- **Status:** IGNORE - FALSE POSITIVE: Func<> names its generics arg1..argN, so locals arg0..argN-1 look transposed. The call is positionally correct.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-glKwihe4vAc1NYg2a
- Fix guidance: see rule `csharpsquid:S2234` below.

## `XLibur/Excel/CalcEngine/Functions/Statistical.cs`

### 121. Extract this nested ternary operation into an independent statement.
- Location: `XLibur/Excel/CalcEngine/Functions/Statistical.cs:852`
- Rule: `csharpsquid:S3358` — Ternary operators should not be nested
- **Status:** FIX - lift the inner ternary into a named local.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-glKqfhe4vAc1NYg2U
- Fix guidance: see rule `csharpsquid:S3358` below.

## `XLibur/Excel/CalcEngine/Functions/Text.cs`

### 122. Do not update the stop condition variable 'i' in the body of the for loop.
- Location: `XLibur/Excel/CalcEngine/Functions/Text.cs:195`
- Rule: `csharpsquid:S127` — "for" loop stop conditions should be invariant
- **Status:** IGNORE - accepted: the loop variable is advanced deliberately while scanning a token run.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-eqn1UBaRu4znB71w1
- Fix guidance: see rule `csharpsquid:S127` below.

## `XLibur/Excel/CalcEngine/NumberParser.cs`

### 123. Use 'string.Contains' instead of 'string.IndexOf' to improve readability
- Location: `XLibur/Excel/CalcEngine/NumberParser.cs:109`
- Rule: `external_roslyn:CA2249` — roslyn:CA2249
- **Status:** FIX - string.Contains.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-dnDiBTAUVQ8sYxMgn
- Fix guidance: see rule `external_roslyn:CA2249` below.

## `XLibur/Excel/CalcEngine/ScalarValue.cs`

### 124. Use span-based 'string.Concat' and 'AsSpan' instead of 'Substring'
- Location: `XLibur/Excel/CalcEngine/ScalarValue.cs:315`
- Rule: `external_roslyn:CA1845` — roslyn:CA1845
- **Status:** FIX - span-based Concat.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-dnDmHTAUVQ8sYxMgp
- Fix guidance: see rule `external_roslyn:CA1845` below.

## `XLibur/Excel/Cells/XLCellFormulaShifter.cs`

### 125. Remove this commented out code.
- Location: `XLibur/Excel/Cells/XLCellFormulaShifter.cs:212`
- **Status:** IGNORE - FALSE POSITIVE on inspection: the flagged lines are prose describing how row deletion clamps a reference, not commented-out code.
- **Status:** FIX - delete the commented-out code.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-nMDZv--pWpbX2ie1c
- Fix guidance: see rule `csharpsquid:S125` below.

### 126. Extract this nested ternary operation into an independent statement.
- Location: `XLibur/Excel/Cells/XLCellFormulaShifter.cs:293`
- Rule: `csharpsquid:S3358` — Ternary operators should not be nested
- **Status:** FIX - lift the inner ternary into a named local.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-nMDZv--pWpbX2ie1e
- Fix guidance: see rule `csharpsquid:S3358` below.

## `XLibur/Excel/Charts/XLChartAxis.cs`

### 127. Method has 11 parameters, which is greater than the 7 authorized.
- Location: `XLibur/Excel/Charts/XLChartAxis.cs:179`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-cV-r7RVSYGxsIdQkZ
- Fix guidance: see rule `csharpsquid:S107` below.

## `XLibur/Excel/Charts/XLChartDataLabels.cs`

### 128. Change return type of method 'AllowedPositions' from 'System.Collections.Generic.IReadOnlyList<XLibur.Excel.XLDataLabelPosition>' to 'XLibur.Excel.XLDataLabelPosition[]' for improved performance
- Location: `XLibur/Excel/Charts/XLChartDataLabels.cs:172`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-cSntRBaRu4znBtB8r
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur/Excel/Charts/XLChartSeries.cs`

### 129. Method has 8 parameters, which is greater than the 7 authorized.
- Location: `XLibur/Excel/Charts/XLChartSeries.cs:174`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-cMSPfZiQzo-vEaiDN
- Fix guidance: see rule `csharpsquid:S107` below.

## `XLibur/Excel/Comments/Threaded/XLPersons.cs`

### 130. Change type of parameter 'left' from 'XLibur.Excel.IXLPerson' to 'XLibur.Excel.XLPerson' for improved performance
- Location: `XLibur/Excel/Comments/Threaded/XLPersons.cs:105`
- Rule: `external_roslyn:CA1859` — roslyn:CA1859
- **Status:** FIX - tighten to the concrete type where the member is not public API.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-gohwNyIH_nJ2fTUr0
- Fix guidance: see rule `external_roslyn:CA1859` below.

## `XLibur/Excel/Drawings/XLPictures.cs`

### 131. The parameter name 'pictures' is not declared in the argument list.
- Location: `XLibur/Excel/Drawings/XLPictures.cs:253`
- Rule: `csharpsquid:S3928` — Parameter names used into ArgumentException constructors should match an existing one
- **Status:** FIX - real defect, paired with CA2208 on the same line.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7lCA5wahmNBK0S5Dqf
- Fix guidance: see rule `csharpsquid:S3928` below.

### 132. Method ValidateMembersAndComputeBounds passes 'pictures' as the paramName argument to a ArgumentException constructor. Replace this argument with one of the method's parameter names. Note that the provided parameter name should have the exact casing as declared on the method.
- Location: `XLibur/Excel/Drawings/XLPictures.cs:253`
- Rule: `external_roslyn:CA2208` — roslyn:CA2208
- **Status:** FIX - real defect: paramName does not match a parameter of the throwing method.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7lCA5wahmNBK0S5Dqh
- Fix guidance: see rule `external_roslyn:CA2208` below.

### 133. The parameter name 'pictures' is not declared in the argument list.
- Location: `XLibur/Excel/Drawings/XLPictures.cs:255`
- Rule: `csharpsquid:S3928` — Parameter names used into ArgumentException constructors should match an existing one
- **Status:** FIX - real defect, paired with CA2208 on the same line.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7lCA5wahmNBK0S5Dqg
- Fix guidance: see rule `csharpsquid:S3928` below.

### 134. Method ValidateMembersAndComputeBounds passes 'pictures' as the paramName argument to a ArgumentException constructor. Replace this argument with one of the method's parameter names. Note that the provided parameter name should have the exact casing as declared on the method.
- Location: `XLibur/Excel/Drawings/XLPictures.cs:255`
- Rule: `external_roslyn:CA2208` — roslyn:CA2208
- **Status:** FIX - real defect: paramName does not match a parameter of the throwing method.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7lCA5wahmNBK0S5Dqi
- Fix guidance: see rule `external_roslyn:CA2208` below.

## `XLibur/Excel/IO/DrawingPartReader.cs`

### 135. Method has 9 parameters, which is greater than the 7 authorized.
- Location: `XLibur/Excel/IO/DrawingPartReader.cs:154`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7aPVYS5NXmL4P0hn9k
- Fix guidance: see rule `csharpsquid:S107` below.

## `XLibur/Excel/IO/WorksheetSheetDataReader.cs`

### 136. Constructor has 8 parameters, which is greater than the 7 authorized.
- Location: `XLibur/Excel/IO/WorksheetSheetDataReader.cs:27`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-SS8zGaG0jWV9Gpqg9
- Fix guidance: see rule `csharpsquid:S107` below.

### 137. LoadRowXml calls TryParseOoxmlNonNegativeInt but does not explicitly check whether the conversion succeeded. Either use the return value in a conditional statement or verify that the call site expects that the out argument will be set to the default value when the conversion fails.
- Location: `XLibur/Excel/IO/WorksheetSheetDataReader.cs:200`
- Rule: `external_roslyn:CA1806` — roslyn:CA1806
- **Status:** FIX - make the tolerated-parse fallback explicit rather than dropping the bool.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-XqigM18-5KNpaldmM
- Fix guidance: see rule `external_roslyn:CA1806` below.

### 138. LoadCellXml calls TryParseOoxmlNonNegativeInt but does not explicitly check whether the conversion succeeded. Either use the return value in a conditional statement or verify that the call site expects that the out argument will be set to the default value when the conversion fails.
- Location: `XLibur/Excel/IO/WorksheetSheetDataReader.cs:303`
- Rule: `external_roslyn:CA1806` — roslyn:CA1806
- **Status:** FIX - make the tolerated-parse fallback explicit rather than dropping the bool.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-XqigM18-5KNpaldmN
- Fix guidance: see rule `external_roslyn:CA1806` below.

### 139. Method has 8 parameters, which is greater than the 7 authorized.
- Location: `XLibur/Excel/IO/WorksheetSheetDataReader.cs:371`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-SS8zGaG0jWV9Gpqg6
- Fix guidance: see rule `csharpsquid:S107` below.

### 140. Method has 9 parameters, which is greater than the 7 authorized.
- Location: `XLibur/Excel/IO/WorksheetSheetDataReader.cs:512`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-SS8zGaG0jWV9Gpqg8
- Fix guidance: see rule `csharpsquid:S107` below.

### 141. Method has 9 parameters, which is greater than the 7 authorized.
- Location: `XLibur/Excel/IO/WorksheetSheetDataReader.cs:707`
- Rule: `csharpsquid:S107` — Methods should not have too many parameters
- **Status:** IGNORE - accepted: OOXML readers/writers carry one attribute per parameter; a parameter object would only rename the problem. Suppress per method, matching the S3776 convention.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-XqigM18-5KNpaldmL
- Fix guidance: see rule `csharpsquid:S107` below.

## `XLibur/Excel/PageSetup/XLHeaderFooter.cs`

### 142. Remove this commented out code.
- Location: `XLibur/Excel/PageSetup/XLHeaderFooter.cs:38`
- Rule: `csharpsquid:S125` — Sections of code should not be commented out
- **Status:** FIX - delete the commented-out code.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZzs6xEUnvODKcpvTk5f
- Fix guidance: see rule `csharpsquid:S125` below.

## `XLibur/Excel/PageSetup/XLPageSetup.cs`

### 143. Remove this commented out code.
- Location: `XLibur/Excel/PageSetup/XLPageSetup.cs:202`
- Rule: `csharpsquid:S125` — Sections of code should not be commented out
- **Status:** FIX - delete the commented-out code.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZzs6xEknvODKcpvTk5r
- Fix guidance: see rule `csharpsquid:S125` below.

## `XLibur/Excel/Ranges/Index/XLRangeIndex.cs`

### 144. Use the "Count" property instead of Enumerable.Count()
- Location: `XLibur/Excel/Ranges/Index/XLRangeIndex.cs:152`
- Rule: `external_roslyn:CA1829` — roslyn:CA1829
- **Status:** FIX - use the Count property.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-4RM-Sp4FixTNOqCwf
- Fix guidance: see rule `external_roslyn:CA1829` below.

### 145. Do not use Enumerable methods on indexable collections. Instead use the collection directly.
- Location: `XLibur/Excel/Ranges/Index/XLRangeIndex.cs:152`
- Rule: `external_roslyn:CA1826` — roslyn:CA1826
- **Status:** FIX - index the collection directly.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-4RM-Sp4FixTNOqCwg
- Fix guidance: see rule `external_roslyn:CA1826` below.

## `XLibur/Excel/Streaming/XLStreamingWorkbook.cs`

### 146. Member 'CreateStyle' does not access instance data and can be marked as static
- Location: `XLibur/Excel/Streaming/XLStreamingWorkbook.cs:157`
- Rule: `external_roslyn:CA1822` — roslyn:CA1822
- **Status:** IGNORE - reversed on inspection. `XLStreamingWorkbook.CreateStyle` is a public method on a public class and is not declared by any interface, so making it static is a source-breaking change for callers. Not worth it to devirtualise one call.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-jJ3bhmmYrCxinzLl2
- Fix guidance: see rule `external_roslyn:CA1822` below.

## `XLibur/Excel/XLWorkbook_Load.cs`

### 147. Use 'Guid.NewGuid()' or 'Guid.Empty' or add arguments to this GUID instantiation.
- Location: `XLibur/Excel/XLWorkbook_Load.cs:884`
- Rule: `csharpsquid:S4581` — "new Guid()" should not be used
- **Status:** FIX - verify and replace the parameterless Guid construction.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-goh0RyIH_nJ2fTUr2
- Fix guidance: see rule `csharpsquid:S4581` below.

## `XLibur/Excel/XLWorksheet.cs`

### 148. Extract this nested ternary operation into an independent statement.
- Location: `XLibur/Excel/XLWorksheet.cs:1552`
- Rule: `csharpsquid:S3358` — Ternary operators should not be nested
- **Status:** FIX - lift the inner ternary into a named local.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7yQVEmFSxzLqliAN63
- Fix guidance: see rule `csharpsquid:S3358` below.

### 149. Extract this nested ternary operation into an independent statement.
- Location: `XLibur/Excel/XLWorksheet.cs:1553`
- Rule: `csharpsquid:S3358` — Ternary operators should not be nested
- **Status:** FIX - lift the inner ternary into a named local.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ7yQVEmFSxzLqliAN64
- Fix guidance: see rule `csharpsquid:S3358` below.

## `XLibur/XLHelper.cs`

### 150. Prefer 'AsSpan' over 'Substring' when span-based overloads are available
- Location: `XLibur/XLHelper.cs:283`
- Rule: `external_roslyn:CA1846` — roslyn:CA1846
- **Status:** FIX - AsSpan instead of Substring.
- Link: https://sonarcloud.io/project/issues?id=XLibur_XLibur&open=AZ-NRIkbQ38ITsHQN8tJ
- Fix guidance: see rule `external_roslyn:CA1846` below.

---

## How to fix — rule guidance

### `csharpsquid:S1066` — Mergeable "if" statements should be combined

**Why this is an issue**

Nested code - blocks of code inside blocks of code - is eventually necessary, but increases complexity. This is why keeping the code as flat as
possible, by avoiding unnecessary nesting, is considered a good practice.

Merging `if` statements when possible will decrease the nesting of the code and improve its readability.

Code like

```
if (condition1)
{
    if (condition2)           // Noncompliant
    {
        // ...
    }
}
```

Will be more readable as

```
if (condition1 && condition2) // Compliant
{
    // ...
}
```

**How to fix it**

If merging the conditions seems to result in a more complex code, extracting the condition or part of it in a named function or variable is a
better approach to fix readability.

#### Noncompliant code example

```
if (file != null)
{
  if (file.isFile() || file.isDirectory())    // Noncompliant
  {
    /* ... */
  }
}
```

#### Compliant solution

```
bool isFileOrDirectory(File file)
{
  return file.isFile() || file.isDirectory();
}

/* ... */

if (file != null && isFileOrDirectory(file))  // Compliant
{
  /* ... */
}
```

### `csharpsquid:S107` — Methods should not have too many parameters

**Why this is an issue**

Methods with a long parameter list are difficult to use because maintainers must figure out the role of each parameter and keep track of their
position.

```
void SetCoordinates(int x1, int y1, int z1, int x2, int y2, int z2) // Noncompliant
{
   // ...
}
```

The solution can be to:

- Split the method into smaller ones

```
// Each function does a part of what the original setCoordinates function was doing, so confusion risks are lower
void SetOrigin(int x, int y, int z)
{
   // ...
}

void SetSize(int width, int height, int depth)
{
   //
}
```

- Find a better data structure for the parameters that group data in a way that makes sense for the specific application domain

```
// In geometry, Point is a logical structure to group data
readonly record struct Point(int X, int Y, int Z);

void SetCoordinates(Point p1, Point p2)
{
    // ...
}
```

This rule raises an issue when a method has more parameters than the provided threshold.

#### Exceptions

The rule does not count the parameters intended for a base class constructor.

With a maximum number of 4 parameters:

```
public class BaseClass
{
    public BaseClass(int param1)
    {
        // ...
    }
}

public class DerivedClass : BaseClass
{
    public DerivedClass(int param1, int param2, int param3, string param4, long param5) : base(param1) // Compliant by exception
    {
        // ...
    }
}
```

### `csharpsquid:S1118` — Utility classes should not have public constructors

**Why this is an issue**

Whenever there are portions of code that are duplicated and do not depend on the state of their container class, they can be centralized inside a
"utility class". A utility class is a class that only has static members, hence it should not be instantiated.

**How to fix it**

To prevent the class from being instantiated, you should define a non-public constructor. This will prevent the compiler from implicitly generating
a public parameterless constructor.

Alternatively, adding the `static` keyword as class modifier will also prevent it from being instantiated.

#### Noncompliant code example

```
public class StringUtils // Noncompliant: implicit public constructor
{
  public static string Concatenate(string s1, string s2)
  {
    return s1 + s2;
  }
}
```

or

```
public class StringUtils // Noncompliant: explicit public constructor
{
  public StringUtils()
  {
  }

  public static string Concatenate(string s1, string s2)
  {
    return s1 + s2;
  }
}
```

#### Compliant solution

```
public static class StringUtils // Compliant: the class is static
{
  public static string Concatenate(string s1, string s2)
  {
    return s1 + s2;
  }
}
```

or

```
public class StringUtils // Compliant: the constructor is not public
{
  private StringUtils()
  {
  }

  public static string Concatenate(string s1, string s2)
  {
    return s1 + s2;
  }
}
```

### `csharpsquid:S1172` — Unused method parameters should be removed

**Why this is an issue**

A typical code smell known as unused function parameters refers to parameters declared in a function but not used anywhere within the function’s
body. While this might seem harmless at first glance, it can lead to confusion and potential errors in your code. Disregarding the values passed to
such parameters, the function’s behavior will be the same, but the programmer’s intention won’t be clearly expressed anymore. Therefore, removing
function parameters that are not being utilized is considered best practice.

This rule raises an issue when a `private` method or constructor of a class/struct takes a parameter without using it.

#### Exceptions

This rule doesn’t raise any issue in the following contexts:

- The `this` parameter of extension methods.

- Methods decorated with attributes.

- Empty methods.

- Methods which only throw `NotImplementedException`.

- The Main method of the application.

- `virtual`, `override` methods.

- interface implementations.

**How to fix it**

Having unused function parameters in your code can lead to confusion and misunderstanding of a developer’s intention. They reduce code readability
and introduce the potential for errors. To avoid these problems, developers should remove unused parameters from function declarations.

#### Noncompliant code example

```
private void DoSomething(int a, int b) // Noncompliant, "b" is unused
{
    Compute(a);
}

private void DoSomething2(int a) // Noncompliant, the value of "a" is unused
{
    a = 10;
    Compute(a);
}
```

#### Compliant solution

```
private void DoSomething(int a)
{
    Compute(a);
}

private void DoSomething2()
{
    var a = 10;
    Compute(a);
}
```

### `csharpsquid:S125` — Sections of code should not be commented out

**Why this is an issue**

Commented-out code distracts the focus from the actual executed code. It creates a noise that increases maintenance code. And because it is never
executed, it quickly becomes out of date and invalid.

Commented-out code should be deleted and can be retrieved from source control history if required.

**How to fix it**

Delete the commented out code.

#### Noncompliant code example

```
void Method(string s)
{
    // if (s.StartsWith('A'))
    // {
    //     s = s.Substring(1);
    // }

    // Do something...
}
```

#### Compliant solution

```
void Method(string s)
{
    // Do something...
}
```

### `csharpsquid:S127` — "for" loop stop conditions should be invariant

**Why this is an issue**

A `for` loop stop condition should test the loop counter against an invariant value, one that is true at both the beginning and ending
of every loop iteration. Ideally, this means that the stop condition is set to a local variable just before the loop begins.

This rule tracks when incremented counters used in the stop condition are updated in the body of the `for` loop.

#### What is the potential impact?

Non-invariant stop conditions can lead to unexpected loop behavior, making the code harder to debug and maintain. If the stop condition changes
unexpectedly during iteration, it may cause:

- infinite loops or premature loop termination

- off-by-one errors that are difficult to trace

- subtle bugs that only manifest under specific conditions

**How to fix it**

It is generally recommended to only update the loop counter in the loop declaration. If skipping elements or iterating at a different pace based on
a condition is needed, consider using a while loop or a different structure that better fits the needs.

#### Noncompliant code example

```
for (int i = 1; i <= 5; i++)
{
    Console.WriteLine(i);
    if (condition)
    {
        i = 20;
    }
}
```

#### Compliant solution

```
int i = 1;
while (i <= 5)
{
    Console.WriteLine(i);
    if (condition)
    {
        i = 20;
    }
    else
    {
        i++;
    }
}
```

#### How does this work?

A `while` loop signals that the iteration logic may be more complex, so readers will naturally look for control flow changes within the
loop body. This makes the code’s intent clearer and easier to reason about.

### `csharpsquid:S1871` — Two branches in a conditional structure should not have exactly the same implementation

**Why this is an issue**

When the same code is duplicated in two or more separate branches of a conditional, it can make the code harder to understand, maintain, and can
potentially introduce bugs if one instance of the code is changed but others are not.

Having two `cases` in a `switch` statement or two branches in an `if` chain with the same implementation is at
best duplicate code, and at worst a coding error.

```
if (a >= 0 && a < 10)
{
  DoFirst();
  DoTheThing();
}
else if (a >= 10 && a < 20)
{
  DoTheOtherThing();
}
else if (a >= 20 && a < 50) // Noncompliant; duplicates first condition
{
  DoFirst();
  DoTheThing();
}
```

```
switch (i)
{
  case 1:
    DoFirst();
    DoSomething();
    break;
  case 2:
    DoSomethingDifferent();
    break;
  case 3:  // Noncompliant; duplicates case 1's implementation
    DoFirst();
    DoSomething();
    break;
  default:
    DoTheRest();
}
```

If the same logic is truly needed for both instances, then:

- in an `if` chain they should be combined

```
if ((a >= 0 && a < 10) || (a >= 20 && a < 50))
{
  DoFirst();
  DoTheThing();
}
else if (a >= 10 && a < 20)
{
  DoTheOtherThing();
}
```

- for a `switch`, one should fall through to the other

```
switch (i)
{
  case 1:
  case 3:
    DoFirst();
    DoSomething();
    break;
  case 2:
    DoSomethingDifferent();
    break;
  default:
    DoTheRest();
}
```

#### Exceptions

The rule does not raise an issue for blocks in an `if` chain that contain a single line of code. The same applies to blocks in a
`switch` statement that contain a single line of code with or without a following `break`.

```
if (a >= 0 && a < 10)
{
  DoTheThing();
}
else if (a >= 10 && a < 20)
{
  DoTheOtherThing();
}
else if (a >= 20 && a < 50)    //no issue, usually this is done on purpose to increase the readability
{
  DoTheThing();
}
```

However, this exception does not apply to `if` chains without an `else` statement or to a `switch` statement
without a `default` clause when all branches have the same single line of code.

```
if (a == 1)
{
  DoSomething();  // Noncompliant, this might have been done on purpose but probably not
}
else if (a == 2)
{
  DoSomething();
}
```

### `csharpsquid:S2234` — Arguments should be passed in the same order as the method parameters

**Why this is an issue**

Calling a method with argument variables whose names match the method parameter names but in a different order can cause confusion. It could
indicate a mistake in the arguments' order, leading to unexpected results.

```
public double Divide(int divisor, int dividend)
{
    return divisor / dividend;
}

public void DoTheThing()
{
    int divisor = 15;
    int dividend = 5;

    double result = Divide(dividend, divisor);  // Noncompliant: arguments' order doesn't match their respective parameter names
    // ...
}
```

However, matching the method parameters' order contributes to clearer and more readable code:

```
public double Divide(int divisor, int dividend)
{
    return divisor / dividend;
}

public void DoTheThing()
{
    int divisor = 15;
    int dividend = 5;

    double result = Divide(divisor, dividend); // Compliant
    // ...
}
```

### `csharpsquid:S3358` — Ternary operators should not be nested

**Why this is an issue**

Nested ternary operators are hard to read because the mapping between each condition and its result is not immediately obvious, and the order in
which the conditions are evaluated is easy to misjudge. Each extra level of nesting increases the effort required to understand which branch produces
which value.

```
public string GetAttendanceStatus(Student student)
{
    return student.IsPresent ? "Present" : student.HasExcuse ? "Excused" : "Absent";  // Noncompliant
}
```

#### Exceptions

The rule does not flag nested ternary operations inside a lambda expression that is converted to an expression tree
(`Expression<TDelegate>`), such as the lambdas passed to `IQueryable<T>` methods (`Select`,
`Where`, `OrderBy`, and similar) when using Entity Framework Core, LINQ to SQL, or another query provider.

None of the usual fixes are available there: a lambda converted to an expression tree cannot have a statement body, so the ternary operation cannot
be rewritten into an `if` statement (compiler error CS0834). A `switch` expression is not an option either, because expression
trees cannot contain one at all (compiler error CS8514). Extracting the logic into a separate method compiles, but most query providers cannot
translate a call to a user-defined method, so the query fails at runtime instead. The nested ternary operation is therefore usually the only construct
that both compiles and translates reliably, typically into a SQL `CASE WHEN`.

```
IQueryable<string> query = context.Students
    .Select(s => s.IsGraduated ? "Graduated" : s.IsEnrolled ? "Enrolled" : "Withdrawn"); // Compliant: expression tree, translated to a SQL CASE WHEN
```

A `switch` expression or an extracted method can still be used here, but only after the query has been materialized, for example after
`AsEnumerable()` or `ToList()`.

**How to fix it**

Extract the nested ternary operation into a separate statement or expression, so that each condition and its result can be read on their own.
Depending on the situation, this can be achieved by:

- extracting the outer condition into an `if` statement

- replacing the chain of conditions with a `switch` expression

- extracting the whole expression into a dedicated, well-named method

#### Noncompliant code example

```
public string GetAttendanceStatus(Student student)
{
    return student.IsPresent ? "Present" : student.HasExcuse ? "Excused" : "Absent";  // Noncompliant
}
```

#### Compliant solution

```
public string GetAttendanceStatus(Student student)
{
    if (student.IsPresent)
    {
        return "Present";
    }
    return student.HasExcuse ? "Excused" : "Absent";
}
```

A `switch` expression is a good substitute when all conditions test the same expression:

#### Noncompliant code example

```
public string GetGradeLabel(Student student) =>
    student.Score >= 90 ? "Excellent" : student.Score >= 70 ? "Good" : "NeedsImprovement"; // Noncompliant
```

#### Compliant solution

```
public string GetGradeLabel(Student student) =>
    student.Score switch
    {
        >= 90 => "Excellent",
        >= 70 => "Good",
        _ => "Needs improvement",
    };
```

When the nested ternary is just one part of a larger statement, extracting it into a well-named method keeps the surrounding code readable:

#### Noncompliant code example

```
public void PrintEnrollmentStatus(Student student)
{
    Console.WriteLine($"Status: {(student.IsGraduated ? "Graduated" : student.IsEnrolled ? "Enrolled" : "Withdrawn")}"); // Noncompliant
}
```

#### Compliant solution

```
public void PrintEnrollmentStatus(Student student)
{
    Console.WriteLine($"Status: {GetEnrollmentStatus(student)}");
}

private static string GetEnrollmentStatus(Student student)
{
    if (student.IsGraduated)
    {
        return "Graduated";
    }
    return student.IsEnrolled ? "Enrolled" : "Withdrawn";
}
```

### `csharpsquid:S3928` — Parameter names used into ArgumentException constructors should match an existing one

**Why this is an issue**

Some constructors of the `ArgumentException`, `ArgumentNullException`, `ArgumentOutOfRangeException` and
`DuplicateWaitObjectException` classes must be fed with a valid parameter name. This rule raises an issue in two cases:

- When this parameter name doesn’t match any existing ones.

- When a call is made to the default (parameterless) constructor

#### Noncompliant code example

```
public void Foo(Bar a, int[] b)
{
  throw new ArgumentException();                                        // Noncompliant
  throw new ArgumentException("My error message", "c");                 // Noncompliant
  throw new ArgumentException("My error message", "c", innerException); // Noncompliant

  throw new ArgumentNullException("c");                     // Noncompliant
  throw new ArgumentNullException(nameof(c));               // Noncompliant
  throw new ArgumentNullException("My error message", "a"); // Noncompliant

  throw new ArgumentOutOfRangeException("c");                           // Noncompliant
  throw new ArgumentOutOfRangeException("c", "My error message");       // Noncompliant
  throw new ArgumentOutOfRangeException("c", b, "My error message");    // Noncompliant

  throw new DuplicateWaitObjectException("c", "My error message");      // Noncompliant
}
```

#### Compliant solution

```
public void Foo(Bar a, int[] b)
{
  throw new ArgumentException("My error message", "a");
  throw new ArgumentException("My error message", "b", innerException);

  throw new ArgumentNullException("a");
  throw new ArgumentNullException(nameof(a));
  throw new ArgumentNullException("a", "My error message");

  throw new ArgumentOutOfRangeException("b");
  throw new ArgumentOutOfRangeException("b", "My error message");
  throw new ArgumentOutOfRangeException("b", b, "My error message");

  throw new DuplicateWaitObjectException("b", "My error message");
}
```

#### Exceptions

The rule won’t raise an issue if the parameter name is not a constant value.

### `csharpsquid:S4581` — "new Guid()" should not be used

**Why this is an issue**

When the syntax `new Guid()` (i.e. parameterless instantiation) is used, it must be that one of three things is wanted:

- An empty GUID, in which case `Guid.Empty` is clearer.

- A randomly-generated GUID, in which case `Guid.NewGuid()` should be used.

- A new GUID with a specific initialization, in which case the initialization parameter is missing.

This rule raises an issue when a parameterless instantiation of the `Guid` struct is found.

#### Noncompliant code example

```
public void Foo()
{
    var g1 = new Guid();    // Noncompliant - what's the intent?
    Guid g2 = new();        // Noncompliant
    var g3 = default(Guid); // Noncompliant
    Guid g4 = default;      // Noncompliant
}
```

#### Compliant solution

```
public void Foo(byte[] bytes)
{
    var g1 = Guid.Empty;
    var g2 = Guid.NewGuid();
    var g3 = new Guid(bytes);
}
```

### `csharpsquid:S6562` — Always set the "DateTimeKind" when creating new "DateTime" instances

**Why this is an issue**

Not knowing the `Kind` of the `DateTime` object that an application is using can lead to misunderstandings when displaying or
comparing them. Explicitly setting the `Kind` property helps the application to stay consistent, and its maintainers understand what kind
of date is being managed. To achieve this, when instantiating a new `DateTime` object you should always use a constructor overload that
allows you to define the `Kind` property.

#### What is the potential impact?

Creating the `DateTime` object without specifying the property `Kind` will set it to the default value of
`DateTimeKind.Unspecified`. In this case, calling the method `ToUniversalTime` will assume that `Kind` is
`DateTimeKind.Local` and calling the method `ToLocalTime` will assume that it’s `DateTimeKind.Utc`. As a result, you
might have mismatched `DateTime` objects in your application.

**How to fix it**

To resolve this issue, use a constructor overload that lets you
specify the `DateTimeKind` when creating the
`DateTime` object. From .Net 6 onwards, use the `DateOnly` type if the time portion of the date is not
relevant.

#### Noncompliant code example

```
void CreateNewTime()
{
    var birthDate = new DateTime(1994, 7, 5, 16, 23, 42);
}
```

#### Compliant solution

```
void CreateNewTime()
{
    var birthDate = new DateTime(1994, 7, 5, 16, 23, 42, DateTimeKind.Utc);
    // or from .Net 6 onwards, use DateOnly:
    var birthDate = new DateOnly(1994, 7, 5);
}
```

### `external_roslyn:CA1510` — Use ArgumentNullException throw helper

Throw helpers are simpler and more efficient than an if block constructing a new exception instance.

### `external_roslyn:CA1513` — Use ObjectDisposedException throw helper

Throw helpers are simpler and more efficient than an if block constructing a new exception instance.

### `external_roslyn:CA1806` — Do not ignore method results

A new object is created but never used; or a method that creates and returns a new string is called and the new string is never used; or a COM or P/Invoke method returns an HRESULT or error code that is never used.

### `external_roslyn:CA1822` — Mark members as static

Members that do not access instance data or call instance methods can be marked as static. After you mark the methods as static, the compiler will emit nonvirtual call sites to these members. This can give you a measurable performance gain for performance-sensitive code.

### `external_roslyn:CA1826` — Do not use Enumerable methods on indexable collections

This collection is directly indexable. Going through LINQ here causes unnecessary allocations and CPU work.

### `external_roslyn:CA1829` — Use Length/Count property instead of Count() when available

Enumerable.Count() potentially enumerates the sequence while a Length/Count property is a direct access.

### `external_roslyn:CA1845` — Use span-based 'string.Concat'

It is more efficient to use 'AsSpan' and 'string.Concat', instead of 'Substring' and a concatenation operator.

### `external_roslyn:CA1846` — Prefer 'AsSpan' over 'Substring'

'AsSpan' is more efficient than 'Substring'. 'Substring' performs an O(n) string copy, while 'AsSpan' does not and has a constant cost.

### `external_roslyn:CA1859` — Use concrete types when possible for improved performance

Using concrete types avoids virtual or interface call overhead and enables inlining.

### `external_roslyn:CA1861` — Avoid constant arrays as arguments

Constant arrays passed as arguments are not reused when called repeatedly, which implies a new array is created each time. Consider extracting them to 'static readonly' fields to improve performance if the passed array is not mutated within the called method.

### `external_roslyn:CA2208` — Instantiate argument exceptions correctly

A call is made to the default (parameterless) constructor of an exception type that is or derives from ArgumentException, or an incorrect string argument is passed to a parameterized constructor of an exception type that is or derives from ArgumentException.

### `external_roslyn:CA2249` — Consider using 'string.Contains' instead of 'string.IndexOf'

Calls to 'string.IndexOf' where the result is used to check for the presence/absence of a substring can be replaced by 'string.Contains'.

### `external_roslyn:TUnit0057` — Hook context parameter available

Hook methods can accept a context parameter for additional test information. Consider adding a parameter of type {0} to access context details.

### `external_roslyn:TUnitAssertions0016` — Collection `.IsEqualTo(...)` compares by reference

`.IsEqualTo(...)` on a collection uses reference equality because collection types don't override `Equals`. Use `.IsEquivalentTo(...)` to compare contents.

