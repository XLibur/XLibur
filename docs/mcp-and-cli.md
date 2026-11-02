# The `xlibur` CLI and the XLibur MCP server

Two tools put XLibur's engine in front of a shell and an AI agent. They read a workbook, query it,
edit it, recalculate it with XLibur's formula engine, and return the computed values. Both are thin
hosts over one library, `XLibur.Operations`. An operation is a typed request and a typed result, so
the CLI and the server have the same operations, the same errors and the same path rules.

| Project | What it is |
|---|---|
| `src/XLibur.Operations` | The operations, the error model and the path policy. Not packed on its own. |
| `src/XLibur.Cli` | The `xlibur` command, a .NET tool. |
| `src/XLibur.Mcp` | `xlibur-mcp`, an MCP server over stdio, a .NET tool. |
| `src/XLibur.Operations.Tests` | The tests of all three. |

Both tools target .NET 10.

## Install

From a clone of the repository:

```
dotnet pack src/XLibur.Cli -c Release -o artifacts
dotnet pack src/XLibur.Mcp -c Release -o artifacts
dotnet tool install -g XLibur.Cli --add-source artifacts --prerelease
dotnet tool install -g XLibur.Mcp --add-source artifacts --prerelease
```

The packages are not published to NuGet yet. To run without installing, use
`dotnet run --project src/XLibur.Cli -- <command> ...`.

## Operations

| Tool (MCP) | Command (CLI) | Writes | What it does |
|---|---|---|---|
| `inspect` | `inspect` | no | Sheets with used ranges and cell and formula counts, tables, pivot tables, charts, pictures, defined names. |
| `read_range` | `read-range` | no | Values, formulas or both, as TSV or JSON, paged. |
| `find` | `find` | no | Cells by text, regular expression or error value. |
| `check_formula` | `check-formula` | no | Parses a formula and evaluates it against a workbook, writing nothing. |
| `recalc` | `recalc` | no | Recalculates every formula and reports error values, circular references, refused formulas, unsupported features, defects, and cells that differ from their cached value. |
| `validate` | `validate` | no | Load check, recalculation report, and a save-and-reload round trip compared cell by cell. |
| `diff` | `diff` | no | Compares two workbooks: sheets, cell values, formulas and styles, defined names, tables. |
| `edit` | `edit` | yes | Sets values or formulas or clears cells, all or none, recalculates, and saves. |
| `convert` | `convert` | yes | `.xlsx`/`.xlsm` to and from `.csv`, `.tsv` and `.json`. |
| `render_report` | `render-report` | yes | An `XLibur.Report` template plus a JSON data file to a generated workbook. |

Every request property is an MCP argument in camelCase and a CLI option in kebab-case.
`xlibur <command> --help` lists them with their descriptions.

### Output

The default output is compact text. `read_range` text is TSV: a header line of column letters, then
one line per row headed by its row number, with trailing blank cells dropped. A formula cell in
`both` mode reads `=B2*C2 -> 4.5`. Tab, newline, return and backslash in text are escaped as `\t`,
`\n`, `\r` and `\\`. A cell that fails to calculate reads `!circular`, `!refused`, `!unsupported` or
`!defect`.

Pass `--json` to the CLI, or `"format": "json"` to a tool, for the structured result. In JSON a blank
cell is `null`, and an error value, a date and a duration are tagged: `{"error":"#N/A"}`,
`{"date":"2026-01-31"}`, `{"time":"01:30:00"}`.

### Paging

`read_range` returns whole rows, at most `maxCells` cells (default 2,000, and never more than the
host's cap of 10,000). When the range has more rows, the result ends with `next:`, the address of
the remaining range. Pass it back as `range` to read the next page. A range is first cut to the
sheet's used range, so `A:C` reads only the rows that have content.

Measured on two real sheets from the test resources, against the per-cell JSON list (one object
with an address and a value per cell) that other Excel MCP servers return:

| Sheet | Mode | Non-blank cells | TSV | xlibur JSON | Per-cell JSON |
|---|---|---:|---:|---:|---:|
| `TemplateWithTableSourcePivotTables.xlsx`, `Hedge detail - Nominal` A1:Y138 | values | 1,022 | 3,745 chars | 16,186 | 31,092 |
| same | both | 1,022 | 64,943 | 88,002 | 101,289 |
| `functions-excel-2010.xlsx`, `Sheet2` A1:U47 | both | 619 | 17,601 | 22,017 | 36,691 |

The values read is 8.3 times smaller than the per-cell list. With formulas the gain is smaller,
1.6 to 2.1 times, because the formula text itself is most of the output.

### Edits

An `edit` request is a list of `{cell, value | formula | clear}`. Each is checked first: the cell
address, the sheet, and that XLibur can parse the formula. If any fails, nothing is applied and
nothing is written. Otherwise the edits are applied, every formula is read again, and the result
lists:

- every edited cell with its computed value,
- every formula cell that held a value before and now holds an error value or fails to calculate,
- the formula cells whose value the edit changed, up to `maxChanges`.

Only the formulas that depend on the edit are recalculated. The others keep the values the file
cached. The workbook is saved over the source, or to `outputPath` to leave the source untouched.

The CLI also takes edits as `--set`: `--set B2=42` sets a number, `--set "B2==SUM(A1:A9)"` sets a
formula, and `--set B2=` clears the cell.

### Report data

`render_report` reads one JSON object, and each property becomes a template variable. An object
becomes a dictionary, so `{{ item.Product }}` reads its `Product` key, and an array becomes a list a
bound range repeats over. A date is written as a cell value is, `{"date":"2026-01-31"}`. A plain
string stays a string, even when it looks like a date.

### Exit codes

| Code | Meaning |
|---|---|
| 0 | Done. `diff`: the workbooks are identical. |
| 1 | Done, with a finding: `diff` found differences, `validate` failed a check, `check-formula` refused or could not evaluate the formula, `render-report` listed template errors. |
| 2 | An error: bad input, a refused path, an unsupported feature, or a defect. |

### Errors

Every error has a kind, a stable snake_case code and a message, for example
`error[sheet_not_found] BadInput: No sheet 'Q3'. Sheets: Sales, Costs.` The kinds follow
[CONTEXT.md](../CONTEXT.md):

| Kind | Meaning |
|---|---|
| `BadInput` | The request or the file is wrong. Fix the request. |
| `Forbidden` | The path policy refused it. |
| `Unsupported` | A valid Excel construct XLibur does not handle, such as an encrypted workbook. |
| `Defect` | A bug in XLibur or in the tool. The message names the exception type. Report it. |

A cell that fails to calculate does not fail the operation. `recalc`, `edit` and `find` report it
per cell. `recalc` also reports a function XLibur does not implement as `unsupported`, because the
engine evaluates such a call to `#NAME?`, the same as a name nobody defined.

## MCP server

```
xlibur-mcp [--root <dir>]... [--allow-write] [--max-input-mb <n>]
```

The server speaks MCP over stdio. It logs to stderr only.

Claude Code:

```
claude mcp add xlibur -- xlibur-mcp --root C:\Work\Reports
```

Claude Desktop, in `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "xlibur": {
      "command": "xlibur-mcp",
      "args": ["--root", "C:\\Work\\Reports", "--allow-write"]
    }
  }
}
```

Without an installed tool, use `"command": "dotnet"` and
`"args": ["run", "--project", "<repo>/src/XLibur.Mcp", "--", "--root", "..."]`.

The read tools are annotated `readOnlyHint: true`. The write tools are annotated
`destructiveHint: true` and are listed only with `--allow-write`. Without it the server's
instructions tell the agent that writes are off. Calls run one at a time, so two edits of the same
file cannot interleave.

## The path policy

The server confines every path to its roots. The roots are the `--root` directories, or the current
directory when none is given. A relative path is relative to the first root. The CLI uses the same
policy with no roots by default, since it runs in the user's own shell, and `--root` turns
confinement on.

A path is refused before the file system is touched when it:

- is a device or network path: `\\server\share`, `//server`, `\\?\`, `\\.\`, `\??\`,
- holds a `:` other than a drive letter's, which names an alternate data stream (`book.xlsx:hidden`),
- has a component that ends in a dot or a space, which Windows strips (`book.xlsx.`),
- uses a reserved device name in any component (`CON`, `NUL`, `AUX`, `PRN`, `COM0`–`COM9`,
  `LPT0`–`LPT9` and their superscript forms), with any extension,
- has an extension other than `.xlsx`, `.xlsm`, `.csv`, `.tsv` or `.json`,
- holds a NUL character, or has more than 1,024 characters or 64 components,
- is not under a root by name once made absolute, with `..` resolved. A root counts by the name it
  was given and by its canonical name.

Nothing outside the roots is probed, so a refusal costs no network round trip, even for a mapped
network drive, and does not tell the caller what exists outside.

Then every symbolic link and junction in the path is followed, one component at a time, with 8.3
short names expanded. A link to a network or device path, or to any path not under a root by name,
is refused without being followed.

Last, the file is opened, and the operating system is asked where the open handle is
(`GetFinalPathNameByHandle` on Windows, `/proc/self/fd` on Linux). That answer decides containment,
so a link swapped in after the walk is still caught. It is compared with the canonical roots
exactly, not ignoring case. Both carry the names as stored on disk, and ignoring case would let a
sibling `data` pass for a root `Data` in a case-sensitive directory, or `dataµ` for `dataΜ`, which
.NET folds together and NTFS keeps apart. The extension is checked again on the final path, so a
`data.json` link to `service.env` is refused. The bytes read are the bytes from that handle, so the
file checked is the file loaded. A root is canonicalised the same way when the server starts, so a
root given through a link or a junction is its target.

Writes:

- are off unless the server is started with `--allow-write`,
- go only inside a root, with one of the allowed extensions,
- never replace a file unless the request sets `overwrite`, except `edit` saving over its own
  source,
- refuse a destination that is itself a link,
- are atomic. The workbook is written to a temporary file in the destination's directory, the
  temporary file's handle is checked against the roots, and the file is renamed over the
  destination when it is complete. A failed save leaves the destination as it was.

Limits:

| Limit | Default |
|---|---|
| Input file size | 64 MiB (`--max-input-mb`) |
| Bytes all entries of one package inflate to | 1 GiB |
| Inflation ratio of an entry over 1 MiB | 1,000 |
| Entries in one package | 10,000 |
| Cells in one `read_range` page | 10,000 |
| Characters in one result | 200,000 for the server, 1,000,000 for the CLI |

The package limits are read from the zip directory before XLibur parses anything, so a zip bomb is
refused at the cost of reading the directory. A file can understate its sizes, but
System.IO.Compression, which XLibur and the Open XML SDK read through, stops each entry at its
declared size. A test pins that behaviour. An encrypted workbook or a legacy `.xls` is refused as
unsupported. XLibur already reads XML with DTD processing prohibited.

What the policy does not cover:

- A hard link inside a root to a file outside it reads that file, because a hard link is the file
  itself. Creating one needs write access to the root.
- A directory swapped for a link between the check of a write's directory and the rename can
  redirect that write. The temporary file's handle is checked, but the rename is not atomic with
  the check.
- On Linux and macOS, a FIFO inside a root that is named like a workbook blocks the open, and with
  it the server, until something writes to it.
- On macOS and the BSDs there is no query of an open handle's path here, so the last gate resolves
  the name again instead. The walk still refuses links out of the roots, but a link swapped in
  between the walk and the open is not caught.

## CLI

```
xlibur [--root <dir>]... <command> [arguments] [--json]
```

```
xlibur inspect book.xlsx
xlibur read-range book.xlsx "Sales!A1:F200" --mode both
xlibur read-range book.xlsx "Sales!A201:F5000"
xlibur find book.xlsx --errors
xlibur find book.xlsx --regex "INV-\d{4}"
xlibur check-formula "=SUMIFS(D:D,A:A,\"Apple\")" book.xlsx --sheet Sales
xlibur edit book.xlsx --set B2=42 --set "D2==B2*C2" --output-path book-edited.xlsx
xlibur recalc book.xlsx
xlibur validate book.xlsx
xlibur diff before.xlsx after.xlsx --ignore-sheets Log --ignore-ranges "Sales!H1:H9"
xlibur diff before.xlsx after.xlsx --values recalculated
xlibur convert book.xlsx sales.csv --sheet Sales
xlibur convert data.json data.xlsx
xlibur render-report template.xlsx data.json report.xlsx
```

### `diff` for XLibur's own regression tests

`diff` compares cell values (cached, or `--values recalculated`), formula text, cell styles,
defined names and tables. Styles are compared by their full style, not by the index in the file,
so a save that renumbers the style table is not a difference. Each style difference names the parts
that differ, such as `font,fill`.

To check that XLibur keeps a file through a load and a save:

```
xlibur convert original.xlsx roundtrip.xlsx
xlibur diff original.xlsx roundtrip.xlsx
```

`validate` does the same in memory. To compare XLibur's calculation with the values Excel cached,
run `xlibur recalc book.xlsx`: its "differ from cached" list is every formula whose recalculated
value is not the one in the file.
