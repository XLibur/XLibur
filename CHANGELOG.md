# Changelog

## Unreleased

### Performance

- **`XmlEncoder.EncodeString` fast-path**: Added a character scan that short-circuits before the `Regex` and `StringBuilder` when a string contains no characters that need encoding (the common case for plain text). For workbooks with ~50K unique shared strings this eliminates ~50K `StringBuilder` allocations, ~50K regex evaluations, and ~50K string copies on save.

- **`IXLWorksheet.SetCellValue(int row, int column, XLCellValue value)`** (new API): Sets a cell value directly on the worksheet's internal storage without allocating an intermediate `XLCell` object. For bulk data population (e.g. 50K rows x 3 columns) this eliminates ~150K object allocations that the `Cell(row, col).SetValue(...)` pattern would create.

### Upgrade Guide

#### Using `SetCellValue` for bulk writes

The existing `Cell(row, col).SetValue(value)` API continues to work and remains the correct choice when you need full cell semantics (formula clearing, merged-range checks, table header refresh). No code changes are required.

For **performance-critical bulk data population** where you are writing values into empty or freshly-created cells, you can switch to the new direct API:

```csharp
// Before (allocates an XLCell per call):
for (int row = 1; row <= 50_000; row++)
{
    ws.Cell(row, 1).SetValue(row);
    ws.Cell(row, 2).SetValue($"Item {row}");
    ws.Cell(row, 3).SetValue(row * 1.5);
}

// After (zero intermediate allocations):
for (int row = 1; row <= 50_000; row++)
{
    ws.SetCellValue(row, 1, row);
    ws.SetCellValue(row, 2, $"Item {row}");
    ws.SetCellValue(row, 3, row * 1.5);
}
```

`SetCellValue` handles date/time number format application and quote-prefix stripping, so the resulting cell content and formatting is identical for data values. The following behaviors are **not** performed by `SetCellValue` — use `Cell().SetValue()` if you need them:

| Behavior | `Cell().SetValue()` | `SetCellValue()` |
|---|---|---|
| Set value and number format | Yes | Yes |
| Clear existing formula | Yes | No |
| Check merged range (inferior cell skip) | Yes | No |
| Refresh table header fields | Yes | No |
