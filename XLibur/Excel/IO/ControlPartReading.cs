using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The plumbing <see cref="SlicerReader"/> and <see cref="TimelineReader"/> share: streaming a part
/// without attaching a DOM to it, pairing worksheet parts with loaded worksheets, and binding a
/// control cache to the pivot tables it names.
/// </summary>
/// <remarks>
/// The two readers differ in every element they read, so they stay separate. What they share is the
/// fidelity rule — never materialise a control part — and that rule is only as good as its weakest
/// copy, so it lives here once.
/// </remarks>
internal static class ControlPartReading
{
    /// <summary>
    /// Reads a part's root element without attaching it to the part.
    /// </summary>
    /// <remarks>
    /// This is the whole fidelity guarantee of the slicer and timeline readers in three lines: the
    /// part is streamed, the element that comes back is detached, and <c>part.RootElement</c> stays
    /// unmaterialised, so the SDK has nothing to write back over the original bytes when the package
    /// is saved.
    /// </remarks>
    internal static T? ReadDetached<T>(OpenXmlPart part) where T : OpenXmlElement
    {
        using var reader = new OpenXmlPartReader(part);

        // Create reads the XML declaration only, so the first Read lands on the root element.
        return reader.Read() ? reader.LoadCurrentElement() as T : null;
    }

    /// <summary>
    /// Pairs each worksheet part with the loaded worksheet it belongs to, in sheet order.
    /// </summary>
    internal static IEnumerable<(WorksheetPart Part, XLWorksheet Worksheet)> WorksheetParts(
        WorkbookPart workbookPart, Sheets sheets, XLWorksheets worksheets)
    {
        foreach (var sheet in sheets.OfType<Sheet>())
        {
            // A sheet with an empty relationship id comes from a non-Excel producer, and the
            // relationship may point at a chartsheet rather than a worksheet.
            if (string.IsNullOrEmpty(sheet.Id?.Value)
                || sheet.Name?.Value is not { } sheetName
                || workbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart worksheetPart
                || !worksheets.TryGetWorksheet(sheetName, out var worksheet))
            {
                continue;
            }

            yield return (worksheetPart, worksheet);
        }
    }

    /// <summary>
    /// Every pivot table in the workbook, keyed by name — which is how a control cache names them.
    /// </summary>
    internal static Dictionary<string, XLPivotTable> PivotTablesByName(XLWorksheets worksheets)
    {
        var pivotTables = new Dictionary<string, XLPivotTable>(XLHelper.NameComparer);
        foreach (var worksheet in worksheets)
        {
            foreach (var pivotTable in worksheet.PivotTables.Cast<XLPivotTable>())
                pivotTables[pivotTable.Name] = pivotTable;
        }

        return pivotTables;
    }

    /// <summary>
    /// Resolves the pivot table names a cache carries, and the pivot cache behind them.
    /// </summary>
    internal static void BindPivotTables(IXLPivotDependentCache cache, Dictionary<string, XLPivotTable> pivotTables)
    {
        // A cache may name several pivot tables — that is how one set of buttons drives a whole
        // dashboard — and may name one that is no longer in the workbook, which is left out rather
        // than reported as a hole in the list.
        foreach (var name in cache.PivotTableNames)
        {
            if (pivotTables.TryGetValue(name, out var pivotTable))
                cache.PivotTables.Add(pivotTable);
        }

        // A slicer's item indices point into the shared items of the pivot cache behind those pivot
        // tables, and a timeline reads its dates from it. Pivot tables sharing a control cache share
        // a pivot cache, so the first one answers for all of them.
        cache.PivotCache = cache.PivotTables.Count > 0 ? cache.PivotTables[0].PivotCache : null;
    }
}
