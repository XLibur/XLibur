using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Everything a worksheet element handler needs and cannot change. Passed by <c>in</c> so the
/// struct is not copied per element.
/// </summary>
internal readonly struct WorksheetElementContext
{
    internal required WorksheetPart Part { get; init; }
    internal required XLWorksheet Worksheet { get; init; }
    internal required StylesheetData Styles { get; init; }
    internal required LoadContext Load { get; init; }
    internal required XLWorkbook Workbook { get; init; }
}

/// <summary>
/// What one worksheet element hands to a later one. <c>sheetPr</c> produces the page-setup
/// properties that <c>pageSetup</c> consumes, so the two are ordered and this carrier is mutable.
/// </summary>
internal struct WorksheetElementState
{
    internal PageSetupProperties? PageSetupProperties;
}
