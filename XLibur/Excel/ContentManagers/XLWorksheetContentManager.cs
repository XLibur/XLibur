using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.ContentManagers;

internal enum XLWorksheetContents
{
    SheetProperties = 1,
    SheetDimension = 2,
    SheetViews = 3,
    SheetFormatProperties = 4,
    Columns = 5,
    SheetData = 6,
    SheetCalculationProperties = 7,
    SheetProtection = 8,
    ProtectedRanges = 9,
    Scenarios = 10,
    AutoFilter = 11,
    SortState = 12,
    DataConsolidate = 13,
    CustomSheetViews = 14,
    MergeCells = 15,
    PhoneticProperties = 16,
    ConditionalFormatting = 17,
    DataValidations = 18,
    Hyperlinks = 19,
    PrintOptions = 20,
    PageMargins = 21,
    PageSetup = 22,
    HeaderFooter = 23,
    RowBreaks = 24,
    ColumnBreaks = 25,
    CustomProperties = 26,
    CellWatches = 27,
    IgnoredErrors = 28,
    SmartTags = 29,
    Drawing = 30,
    LegacyDrawing = 31,
    LegacyDrawingHeaderFooter = 32,
    DrawingHeaderFooter = 33,
    Picture = 34,
    OleObjects = 35,
    Controls = 36,
    AlternateContent = 37,
    WebPublishItems = 38,
    TableParts = 39,
    WorksheetExtensionList = 40
}

internal sealed class XLWorksheetContentManager : XLBaseContentManager<XLWorksheetContents>
{
    /// <remarks>
    /// One pass over the children, keeping the last element seen for each slot. This used to run
    /// <c>Elements&lt;T&gt;().LastOrDefault()</c> once per slot — thirty-nine filtered traversals
    /// of the same child list, each allocating its own iterator — which made building the manager
    /// a significant share of the per-worksheet save cost. The last-wins behaviour is preserved
    /// because a later child simply overwrites its slot.
    /// <para>
    /// <see cref="XLWorksheetContents.SmartTags"/> is deliberately not mapped, matching the
    /// original: that slot stays empty so it never becomes an insertion anchor.
    /// </para>
    /// </remarks>
    public XLWorksheetContentManager(Worksheet opWorksheet)
        : base(XLWorksheetContents.WorksheetExtensionList)
    {
        foreach (var child in opWorksheet.ChildElements)
        {
            if (SlotOf(child) is { } slot)
                SetElement(slot, child);
        }
    }

    /// <summary>
    /// The schema slot an element belongs to, or null when the element is not tracked.
    /// </summary>
    private static XLWorksheetContents? SlotOf(OpenXmlElement child) => child switch
    {
        SheetProperties => XLWorksheetContents.SheetProperties,
        SheetDimension => XLWorksheetContents.SheetDimension,
        SheetViews => XLWorksheetContents.SheetViews,
        SheetFormatProperties => XLWorksheetContents.SheetFormatProperties,
        Columns => XLWorksheetContents.Columns,
        SheetData => XLWorksheetContents.SheetData,
        SheetCalculationProperties => XLWorksheetContents.SheetCalculationProperties,
        SheetProtection => XLWorksheetContents.SheetProtection,
        ProtectedRanges => XLWorksheetContents.ProtectedRanges,
        Scenarios => XLWorksheetContents.Scenarios,
        AutoFilter => XLWorksheetContents.AutoFilter,
        SortState => XLWorksheetContents.SortState,
        DataConsolidate => XLWorksheetContents.DataConsolidate,
        CustomSheetViews => XLWorksheetContents.CustomSheetViews,
        MergeCells => XLWorksheetContents.MergeCells,
        PhoneticProperties => XLWorksheetContents.PhoneticProperties,
        ConditionalFormatting => XLWorksheetContents.ConditionalFormatting,
        DataValidations => XLWorksheetContents.DataValidations,
        Hyperlinks => XLWorksheetContents.Hyperlinks,
        PrintOptions => XLWorksheetContents.PrintOptions,
        PageMargins => XLWorksheetContents.PageMargins,
        PageSetup => XLWorksheetContents.PageSetup,
        HeaderFooter => XLWorksheetContents.HeaderFooter,
        RowBreaks => XLWorksheetContents.RowBreaks,
        ColumnBreaks => XLWorksheetContents.ColumnBreaks,
        CustomProperties => XLWorksheetContents.CustomProperties,
        CellWatches => XLWorksheetContents.CellWatches,
        IgnoredErrors => XLWorksheetContents.IgnoredErrors,
        Drawing => XLWorksheetContents.Drawing,
        LegacyDrawing => XLWorksheetContents.LegacyDrawing,
        LegacyDrawingHeaderFooter => XLWorksheetContents.LegacyDrawingHeaderFooter,
        DrawingHeaderFooter => XLWorksheetContents.DrawingHeaderFooter,
        Picture => XLWorksheetContents.Picture,
        OleObjects => XLWorksheetContents.OleObjects,
        Controls => XLWorksheetContents.Controls,
        AlternateContent => XLWorksheetContents.AlternateContent,
        WebPublishItems => XLWorksheetContents.WebPublishItems,
        TableParts => XLWorksheetContents.TableParts,
        WorksheetExtensionList => XLWorksheetContents.WorksheetExtensionList,
        _ => null,
    };
}
