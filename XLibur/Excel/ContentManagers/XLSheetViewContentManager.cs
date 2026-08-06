using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.ContentManagers;

internal enum XLSheetViewContents
{
    Pane,
    Selection,
    PivotSelection,
    ExtensionList
}

internal sealed class XLSheetViewContentManager : XLBaseContentManager<XLSheetViewContents>
{
    /// <inheritdoc cref="XLWorksheetContentManager(Worksheet)"/>
    public XLSheetViewContentManager(SheetView sheetView)
        : base(XLSheetViewContents.ExtensionList)
    {
        foreach (var child in sheetView.ChildElements)
        {
            switch (child)
            {
                case Pane: SetElement(XLSheetViewContents.Pane, child); break;
                case Selection: SetElement(XLSheetViewContents.Selection, child); break;
                case PivotSelection: SetElement(XLSheetViewContents.PivotSelection, child); break;
                case ExtensionList: SetElement(XLSheetViewContents.ExtensionList, child); break;
            }
        }
    }
}
