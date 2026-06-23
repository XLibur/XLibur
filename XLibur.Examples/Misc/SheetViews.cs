using XLibur.Excel;

namespace XLibur.Examples.Misc;

public class SheetViews : IXLExample
{
    public void Create(string filePath)
    {
        using var wb = new XLWorkbook();

        var ws = wb.AddWorksheet("ZoomScale");
        ws.FirstCell().SetValue(ws.Name);
        ws.SheetView.ZoomScale = 50;

        ws = wb.AddWorksheet("ZoomScaleNormal");
        ws.FirstCell().SetValue(ws.Name);
        ws.SheetView.ZoomScaleNormal = 70;

        ws = wb.AddWorksheet("ZoomScalePageLayoutView");
        ws.FirstCell().SetValue(ws.Name);
        ws.SheetView.ZoomScalePageLayoutView = 85;

        ws = wb.AddWorksheet("ZoomScaleSheetLayoutView");
        ws.FirstCell().SetValue(ws.Name);
        ws.SheetView.ZoomScaleSheetLayoutView = 120;

        ws = wb.AddWorksheet("ZoomScaleTooSmall");
        ws.FirstCell().SetValue(ws.Name);
        ws.SheetView.ZoomScale = 5;

        ws = wb.AddWorksheet("ZoomScaleTooBig");
        ws.FirstCell().SetValue(ws.Name);
        ws.SheetView.ZoomScale = 500;

        ws = wb.AddWorksheet("TopLeftCell");
        ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000").Address;

        // FocusCell sets the active cell AND scrolls the sheet so it is visible at the
        // top-left of the scrollable region. On a frozen sheet it drives <pane topLeftCell>;
        // on an unfrozen sheet it drives <sheetView topLeftCell>.
        ws = wb.AddWorksheet("FocusCell");
        ws.SheetView.FreezeRows(2);
        ws.Cell("A1").SetValue("Frozen header");
        ws.Cell("A500").SetValue("Focused row");
        ws.FocusCell("A500");

        wb.SaveAs(filePath);
    }
}
