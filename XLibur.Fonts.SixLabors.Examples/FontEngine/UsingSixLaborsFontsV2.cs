using XLibur.Excel;
using XLibur.Fonts.SixLabors;

namespace XLibur.Fonts.SixLabors.Examples.FontEngine;

/// <summary>
/// Demonstrates using SixLabors.Fonts 2.x as the font engine for text measurement.
/// Install the XLibur.Fonts.SixLabors NuGet package to use this.
/// </summary>
public class UsingSixLaborsFontsV2
{
    public static void Create(string filePath)
    {
        // Create a SixLabors.Fonts 2.x font engine with system fonts
        var fontEngine = new SixLaborsFontEngine("Microsoft Sans Serif");

        // Pass it via LoadOptions when creating a workbook
        var options = new LoadOptions { FontEngine = fontEngine };
        using var wb = new XLWorkbook(options);
        var ws = wb.Worksheets.Add("SixLabors v2 Fonts");

        ws.Cell(1, 1).Value = "Using SixLabors.Fonts 2.x";
        ws.Cell(1, 1).Style.Font.FontSize = 20;
        ws.Cell(1, 1).Style.Font.Bold = true;

        ws.Cell(3, 1).Value = "The font engine handles text measurement for:";
        ws.Cell(4, 1).Value = "  - Column auto-fit (AdjustToContents)";
        ws.Cell(5, 1).Value = "  - Row auto-fit";
        ws.Cell(6, 1).Value = "  - Glyph metrics and max digit width";

        // Auto-fit uses the custom font engine for measurement
        ws.Column(1).AdjustToContents();
        ws.Rows(1, 6).AdjustToContents();

        // You can also set a global default so all new workbooks use it
        LoadOptions.DefaultFontEngine = fontEngine;
        using var wb2 = new XLWorkbook();
        var ws2 = wb2.Worksheets.Add("Global Default");
        ws2.Cell(1, 1).Value = "This workbook also uses SixLabors.Fonts 2.x";
        ws2.Column(1).AdjustToContents();

        wb.SaveAs(filePath);
    }
}
