using System.IO;
using XLibur.Excel;
using XLibur.Fonts.SixLabors;

namespace XLibur.Examples.FontEngine;

/// <summary>
/// Demonstrates using stream-based fonts for environments without system font access
/// (e.g., Blazor WebAssembly, serverless, Docker containers).
/// </summary>
public class UsingStreamBasedFonts : IXLExample
{
    public void Create(string filePath)
    {
        // Load a font from a file stream — no system fonts needed
        using var fontStream = File.OpenRead("path/to/MyFont.ttf");
        var fontEngine = SixLaborsFontEngine.CreateOnlyWithFonts(fontStream);

        var options = new LoadOptions { FontEngine = fontEngine };
        using var wb = new XLWorkbook(options);
        var ws = wb.Worksheets.Add("Stream Fonts");

        ws.Cell(1, 1).Value = "Created with stream-based fonts";
        ws.Cell(2, 1).Value = "No system font access required";

        ws.Column(1).AdjustToContents();

        wb.SaveAs(filePath);
    }
}
