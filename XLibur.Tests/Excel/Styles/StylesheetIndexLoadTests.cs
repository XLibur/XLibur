using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// A cell format points into the stylesheet's lists of fonts, fills and borders by index. The load
/// takes each list's children once and looks them up by position; these pin that every format still
/// finds its own entries, the last ones of long lists included.
/// </summary>
public class StylesheetIndexLoadTests
{
    private const int Cells = 300;

    private static readonly XLBorderStyleValues[] Borders =
        [XLBorderStyleValues.Thin, XLBorderStyleValues.Dotted, XLBorderStyleValues.Dashed, XLBorderStyleValues.Medium];

    [Test]
    public async Task Every_cell_format_of_a_long_stylesheet_finds_its_own_font_fill_and_border()
    {
        using var source = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            for (var i = 0; i < Cells; i++)
            {
                var cell = ws.Cell(i + 1, 1);
                cell.Value = i;
                cell.Style.Font.FontSize = 6 + i * 0.5;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(i % 256, 255 - i % 256, i / 256);
                cell.Style.Border.LeftBorder = Borders[i % Borders.Length];
                cell.Style.Border.LeftBorderColor = XLColor.FromArgb(i % 256, 0, 0);
            }

            wb.SaveAs(source);
        }

        using var loaded = new XLWorkbook(source);
        var ws1 = loaded.Worksheet(1);
        var wrong = Enumerable.Range(0, Cells)
            .Where(i =>
            {
                var style = ws1.Cell(i + 1, 1).Style;
                return style.Font.FontSize != 6 + i * 0.5 ||
                       style.Fill.BackgroundColor != XLColor.FromArgb(i % 256, 255 - i % 256, i / 256) ||
                       style.Border.LeftBorder != Borders[i % Borders.Length] ||
                       style.Border.LeftBorderColor != XLColor.FromArgb(i % 256, 0, 0);
            })
            .ToArray();

        await Assert.That(string.Join(", ", wrong)).IsEqualTo(string.Empty);
    }
}
