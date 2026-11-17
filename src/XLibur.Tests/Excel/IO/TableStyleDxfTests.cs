using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Issue #901. A custom table style names its dxfs by index, and the save rebuilds
/// <c>&lt;dxfs&gt;</c>. Excel cannot open a file whose table style points past the end of the list.
/// </summary>
internal class TableStyleDxfTests
{
    [Test]
    [Arguments(@"TryToLoad\TableWithCustomTheme.xlsx")]
    [Arguments(@"Other\GoogleSheets\2619_exported-broken2.xlsx")]
    public async Task Every_table_style_element_points_at_its_original_dxf_after_a_save(string fixture)
    {
        using var original = TestHelper.OpenResource(fixture);
        var expected = ReadTableStyleDxfs(original);
        await Assert.That(expected.Count).IsGreaterThan(0);

        original.Position = 0;
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(original))
            wb.SaveAs(saved);

        var actual = ReadTableStyleDxfs(saved);
        await Assert.That(actual).IsEquivalentTo(expected);
    }

    private static List<string> ReadTableStyleDxfs(Stream package)
    {
        using var document = SpreadsheetDocument.Open(package, false);
        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        var dxfs = stylesheet.DifferentialFormats?.Elements<DifferentialFormat>().ToList() ?? [];

        return stylesheet.TableStyles!.Elements<TableStyle>()
            .SelectMany(style => style.Elements<TableStyleElement>()
                .Where(element => element.FormatId is not null)
                .Select(element =>
                {
                    var id = (int)element.FormatId!.Value;
                    var dxf = id < dxfs.Count ? dxfs[id].OuterXml : "<missing>";
                    return $"{style.Name}/{element.Type!.InnerText}/{element.Size?.Value}: {dxf}";
                }))
            .ToList();
    }
}
