using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Issue #966. A custom slicer or timeline style names its dxfs by index in the stylesheet's
/// <c>extLst</c>, and the save rebuilds <c>&lt;dxfs&gt;</c>. Each style element must still name the
/// dxf it named before the save.
/// </summary>
internal class SlicerTimelineStyleDxfTests
{
    private const string Extensions =
        """
        <x:extLst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <x:ext uri="{EB79DEF2-80B8-4f9b-8A8C-5B6D0D4C2E7A}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1">
              <x14:slicerStyle name="Custom Slicer">
                <x14:slicerStyleElements>
                  <x14:slicerStyleElement type="selectedItemWithData" dxfId="2"/>
                  <x14:slicerStyleElement type="unselectedItemWithData" dxfId="1"/>
                </x14:slicerStyleElements>
              </x14:slicerStyle>
            </x14:slicerStyles>
          </x:ext>
          <x:ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1">
              <x15:timelineStyle name="Custom Timeline">
                <x15:timelineStyleElements>
                  <x15:timelineStyleElement type="selectionLabel" dxfId="3"/>
                  <x15:timelineStyleElement type="timeLevel" dxfId="2"/>
                </x15:timelineStyleElements>
              </x15:timelineStyle>
            </x15:timelineStyles>
          </x:ext>
        </x:extLst>
        """;

    [Test]
    public async Task Every_slicer_and_timeline_style_element_points_at_its_original_dxf_after_a_save()
    {
        using var original = BuildWorkbook();
        var expected = ReadStyleDxfs(original);
        await Assert.That(expected.Count).IsEqualTo(4);

        original.Position = 0;
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(original))
            wb.SaveAs(saved);

        var actual = ReadStyleDxfs(saved);
        await Assert.That(actual).IsEquivalentTo(expected);
    }

    /// <summary>
    /// A conditional format on A1:A3 is dxf 0. Dxfs 1 to 3 are named only by the custom slicer
    /// and timeline styles, so the rebuilt collection must keep them.
    /// </summary>
    private static MemoryStream BuildWorkbook()
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("A3").Value = 3;
            ws.Range("A1:A3").AddConditionalFormat().WhenGreaterThan(1).Fill.SetBackgroundColor(XLColor.Red);
            wb.SaveAs(stream);
        }

        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            var dxfs = stylesheet.DifferentialFormats!;
            dxfs.Append(FillDxf("FF00FF00"));
            dxfs.Append(FillDxf("FF0000FF"));
            dxfs.Append(FillDxf("FFFFFF00"));
            dxfs.Count = (uint)dxfs.ChildElements.Count;

            stylesheet.RemoveAllChildren<StylesheetExtensionList>();
            stylesheet.Append(new StylesheetExtensionList(Extensions));
            stylesheet.Save();
        }

        stream.Position = 0;
        return stream;
    }

    private static DifferentialFormat FillDxf(string rgb) =>
        new(new Fill(new PatternFill(new BackgroundColor { Rgb = HexBinaryValue.FromString(rgb) })
            { PatternType = PatternValues.Solid }));

    private static List<string> ReadStyleDxfs(Stream package)
    {
        using var document = SpreadsheetDocument.Open(package, false);
        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        var dxfs = stylesheet.DifferentialFormats?.Elements<DifferentialFormat>().ToList() ?? [];

        return stylesheet.StylesheetExtensionList!.Descendants()
            .Where(e => e.LocalName is "slicerStyleElement" or "timelineStyleElement")
            .Select(e =>
            {
                var type = e.GetAttribute("type", "").Value;
                var id = int.Parse(e.GetAttribute("dxfId", "").Value!);
                var dxf = id < dxfs.Count ? dxfs[id].OuterXml : "<missing>";
                return $"{e.LocalName}/{type}: {dxf}";
            })
            .ToList();
    }
}
