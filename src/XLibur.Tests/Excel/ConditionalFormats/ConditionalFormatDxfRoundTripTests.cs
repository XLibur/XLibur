using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// A conditional format's dxf states only what it overrides, and the style model cannot tell an
/// explicit <c>&lt;i val="0"/&gt;</c>, a <c>&lt;charset&gt;</c> or a colourless solid fill from what it
/// leaves out. A rule the caller did not restyle keeps the dxf it was loaded with (#979).
/// </summary>
public class ConditionalFormatDxfRoundTripTests
{
    [Test]
    [Arguments(@"TryToLoad\InvalidPrintTitles2.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\formats.xlsx")]
    [Arguments(@"TryToLoad\DrawingSizeInPx.xlsx")]
    [Arguments(@"Other\StyleReferenceFiles\TransparentBackgroundFill\inputfile.xlsx")]
    [Arguments(@"Other\PivotTableReferenceFiles\TwoPivotTablesWithSingleSource\input.xlsx")]
    public async Task A_save_keeps_the_dxf_of_each_rule(string resource)
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource));
        var before = RuleDxfs(source);
        source.Position = 0;
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(source))
            wb.SaveAs(saved);

        await Assert.That(before).IsNotEmpty();
        await Assert.That(RuleDxfs(saved)).IsEquivalentTo(before);
    }

    [Test]
    public async Task A_restyled_rule_is_written_from_its_style()
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\InvalidPrintTitles2.xlsx"));
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            var cf = wb.Worksheets.SelectMany(ws => ws.ConditionalFormats).Single(c => c.Style.Font.Bold);
            cf.Style.Fill.BackgroundColor = XLColor.FromArgb(0x12, 0x34, 0x56);
            wb.SaveAs(saved);
        }

        var dxf = RuleDxfs(saved).Single(d => d.Contains("<b />"));
        await Assert.That(dxf).Contains("<bgColor rgb=\"FF123456\" />");
    }

    /// <summary>
    /// The distinct dxfs the sheets' rules point at, each without its alignment, which a conditional
    /// format cannot apply and XLibur does not write for one.
    /// </summary>
    private static HashSet<string> RuleDxfs(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        var dxfs = workbookPart.WorkbookStylesPart?.Stylesheet?.DifferentialFormats?
            .Elements<DifferentialFormat>().ToArray() ?? [];
        return workbookPart.WorksheetParts
            .SelectMany(part => part.Worksheet!.Descendants<ConditionalFormattingRule>())
            .Where(rule => rule.FormatId is not null)
            .Select(rule => Canonical(dxfs[(int)rule.FormatId!.Value]))
            .ToHashSet();
    }

    private static string Canonical(DifferentialFormat dxf)
    {
        var element = XElement.Parse(dxf.OuterXml);
        element.Elements().Where(e => e.Name.LocalName == "alignment").Remove();
        element.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        return element.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }
}
