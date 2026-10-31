using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// Optional attributes and elements of a conditional format rule that XLibur's own writer always
/// emits, so a file saved by XLibur never leaves them out but a file from elsewhere may (spec 48).
/// Each input is a saved package with one optional part removed from the sheet XML.
/// </summary>
public class ConditionalFormatOptionalElementTests
{
    private static MemoryStream SaveCellIs()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A1").Value = 5;
        ws.Range("A1:A3").AddConditionalFormat().WhenLessThan(10).Fill.SetBackgroundColor(XLColor.Yellow);
        var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms;
    }

    private static string Edit(string xml, string pattern, string replacement)
    {
        var edited = Regex.Replace(xml, pattern, replacement);
        if (edited == xml)
            throw new InvalidOperationException($"The edit '{pattern}' matched nothing in: {xml}");
        return edited;
    }

    private static XLWorkbook Load(MemoryStream package)
    {
        package.Position = 0;
        return new XLWorkbook(package);
    }

    private static MemoryStream Resave(MemoryStream package)
    {
        using var wb = Load(package);
        var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms;
    }

    [Test]
    public async Task CellIsRuleWithoutAnOperatorTakesTheSchemaDefault()
    {
        // ST_ConditionalFormattingOperator is optional on cfRule; Excel reads a missing one as lessThan.
        using var package = SaveCellIs().RewriteSheet1(xml => Edit(xml, " operator=\"[A-Za-z]+\"", ""));

        using var wb = Load(package);
        var cf = wb.Worksheet("S").ConditionalFormats.Single();
        await Assert.That(cf.Operator).IsEqualTo(XLCFOperator.LessThan);
        using var saved = Resave(package);
        await Assert.That(saved.Sheet1Xml()).Contains("<x:cfRule type=\"cellIs\"");
    }

    [Test]
    public async Task RuleWhoseDifferentialFormatIsMissingLoadsWithoutAStyle()
    {
        using var package = SaveCellIs().RewriteSheet1(xml => Edit(xml, "dxfId=\"\\d+\"", "dxfId=\"99\""));

        using var wb = Load(package);
        var cf = wb.Worksheet("S").ConditionalFormats.Single();
        await Assert.That(cf.Operator).IsEqualTo(XLCFOperator.LessThan);
        await Assert.That(cf.Style.Fill.BackgroundColor).IsNotEqualTo(XLColor.Yellow);
    }

    /// <summary>
    /// The rule is kept, not dropped: its condition still applies, only its format is lost. Dropping it
    /// would silently delete a rule the author wrote. On a save it comes back without a format.
    /// </summary>
    [Test]
    public async Task RuleWhoseDifferentialFormatIsMissingIsKeptThroughASave()
    {
        using var package = SaveCellIs().RewriteSheet1(xml => Edit(xml, "dxfId=\"\\d+\"", "dxfId=\"99\""));
        using var saved = Resave(package);

        using var wb = Load(saved);
        var cf = wb.Worksheet("S").ConditionalFormats.Single();
        await Assert.That(cf.Operator).IsEqualTo(XLCFOperator.LessThan);
        await Assert.That(cf.Values[1].Value).IsEqualTo("10");
        await Assert.That(cf.Style.Fill.BackgroundColor).IsNotEqualTo(XLColor.Yellow);
    }

    /// <summary>
    /// A bar from Excel 2007 with no lengths is drawn 10 to 90 (checked over COM); a bar XLibur
    /// builds is drawn the whole cell. They look different, so they must not compare equal and be
    /// merged into one.
    /// </summary>
    [Test]
    public async Task LoadedLegacyDataBarIsNotEqualToABarBuiltWithDefaults()
    {
        using var legacy = SaveDataBar().RewriteSheet1(xml =>
            Edit(xml, "<x:extLst><x:ext [^>]*uri=\"\\{78C0D931.*?</x:ext></x:extLst>", ""));
        using var loaded = Load(legacy);
        using var built = new XLWorkbook();
        var ws = built.AddWorksheet("S");
        ws.Range("A1:A2").AddConditionalFormat().DataBar(XLColor.Red, XLColor.Blue).LowestValue().HighestValue();

        var a = loaded.Worksheet("S").ConditionalFormats.Single();
        var b = ws.ConditionalFormats.Single();
        await Assert.That(XLibur.Excel.ConditionalFormats.XLConditionalFormat.NoRangeComparer.Equals(a, b)).IsFalse();
    }

    [Test]
    public async Task ConditionalFormattingWithoutARangeIsSkipped()
    {
        using var package = SaveCellIs().RewriteSheet1(xml => Edit(xml, " sqref=\"A1:A3\"", ""));

        using var wb = Load(package);
        await Assert.That(wb.Worksheet("S").ConditionalFormats.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task DataBarTwinIdWithoutBracesStillPairsWithItsRule()
    {
        using var package = SaveDataBar().RewriteSheet1(xml =>
            Edit(xml, "\\{([0-9a-fA-F-]{36})\\}", "$1"));

        using var wb = Load(package);
        await Assert.That(wb.Worksheet("S").ConditionalFormats.Count()).IsEqualTo(1);
    }

    private static MemoryStream SaveDataBar()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A1").Value = 1;
        ws.Range("A1:A2").AddConditionalFormat().DataBar(XLColor.Red, XLColor.Blue).LowestValue().HighestValue();
        var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms;
    }
}
