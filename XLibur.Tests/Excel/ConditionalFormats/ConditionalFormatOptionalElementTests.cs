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
