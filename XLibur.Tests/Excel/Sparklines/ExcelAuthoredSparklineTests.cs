using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Sparklines;

/// <summary>
/// Sparklines as Excel writes them, or as a hand-edited file spells them, where XLibur's writer
/// always emits every optional child (issue 845, spec 48). Each input is made by saving sparklines
/// through XLibur and hand-editing the part, as the conditional-format tests do.
/// </summary>
public class ExcelAuthoredSparklineTests
{
    private static MemoryStream Save(XLWorkbook wb)
    {
        var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms;
    }

    /// <summary>Two sparklines (G1, G2) in one group, drawing A1:C1 and A2:C2.</summary>
    private static MemoryStream SaveSparklines()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 3;
        ws.Cell("C1").Value = 2;
        ws.Cell("A2").Value = 4;
        ws.Cell("B2").Value = 6;
        ws.Cell("C2").Value = 5;
        ws.SparklineGroups.Add("G1:G2", "A1:C2");
        return Save(wb);
    }

    private static string Edit(string xml, string pattern, string replacement)
    {
        var edited = Regex.Replace(xml, pattern, replacement);
        if (edited == xml)
            throw new InvalidOperationException($"The edit '{pattern}' matched nothing in: {xml}");
        return edited;
    }

    private static MemoryStream SaveAgain(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        return Save(wb);
    }

    private static MemoryStream WithSecondSparkline(Func<string, string> edit) =>
        SaveSparklines().RewriteSheet1(xml =>
        {
            // Edit only the second <x14:sparkline>, so the first one proves the others are kept.
            var match = Regex.Matches(xml, "<x14:sparkline>.*?</x14:sparkline>")[1];
            return xml.Replace(match.Value, edit(match.Value));
        });

    private static async Task AssertFirstSparklineKeptAndRoundTrips(MemoryStream package)
    {
        using var wb = new XLWorkbook(package);
        var sparklines = wb.Worksheet("S").SparklineGroups.SelectMany(g => g).ToList();
        await Assert.That(sparklines.Count).IsEqualTo(1);
        await Assert.That(sparklines[0].Location.Address.ToString()).IsEqualTo("G1");

        using var saved = Save(wb);
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("S").SparklineGroups.SelectMany(g => g).Count()).IsEqualTo(1);
    }

    #region Sparkline without its cell or data range

    [Test]
    public async Task SparklineWithoutACellIsSkippedAndTheOthersKept()
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:sqref>.*?</xm:sqref>", ""));
        await AssertFirstSparklineKeptAndRoundTrips(package);
    }

    [Test]
    public async Task SparklineWithoutADataRangeIsSkippedAndTheOthersKept()
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:f>.*?</xm:f>", ""));
        await AssertFirstSparklineKeptAndRoundTrips(package);
    }

    [Test]
    public async Task SparklineWithAnEmptyCellIsSkippedAndTheOthersKept()
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:sqref>.*?</xm:sqref>", "<xm:sqref></xm:sqref>"));
        await AssertFirstSparklineKeptAndRoundTrips(package);
    }

    [Test]
    public async Task SparklineWithAnEmptyDataRangeIsSkippedAndTheOthersKept()
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:f>.*?</xm:f>", "<xm:f></xm:f>"));
        await AssertFirstSparklineKeptAndRoundTrips(package);
    }

    [Test]
    public async Task SparklineWithAnUnparseableCellIsSkippedAndTheOthersKept()
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:sqref>.*?</xm:sqref>", "<xm:sqref>not a cell</xm:sqref>"));
        await AssertFirstSparklineKeptAndRoundTrips(package);
    }

    [Test]
    [Arguments("S!nonsense")]
    [Arguments("Gone!A2:C2")]
    [Arguments("#REF!A2:C2")]
    public async Task SparklineWhoseDataNamesNoRangeIsKeptAsAnInvalidOneAndSaves(string reference)
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:f>.*?</xm:f>", $"<xm:f>{reference}</xm:f>"));

        using var wb = new XLWorkbook(package);
        var sparklines = wb.Worksheet("S").SparklineGroups.SelectMany(g => g).OrderBy(s => s.Location.Address.RowNumber).ToList();
        await Assert.That(sparklines.Count).IsEqualTo(2);
        await Assert.That(sparklines[0].IsValid).IsTrue();
        await Assert.That(sparklines[1].IsValid).IsFalse();

        using var saved = Save(wb);
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("S").SparklineGroups.SelectMany(g => g).Count()).IsEqualTo(2);
    }

    [Test]
    [Arguments("A1:")]
    [Arguments("S!")]
    [Arguments("foo bar")]
    [Arguments("!!")]
    [Arguments("Other!G2")]
    public async Task SparklineWithAMalformedOrForeignCellLoadsAndTheOthersAreKept(string cell)
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:sqref>.*?</xm:sqref>", $"<xm:sqref>{cell}</xm:sqref>"));
        await AssertFirstSparklineKeptAndRoundTrips(package);
    }

    [Test]
    [Arguments("A1:")]
    [Arguments("S!")]
    [Arguments("!!")]
    public async Task SparklineWithAMalformedDataRangeLoads(string reference)
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:f>.*?</xm:f>", $"<xm:f>{reference}</xm:f>"));

        using var wb = new XLWorkbook(package);
        await Assert.That(wb.Worksheet("S").SparklineGroups.SelectMany(g => g).Count()).IsGreaterThanOrEqualTo(1);
        using var saved = Save(wb);
    }

    [Test]
    public async Task InvalidLoadedSparklineIsLeftOutWhenTheSheetIsCopied()
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:f>.*?</xm:f>", "<xm:f>Gone!A2:C2</xm:f>"));
        using var wb = new XLWorkbook(package);
        var copy = wb.Worksheet("S").CopyTo("S2");

        await Assert.That(copy.SparklineGroups.SelectMany(g => g).Count()).IsEqualTo(1);
    }

    [Test]
    public async Task SparklineWithATwoDimensionalDataRangeIsSkippedAndTheOthersKept()
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:f>.*?</xm:f>", "<xm:f>S!A1:C2</xm:f>"));
        await AssertFirstSparklineKeptAndRoundTrips(package);
    }

    [Test]
    public async Task SparklineWithAMultiCellLocationIsSkippedAndTheOthersKept()
    {
        using var package = WithSecondSparkline(s => Edit(s, "<xm:sqref>.*?</xm:sqref>", "<xm:sqref>G2:G3</xm:sqref>"));
        await AssertFirstSparklineKeptAndRoundTrips(package);
    }

    [Test]
    public async Task GroupWithNoSparklinesLoadsAndIsNotKept()
    {
        using var package = SaveSparklines()
            .RewriteSheet1(xml => Edit(xml, "<x14:sparklines>.*?</x14:sparklines>", "<x14:sparklines/>"));

        using var wb = new XLWorkbook(package);
        await Assert.That(wb.Worksheet("S").SparklineGroups.Count()).IsEqualTo(0);

        using var saved = Save(wb);
        await Assert.That(saved.Sheet1Xml()).DoesNotContain("sparklineGroup");
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("S").SparklineGroups.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task GroupWithOnlySkippedSparklinesLoadsAndIsNotKept()
    {
        using var package = SaveSparklines()
            .RewriteSheet1(xml => Edit(xml, "<xm:f>.*?</xm:f>", ""));

        using var wb = new XLWorkbook(package);
        await Assert.That(wb.Worksheet("S").SparklineGroups.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task GroupWithoutASparklinesElementLoads()
    {
        using var package = SaveSparklines()
            .RewriteSheet1(xml => Edit(xml, "<x14:sparklines>.*?</x14:sparklines>", ""));

        using var wb = new XLWorkbook(package);
        await Assert.That(wb.Worksheet("S").SparklineGroups.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task GroupDateRangeThatIsEmptyOrUnparseableLoads()
    {
        using var package = SaveSparklines()
            .RewriteSheet1(xml => Edit(xml, "<x14:sparklines>", "<xm:f>nonsense</xm:f><x14:sparklines>"));

        using var wb = new XLWorkbook(package);
        var group = wb.Worksheet("S").SparklineGroups.Single();
        await Assert.That(group.DateRange).IsNull();
        await Assert.That(group.Count()).IsEqualTo(2);
    }

    [Test]
    public async Task GroupDateRangeThatIsEmptyLoads()
    {
        using var package = SaveSparklines()
            .RewriteSheet1(xml => Edit(xml, "<x14:sparklines>", "<xm:f></xm:f><x14:sparklines>"));

        using var wb = new XLWorkbook(package);
        var group = wb.Worksheet("S").SparklineGroups.Single();
        await Assert.That(group.DateRange).IsNull();
    }

    [Test]
    public async Task GroupDateRangeThatIsNotOneDimensionalLoads()
    {
        using var package = SaveSparklines()
            .RewriteSheet1(xml => Edit(xml, "<x14:sparklines>", "<xm:f>S!A1:C2</xm:f><x14:sparklines>"));

        using var wb = new XLWorkbook(package);
        var group = wb.Worksheet("S").SparklineGroups.Single();
        await Assert.That(group.DateRange).IsNull();
    }

    [Test]
    public async Task GroupDateRangeSurvivesASave()
    {
        using var package = SaveSparklines()
            .RewriteSheet1(xml => Edit(xml, "<x14:sparklines>", "<xm:f>S!A1:C1</xm:f><x14:sparklines>"));
        using var saved = SaveAgain(package);

        using var wb = new XLWorkbook(saved);
        var group = wb.Worksheet("S").SparklineGroups.Single();
        await Assert.That(group.DateRange!.RangeAddress.ToString()).IsEqualTo("A1:C1");
    }

    #endregion

    #region Axis colour that is not an rgb

    private static MemoryStream WithAxisColor(string attributes) =>
        SaveSparklines().RewriteSheet1(xml => Edit(xml, "<x14:colorAxis [^>]*/>", $"<x14:colorAxis {attributes}/>"));

    [Test]
    public async Task AxisColourWithoutAnyAttributeIsAutomatic()
    {
        using var package = SaveSparklines().RewriteSheet1(xml => Edit(xml, "<x14:colorAxis [^>]*/>", "<x14:colorAxis/>"));

        using var wb = new XLWorkbook(package);
        var color = wb.Worksheet("S").SparklineGroups.Single().HorizontalAxis.Color;
        await Assert.That(color.ColorType).IsEqualTo(XLColorType.Automatic);
    }

    [Test]
    public async Task AxisColourAsATheme()
    {
        using var package = WithAxisColor("theme=\"4\" tint=\"-0.25\"");
        using var saved = SaveAgain(package);

        await Assert.That(saved.Sheet1Xml()).Contains("<x14:colorAxis theme=\"4\"");
        using var wb = new XLWorkbook(saved);
        var color = wb.Worksheet("S").SparklineGroups.Single().HorizontalAxis.Color;
        await Assert.That(color.ColorType).IsEqualTo(XLColorType.Theme);
        await Assert.That(color.ThemeColor).IsEqualTo(XLThemeColor.Accent1);
        await Assert.That(color.ThemeTint).IsEqualTo(-0.25);
    }

    [Test]
    public async Task AxisColourAsAnIndex()
    {
        using var package = WithAxisColor("indexed=\"10\"");
        using var saved = SaveAgain(package);

        await Assert.That(saved.Sheet1Xml()).Contains("<x14:colorAxis indexed=\"10\"");
        using var wb = new XLWorkbook(saved);
        var color = wb.Worksheet("S").SparklineGroups.Single().HorizontalAxis.Color;
        await Assert.That(color.ColorType).IsEqualTo(XLColorType.Indexed);
        await Assert.That(color.Indexed).IsEqualTo(10);
    }

    [Test]
    public async Task AxisColourAsAuto()
    {
        using var package = WithAxisColor("auto=\"1\"");
        using var saved = SaveAgain(package);

        await Assert.That(saved.Sheet1Xml()).Contains("<x14:colorAxis auto=\"1\"");
        using var wb = new XLWorkbook(saved);
        var color = wb.Worksheet("S").SparklineGroups.Single().HorizontalAxis.Color;
        await Assert.That(color.ColorType).IsEqualTo(XLColorType.Automatic);
    }

    [Test]
    public async Task AxisColourAsRgb()
    {
        using var package = WithAxisColor("rgb=\"FF123456\"");
        using var saved = SaveAgain(package);

        await Assert.That(saved.Sheet1Xml()).Contains("<x14:colorAxis rgb=\"FF123456\"");
        using var wb = new XLWorkbook(saved);
        var color = wb.Worksheet("S").SparklineGroups.Single().HorizontalAxis.Color;
        await Assert.That(color).IsEqualTo(XLColor.FromArgb(0xFF123456));
    }

    [Test]
    public async Task AxisColourSetThroughTheApiAsAThemeIsWrittenAsATheme()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        var group = ws.SparklineGroups.Add("G1", "A1:C1");
        group.HorizontalAxis.Color = XLColor.FromTheme(XLThemeColor.Accent2);
        using var saved = Save(wb);

        await Assert.That(saved.Sheet1Xml()).Contains("<x14:colorAxis theme=\"5\"");
        using var reloaded = new XLWorkbook(saved);
        var color = reloaded.Worksheet("S").SparklineGroups.Single().HorizontalAxis.Color;
        await Assert.That(color).IsEqualTo(XLColor.FromTheme(XLThemeColor.Accent2));
    }

    #endregion
}
