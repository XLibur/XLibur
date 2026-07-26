using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Clearing;

public class ClearingTests
{
    private static readonly XLColor BackgroundColor = XLColor.LightBlue;
    private static readonly XLColor ForegroundColor = XLColor.DarkBrown;

    private static async Task<XLWorkbook> SetupWorkbook()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        var c = ws.FirstCell()
            .SetValue("Hello world!");

        c.GetComment().AddText("Some comment");

        c.Style.Fill.BackgroundColor = BackgroundColor;
        c.Style.Font.FontColor = ForegroundColor;
        c.CreateDataValidation().Custom("B1");

        c = ws.FirstCell()
            .CellBelow()
            .SetFormulaA1("=LEFT(A1,5)");

        c.GetComment().AddText("Another comment");

        c.Style.Fill.BackgroundColor = BackgroundColor;
        c.Style.Font.FontColor = ForegroundColor;

        c = ws.FirstCell()
            .CellBelow(2)
            .SetValue(new DateTime(2018, 1, 15, 0, 0, 0, DateTimeKind.Unspecified));

        c.GetComment().AddText("A date");

        c.Style.Fill.BackgroundColor = BackgroundColor;
        c.Style.Font.FontColor = ForegroundColor;

        ws.Column(1)
            .AddConditionalFormat().WhenStartsWith("Hell")
            .Fill.SetBackgroundColor(XLColor.Red)
            .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
            .Border.SetOutsideBorderColor(XLColor.Blue)
            .Font.SetBold();

        await Assert.That(ws.Cell("A1").Value.Type).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A2").Value.Type).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A3").Value.Type).IsEqualTo(XLDataType.DateTime);

        await Assert.That(ws.Cell("A1").HasFormula).IsEqualTo(false);
        await Assert.That(ws.Cell("A2").HasFormula).IsEqualTo(true);
        await Assert.That(ws.Cell("A1").HasFormula).IsEqualTo(false);

        foreach (var cell in ws.Range("A1:A3").Cells())
        {
            await Assert.That(cell.Style.Fill.BackgroundColor).IsEqualTo(BackgroundColor);
            await Assert.That(cell.Style.Font.FontColor).IsEqualTo(ForegroundColor);
            await Assert.That(ws.ConditionalFormats.Any()).IsTrue();
            await Assert.That(cell.HasComment).IsTrue();
        }

        await Assert.That(ws.Cell("A1").GetDataValidation().Value).IsEqualTo("B1");

        return wb;
    }

    [Test]
    public async Task WorksheetClearAll()
    {
        using var wb = await SetupWorkbook();
        var ws = wb.Worksheets.First();

        ws.Clear();

        foreach (var c in ws.Range("A1:A10").Cells())
        {
            await Assert.That(c.IsEmpty()).IsTrue();
            await Assert.That(c.DataType).IsEqualTo(XLDataType.Blank);
            await Assert.That(c.Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
            await Assert.That(c.Style.Font.FontColor).IsEqualTo(ws.Style.Font.FontColor);
            await Assert.That(ws.ConditionalFormats.Any()).IsFalse();
            await Assert.That(c.HasComment).IsFalse();
            await Assert.That(c.GetDataValidation().Value).IsEqualTo(string.Empty);
        }
    }

    [Test]
    public async Task WorksheetClearContents()
    {
        using var wb = await SetupWorkbook();
        var ws = wb.Worksheets.First();

        ws.Clear(XLClearOptions.Contents);

        foreach (var c in ws.Range("A1:A3").Cells())
        {
            await Assert.That(ws.Cell("A1").DataType).IsEqualTo(XLDataType.Blank);
            await Assert.That(c.IsEmpty(XLCellsUsedOptions.Contents)).IsTrue();

            await Assert.That(c.Style.Fill.BackgroundColor).IsEqualTo(BackgroundColor);
            await Assert.That(c.Style.Font.FontColor).IsEqualTo(ForegroundColor);
            await Assert.That(ws.ConditionalFormats.Any()).IsTrue();
            await Assert.That(c.HasComment).IsTrue();
        }

        await Assert.That(ws.Cell("A1").GetDataValidation().Value).IsEqualTo("B1");
    }

    [Test]
    public async Task WorksheetClearNormalFormats()
    {
        using var wb = await SetupWorkbook();
        var ws = wb.Worksheets.First();

        ws.Clear(XLClearOptions.NormalFormats);

        foreach (var c in ws.Range("A1:A3").Cells())
        {
            await Assert.That(c.IsEmpty()).IsFalse();
            await Assert.That(c.Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
            await Assert.That(c.Style.Font.FontColor).IsEqualTo(ws.Style.Font.FontColor);
            await Assert.That(ws.ConditionalFormats.Any()).IsTrue();
            await Assert.That(c.HasComment).IsTrue();
        }

        await Assert.That(ws.Cell("A1").DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A2").DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A3").DataType).IsEqualTo(XLDataType.DateTime);

        await Assert.That(ws.Cell("A1").GetDataValidation().Value).IsEqualTo("B1");
    }

    [Test]
    public async Task WorksheetClearConditionalFormats()
    {
        using var wb = await SetupWorkbook();
        var ws = wb.Worksheets.First();

        ws.Clear(XLClearOptions.ConditionalFormats);

        foreach (var c in ws.Range("A1:A3").Cells())
        {
            await Assert.That(c.IsEmpty()).IsFalse();
            await Assert.That(c.Style.Fill.BackgroundColor).IsEqualTo(BackgroundColor);
            await Assert.That(c.Style.Font.FontColor).IsEqualTo(ForegroundColor);
            await Assert.That(ws.ConditionalFormats.Any()).IsFalse();
            await Assert.That(c.HasComment).IsTrue();
        }

        await Assert.That(ws.Cell("A1").DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A2").DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A3").DataType).IsEqualTo(XLDataType.DateTime);

        await Assert.That(ws.Cell("A1").GetDataValidation().Value).IsEqualTo("B1");
    }

    [Test]
    public async Task WorksheetClearComments()
    {
        using var wb = await SetupWorkbook();
        var ws = wb.Worksheets.First();

        ws.Clear(XLClearOptions.Comments);

        foreach (var c in ws.Range("A1:A3").Cells())
        {
            await Assert.That(c.IsEmpty()).IsFalse();
            await Assert.That(c.Style.Fill.BackgroundColor).IsEqualTo(BackgroundColor);
            await Assert.That(c.Style.Font.FontColor).IsEqualTo(ForegroundColor);
            await Assert.That(ws.ConditionalFormats.Any()).IsTrue();
            await Assert.That(c.HasComment).IsFalse();
        }

        await Assert.That(ws.Cell("A1").DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A2").DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A3").DataType).IsEqualTo(XLDataType.DateTime);

        await Assert.That(ws.Cell("A1").GetDataValidation().Value).IsEqualTo("B1");
    }

    [Test]
    public async Task WorksheetClearDataValidation()
    {
        using var wb = await SetupWorkbook();
        var ws = wb.Worksheets.First();

        ws.Clear(XLClearOptions.DataValidation);

        foreach (var c in ws.Range("A1:A3").Cells())
        {
            await Assert.That(c.IsEmpty()).IsFalse();
            await Assert.That(c.Style.Fill.BackgroundColor).IsEqualTo(BackgroundColor);
            await Assert.That(c.Style.Font.FontColor).IsEqualTo(ForegroundColor);
            await Assert.That(ws.ConditionalFormats.Any()).IsTrue();
            await Assert.That(c.HasComment).IsTrue();
        }

        await Assert.That(ws.Cell("A1").DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A2").DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.Cell("A3").DataType).IsEqualTo(XLDataType.DateTime);

        await Assert.That(ws.Cell("A1").GetDataValidation().Value).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task DeleteClearedCellValue()
    {
        using var ms = new MemoryStream();
        using (var wb = await SetupWorkbook())
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.Cell("A1").GetText()).IsEqualTo("Hello world!");
            await Assert.That(ws.Cell("A3").GetDateTime()).IsEqualTo(new DateTime(2018, 1, 15, 0, 0, 0, DateTimeKind.Unspecified));

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            ws.Clear(XLClearOptions.Contents);
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(Blank.Value);
            await Assert.That(() => ws.Cell("A3").GetDateTime()).Throws<InvalidCastException>();

            wb.Save();
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(Blank.Value);
            await Assert.That(() => ws.Cell("A3").GetDateTime()).Throws<InvalidCastException>();
        }
    }

    [Test]
    [Arguments(XLClearOptions.All, 2)]
    [Arguments(XLClearOptions.AllContents, 4)]
    [Arguments(XLClearOptions.AllFormats, 4)]
    [Arguments(XLClearOptions.Contents, 4)]
    [Arguments(XLClearOptions.MergedRanges, 2)]
    public async Task CanClearMergedRanges(XLClearOptions options, int expectedCount)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");

        ws.Range("A1:C3").Merge();
        ws.Range("A4:B6").Merge();
        ws.Range("D1:F3").Merge();
        ws.Range("E4:F6").Merge();

        ws.Range("C1:D6").Clear(options);

        await Assert.That(ws.MergedRanges.Count).IsEqualTo(expectedCount);
    }
}
