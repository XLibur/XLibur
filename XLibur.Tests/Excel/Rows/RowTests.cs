using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.Rows;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Rows;

public class RowTests
{
    [Test]
    public async Task RowsUsedIsFast()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().SetValue("Hello world!");
        var rowsUsed = ws.Column(1).AsRange().RowsUsed();
        await Assert.That(rowsUsed.Count()).IsEqualTo(1);
    }

    [Test]
    public async Task CopyRow()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Test").Style.Font.SetBold();
        ws.FirstRow().CopyTo(ws.Row(2));

        await Assert.That(ws.Cell("A2").Style.Font.Bold).IsTrue();
    }

    [Test]
    public async Task InsertingRowsAbove1()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
        ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

        var row1 = ws.Row(1);
        var row2 = ws.Row(2);
        var row3 = ws.Row(3);

        var rowIns = ws.Row(1).InsertRowsAbove(1).First();

        await Assert.That(ws.Row(1).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(ws.Row(1).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(ws.Row(1).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);

        await Assert.That(ws.Row(2).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(2).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(2).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Row(3).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(ws.Row(3).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(ws.Row(3).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(ws.Row(4).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(4).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(4).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Row(3).Cell(2).GetText()).IsEqualTo("X");

        await Assert.That(rowIns.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(rowIns.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);
        await Assert.That(rowIns.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(ws.Style.Fill.BackgroundColor);

        await Assert.That(row1.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row1.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row1.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(row2.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(row2.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(row2.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(row3.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row3.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row3.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(row2.Cell(2).GetText()).IsEqualTo("X");
    }

    [Test]
    public async Task InsertingRowsAbove2()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
        ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

        var row1 = ws.Row(1);
        var row2 = ws.Row(2);
        var row3 = ws.Row(3);

        var rowIns = ws.Row(2).InsertRowsAbove(1).First();

        await Assert.That(ws.Row(1).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(1).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(1).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Row(2).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(2).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(2).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Row(3).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(ws.Row(3).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(ws.Row(3).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(ws.Row(4).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(4).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(4).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Row(3).Cell(2).GetText()).IsEqualTo("X");

        await Assert.That(rowIns.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(rowIns.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(rowIns.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(row1.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row1.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row1.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(row2.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(row2.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(row2.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(row3.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row3.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row3.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(row2.Cell(2).GetText()).IsEqualTo("X");
    }

    [Test]
    public async Task InsertingRowsAbove3()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
        ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

        var row1 = ws.Row(1);
        var row2 = ws.Row(2);
        var row3 = ws.Row(3);

        var rowIns = ws.Row(3).InsertRowsAbove(1).First();

        await Assert.That(ws.Row(1).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(1).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(1).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Row(2).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(ws.Row(2).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(ws.Row(2).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(ws.Row(3).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(ws.Row(3).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(ws.Row(3).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(ws.Row(4).Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(4).Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Row(4).Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(ws.Row(2).Cell(2).GetText()).IsEqualTo("X");

        await Assert.That(rowIns.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(rowIns.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(rowIns.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(row1.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row1.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row1.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(row2.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        await Assert.That(row2.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Green);
        await Assert.That(row2.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(row3.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row3.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row3.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        await Assert.That(row2.Cell(2).GetText()).IsEqualTo("X");
    }

    [Test]
    public async Task InsertingRowsAbove4()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Row(2).Height = 15;
        ws.Row(3).Height = 20;
        ws.Row(4).Height = 25;
        ws.Row(5).Height = 35;

        ws.Row(2).FirstCell().SetValue("Row height: 15");
        ws.Row(3).FirstCell().SetValue("Row height: 20");
        ws.Row(4).FirstCell().SetValue("Row height: 25");
        ws.Row(5).FirstCell().SetValue("Row height: 35");

        ws.Range("3:3").InsertRowsAbove(1);

        await Assert.That(ws.Row(2).Height).IsEqualTo(15);
        await Assert.That(ws.Row(4).Height).IsEqualTo(20);
        await Assert.That(ws.Row(5).Height).IsEqualTo(25);
        await Assert.That(ws.Row(6).Height).IsEqualTo(35);

        await Assert.That(ws.Row(3).Height).IsEqualTo(20);
        ws.Row(3).ClearHeight();
        await Assert.That(ws.Row(3).Height).IsEqualTo(ws.RowHeight);
    }

    [Test]
    public async Task NoRowsUsed()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var count = ws.RowsUsed().Count();
        count += ws.Range("A1:C3").RowsUsed().Count();

        await Assert.That(count).IsEqualTo(0);
    }

    [Test]
    public async Task RowUsed()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 2).SetValue("Test");
        ws.Cell(1, 3).SetValue("Test");

        var fromRow = ws.Row(1).RowUsed();
        await Assert.That(fromRow.RangeAddress.ToStringRelative()).IsEqualTo("B1:C1");

        var fromRange = ws.Range("A1:E1").FirstRow()!.RowUsed();
        await Assert.That(fromRange.RangeAddress.ToStringRelative()).IsEqualTo("B1:C1");
    }

    [Test]
    public async Task RowsUsedWithDataValidation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().SetValue("Hello world!");
        ws.Range("A1:A100").CreateDataValidation().WholeNumber.EqualTo(1);

        var range = ws.Column(1).AsRange();

        await Assert.That(range.RowsUsed(XLCellsUsedOptions.DataValidation).Count()).IsEqualTo(100);
        await Assert.That(range.RowsUsed(XLCellsUsedOptions.All).Count()).IsEqualTo(100);
    }

    [Test]
    public async Task RowsUsedWithConditionalFormatting()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().SetValue("Hello world!");
        ws.Range("A1:A100").AddConditionalFormat().WhenStartsWith("Hell").Fill.SetBackgroundColor(XLColor.Red).Font.SetFontColor(XLColor.White);

        var range = ws.Column(1).AsRange();

        await Assert.That(range.RowsUsed(XLCellsUsedOptions.ConditionalFormats).Count()).IsEqualTo(100);
        await Assert.That(range.RowsUsed(XLCellsUsedOptions.All).Count()).IsEqualTo(100);
    }

    [Test]
    public async Task UngroupFromAll()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        ws.Rows(1, 2).Group();
        ws.Rows(1, 2).Ungroup(true);

        await Assert.That(ws.Row(1).OutlineLevel).IsEqualTo(0);
        await Assert.That(ws.Row(2).OutlineLevel).IsEqualTo(0);
    }

    [Test]
    public async Task NegativeRowNumberIsInvalid()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1") as XLWorksheet;

        var row = new XLRow(ws!, -1);

        await Assert.That(row.RangeAddress.IsValid).IsFalse();
    }

    [Test]
    public async Task DeleteRowOnWorksheetWithComment()
    {
        var ws = new XLWorkbook().AddWorksheet();
        ws.Cell(4, 1).GetComment().AddText("test");
        ws.Column(1).Width = 100;
        await Assert.That(() => ws.Row(1).Delete()).ThrowsNothing();
    }

    [Test]
    public async Task AssignWorksheetRowHeightWhenAllRowsChanged()
    {
        var ws = new XLWorkbook().AddWorksheet();
        var rows = ws.Rows();

        rows.Height = 30;

        await Assert.That(ws.Row(11).Height).IsEqualTo(30).Within(XLHelper.Epsilon);
        await Assert.That(ws.RowHeight).IsEqualTo(30).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task PreserveWorksheetRowHeightWhenNotAllRowsChanged()
    {
        var ws = new XLWorkbook().AddWorksheet();
        var defaultRowHeight = ws.RowHeight;
        var rows = ws.Rows(1, XLHelper.MaxRowNumber);

        rows.Height = 30;

        await Assert.That(ws.Row(11).Height).IsEqualTo(30).Within(XLHelper.Epsilon);
        await Assert.That(ws.RowHeight).IsEqualTo(defaultRowHeight).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task PreserveWorksheetRowHeightWhenUsedRowsChanged()
    {
        var ws = new XLWorkbook().AddWorksheet();
        ws.Cells("A1:E5").Value = "Not empty";
        var defaultRowHeight = ws.RowHeight;
        var rows = ws.RowsUsed(XLCellsUsedOptions.Contents);

        rows.Height = 30;

        await Assert.That(ws.Row(3).Height).IsEqualTo(30).Within(XLHelper.Epsilon);
        await Assert.That(ws.Row(11).Height).IsEqualTo(defaultRowHeight).Within(XLHelper.Epsilon);
        await Assert.That(ws.RowHeight).IsEqualTo(defaultRowHeight).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task LoadingDataOnlyRows_DoesNotCreateXLRowObjects()
    {
        // Data-only rows (no custom height, style, hidden, etc.) should not
        // create XLRow objects in RowsCollection during loading.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "Hello";
        ws.Cell("A2").Value = "World";
        ws.Cell("A3").Value = 42;

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var loaded = new XLWorkbook(ms);
        var loadedWs = (XLWorksheet)loaded.Worksheets.First();

        // No rows should be in RowsCollection since none have custom properties
        await Assert.That(loadedWs.Internals.RowsCollection).IsEmpty();

        // But cell data should still be accessible
        await Assert.That(loadedWs.Cell("A1")!.GetString()).IsEqualTo("Hello");
        await Assert.That(loadedWs.Cell("A2")!.GetString()).IsEqualTo("World");
        await Assert.That(loadedWs.Cell("A3")!.GetValue<int>()).IsEqualTo(42);
    }

    [Test]
    public async Task LoadingRowsWithCustomHeight_CreatesXLRowObjects()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "Normal row";
        ws.Cell("A2").Value = "Custom height row";
        ws.Row(2).Height = 30;
        ws.Cell("A3").Value = "Normal row";

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var loaded = new XLWorkbook(ms);
        var loadedWs = (XLWorksheet)loaded.Worksheets.First();

        // Only row 2 should be in RowsCollection (it has custom height)
        await Assert.That(loadedWs.Internals.RowsCollection).Count().IsEqualTo(1);
        await Assert.That(loadedWs.Internals.RowsCollection.ContainsKey(2)).IsTrue();
        await Assert.That(loadedWs.Internals.RowsCollection[2].Height).IsEqualTo(30).Within(XLHelper.Epsilon);

        // All cell data should still be accessible
        await Assert.That(loadedWs.Cell("A1")!.GetString()).IsEqualTo("Normal row");
        await Assert.That(loadedWs.Cell("A2")!.GetString()).IsEqualTo("Custom height row");
        await Assert.That(loadedWs.Cell("A3")!.GetString()).IsEqualTo("Normal row");
    }

    [Test]
    public async Task LoadingRowsWithHiddenFlag_CreatesXLRowObjects()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "Visible";
        ws.Cell("A2").Value = "Hidden";
        ws.Row(2).Hide();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var loaded = new XLWorkbook(ms);
        var loadedWs = (XLWorksheet)loaded.Worksheets.First();

        await Assert.That(loadedWs.Internals.RowsCollection.ContainsKey(2)).IsTrue();
        await Assert.That(loadedWs.Row(2).IsHidden).IsTrue();
    }

    [Test]
    public async Task LoadAndSaveRoundTrip_DataOnlyRows_PreservesData()
    {
        // Verify that skipping XLRow creation doesn't break save round-trip
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        for (var i = 1; i <= 100; i++)
        {
            ws.Cell(i, 1).Value = $"Row {i}";
            ws.Cell(i, 2).Value = i * 10;
        }
        // Set custom height on one row
        ws.Row(50).Height = 25;

        using var ms1 = new MemoryStream();
        wb.SaveAs(ms1);
        ms1.Position = 0;

        // Load and re-save
        using var loaded = new XLWorkbook(ms1);
        using var ms2 = new MemoryStream();
        loaded.SaveAs(ms2);
        ms2.Position = 0;

        // Load again and verify
        using var reloaded = new XLWorkbook(ms2);
        var rws = reloaded.Worksheets.First();
        await Assert.That(rws.Cell("A1").GetString()).IsEqualTo("Row 1");
        await Assert.That(rws.Cell("A100").GetString()).IsEqualTo("Row 100");
        await Assert.That(rws.Cell("B50").GetValue<int>()).IsEqualTo(500);
        await Assert.That(rws.Row(50).Height).IsEqualTo(25).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task AdjustToContents_MultilineText_HeightIsLargerThanSingleLine()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Value = "Single line";
        ws.Cell("A2").Value = "Line 1\nLine 2";
        ws.Cell("A3").Value = "Line 1\nLine 2\nLine 3";

        ws.Row(1).AdjustToContents();
        ws.Row(2).AdjustToContents();
        ws.Row(3).AdjustToContents();

        var singleLineHeight = ws.Row(1).Height;
        var twoLineHeight = ws.Row(2).Height;
        var threeLineHeight = ws.Row(3).Height;

        await Assert.That(twoLineHeight).IsGreaterThan(singleLineHeight).Because("Two-line text should produce a taller row than single-line text");
        await Assert.That(threeLineHeight).IsGreaterThan(twoLineHeight).Because("Three-line text should produce a taller row than two-line text");
    }

    [Test]
    public async Task AdjustToContents_ConsecutiveNewlines_EachContributesHeight()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Value = "A\nB";
        ws.Cell("A2").Value = "A\n\nB";

        ws.Row(1).AdjustToContents();
        ws.Row(2).AdjustToContents();

        await Assert.That(ws.Row(2).Height).IsGreaterThan(ws.Row(1).Height).Because("Consecutive newlines (empty line) should add height");
    }

    [Test]
    public async Task AdjustToContents_CrLfNewlines_DetectedSameAsLf()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Value = "Line 1\nLine 2";
        ws.Cell("A2").Value = "Line 1\r\nLine 2";

        ws.Row(1).AdjustToContents();
        ws.Row(2).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsEqualTo(ws.Row(2).Height).Within(XLHelper.Epsilon).Because("LF and CRLF newlines should produce the same row height");
    }

    // The grapheme after a \r\n was skipped (#618). A one-character line after the break, or a
    // second \r\n, is the whole of a line, so losing it changes the height.
    [Test]
    [Arguments("A\nB", "A\r\nB")]
    [Arguments("A\n\nB", "A\r\n\r\nB")]
    [Arguments("A\n\n\nB", "A\r\n\r\n\r\nB")]
    public async Task AdjustToContents_CrLfNewlines_KeepTheLineAfterThem(string lf, string crLf)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = lf;
        ws.Cell("A2").Value = crLf;

        ws.Rows(1, 2).AdjustToContents();

        await Assert.That(ws.Row(2).Height).IsEqualTo(ws.Row(1).Height).Within(XLHelper.Epsilon);
    }

    // Wrapped text is compared with the same text split by hard line breaks, so the tests don't
    // depend on the metrics of a particular font engine. Column A is 6 characters wide: one
    // 3-letter word fits on a line, two don't.

    [Test]
    public async Task AdjustToContents_WrappedText_HeightMatchesHardLineBreaks()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Column(1).Width = 6;
        ws.Cell("A1").Value = "abc def ghi";
        ws.Cell("A1").Style.Alignment.WrapText = true;
        ws.Cell("A2").Value = "abc\ndef\nghi";

        ws.Rows(1, 2).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsEqualTo(ws.Row(2).Height).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task AdjustToContents_WrapTextOff_TextStaysOnOneLine()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Column(1).Width = 6;
        ws.Cell("A1").Value = "abc def ghi";
        ws.Cell("A2").Value = "abc";

        ws.Rows(1, 2).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsEqualTo(ws.Row(2).Height).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task AdjustToContents_WrappedText_BreaksAfterHyphen()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Column(1).Width = 6;
        ws.Cell("A1").Value = "abc-def";
        ws.Cell("A1").Style.Alignment.WrapText = true;
        ws.Cell("A2").Value = "abc-\ndef";

        ws.Rows(1, 2).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsEqualTo(ws.Row(2).Height).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task AdjustToContents_WrappedText_TrailingSpacesDoNotAddLine()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Column(1).Width = 6;
        ws.Cell("A1").Value = "abc                    ";
        ws.Cell("A1").Style.Alignment.WrapText = true;
        ws.Cell("A2").Value = "abc";

        ws.Rows(1, 2).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsEqualTo(ws.Row(2).Height).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task AdjustToContents_WrappedText_OverlongWordIsSplit()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Column(1).Width = 6;
        ws.Cell("A1").Value = "abcdefghijklmnopqrstuvwxyz";
        ws.Cell("A1").Style.Alignment.WrapText = true;
        ws.Cell("A2").Value = "a\nb\nc";

        ws.Rows(1, 2).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsGreaterThanOrEqualTo(ws.Row(2).Height)
            .Because("26 letters in a 6-character column need at least 3 lines");
    }

    [Test]
    public async Task AdjustToContents_WrappedText_IndentNarrowsLine()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Column(1).Width = 10;
        ws.Cell("A1").Value = "abc def";
        ws.Cell("A1").Style.Alignment.WrapText = true;
        ws.Cell("A2").Value = "abc def";
        ws.Cell("A2").Style.Alignment.WrapText = true;
        ws.Cell("A2").Style.Alignment.Indent = 5;
        ws.Cell("A3").Value = "abc\ndef";

        ws.Rows(1, 3).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsLessThan(ws.Row(2).Height).Because("without an indent, both words fit on one line");
        await Assert.That(ws.Row(2).Height).IsEqualTo(ws.Row(3).Height).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task AdjustToContents_WrappedText_ColumnWiderAfterAutofitKeepsOneLine()
    {
        // A column autofitted to a text keeps that text on one line when it is wrapped.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "The quick brown fox jumps over the lazy dog";
        ws.Cell("A1").Style.Alignment.WrapText = true;
        ws.Cell("A2").Value = "The";
        ws.Column(1).AdjustToContents();

        ws.Rows(1, 2).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsEqualTo(ws.Row(2).Height).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task AdjustToContents_Issue2867_WrappedFirstColumnGrowsRow()
    {
        // ClosedXML/ClosedXML#2867: a wrapped cell in a fixed-width column must make the row taller.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Cell("A1").InsertData(new[]
        {
            new object[] { "Cheesecake sablfdsa daskjfhdsakjdsa and what more tging we have...!", 14 },
            new object[] { "Medovik", 6 },
            new object[] { "Muffin", 10 },
        })!;
        var firstColumn = range.FirstColumn()!;
        firstColumn.Style.Alignment.SetWrapText(true);
        firstColumn.WorksheetColumn().Width = 20;
        foreach (var column in ws.ColumnsUsed().Skip(1))
            column.AdjustToContents();
        ws.Cell("D5").Value = "a\nb\nc";

        ws.RowsUsed().AdjustToContents();

        await Assert.That(ws.Column(1).Width).IsEqualTo(20);
        await Assert.That(ws.Row(1).Height).IsGreaterThanOrEqualTo(ws.Row(5).Height)
            .Because("the text needs at least 3 lines in a 20-character column");
        await Assert.That(ws.Row(2).Height).IsLessThan(ws.Row(1).Height);
    }
}
