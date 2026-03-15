using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.Rows;

[TestFixture]
public class RowTests
{
    [Test]
    public void RowsUsedIsFast()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().SetValue("Hello world!");
        var rowsUsed = ws.Column(1).AsRange().RowsUsed();
        Assert.AreEqual(1, rowsUsed.Count());
    }

    [Test]
    public void CopyRow()
    {
        var wb = new XLWorkbook();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue("Test").Style.Font.SetBold();
        ws.FirstRow().CopyTo(ws.Row(2));

        Assert.IsTrue(ws.Cell("A2").Style.Font.Bold);
    }

    [Test]
    public void InsertingRowsAbove1()
    {
        var wb = new XLWorkbook();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

        ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
        ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

        IXLRow row1 = ws.Row(1);
        IXLRow row2 = ws.Row(2);
        IXLRow row3 = ws.Row(3);

        IXLRow rowIns = ws.Row(1).InsertRowsAbove(1).First();

        Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(ws.Style.Fill.BackgroundColor, ws.Row(1).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Green, ws.Row(3).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual("X", ws.Row(3).Cell(2).GetText());

        Assert.AreEqual(ws.Style.Fill.BackgroundColor, rowIns.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(ws.Style.Fill.BackgroundColor, rowIns.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(ws.Style.Fill.BackgroundColor, rowIns.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, row1.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row1.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row1.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Yellow, row2.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Green, row2.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Yellow, row2.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, row3.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row3.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row3.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual("X", row2.Cell(2).GetText());
    }

    [Test]
    public void InsertingRowsAbove2()
    {
        var wb = new XLWorkbook();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

        ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
        ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

        IXLRow row1 = ws.Row(1);
        IXLRow row2 = ws.Row(2);
        IXLRow row3 = ws.Row(3);

        IXLRow rowIns = ws.Row(2).InsertRowsAbove(1).First();

        Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(2).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Green, ws.Row(3).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual("X", ws.Row(3).Cell(2).GetText());

        Assert.AreEqual(XLColor.Red, rowIns.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, rowIns.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, rowIns.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, row1.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row1.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row1.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Yellow, row2.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Green, row2.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Yellow, row2.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, row3.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row3.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row3.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual("X", row2.Cell(2).GetText());
    }

    [Test]
    public void InsertingRowsAbove3()
    {
        var wb = new XLWorkbook();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

        ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
        ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

        IXLRow row1 = ws.Row(1);
        IXLRow row2 = ws.Row(2);
        IXLRow row3 = ws.Row(3);

        IXLRow rowIns = ws.Row(3).InsertRowsAbove(1).First();

        Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(1).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Yellow, ws.Row(2).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Green, ws.Row(2).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Yellow, ws.Row(2).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Green, ws.Row(3).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Yellow, ws.Row(3).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, ws.Row(4).Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual("X", ws.Row(2).Cell(2).GetText());

        Assert.AreEqual(XLColor.Yellow, rowIns.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Green, rowIns.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Yellow, rowIns.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, row1.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row1.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row1.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Yellow, row2.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Green, row2.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Yellow, row2.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual(XLColor.Red, row3.Cell(1).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row3.Cell(2).Style.Fill.BackgroundColor);
        Assert.AreEqual(XLColor.Red, row3.Cell(3).Style.Fill.BackgroundColor);

        Assert.AreEqual("X", row2.Cell(2).GetText());
    }

    [Test]
    public void InsertingRowsAbove4()
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

        Assert.AreEqual(15, ws.Row(2).Height);
        Assert.AreEqual(20, ws.Row(4).Height);
        Assert.AreEqual(25, ws.Row(5).Height);
        Assert.AreEqual(35, ws.Row(6).Height);

        Assert.AreEqual(20, ws.Row(3).Height);
        ws.Row(3).ClearHeight();
        Assert.AreEqual(ws.RowHeight, ws.Row(3).Height);
    }

    [Test]
    public void NoRowsUsed()
    {
        var wb = new XLWorkbook();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        Int32 count = 0;

        foreach (IXLRow row in ws.RowsUsed())
            count++;

        foreach (IXLRangeRow row in ws.Range("A1:C3").RowsUsed())
            count++;

        Assert.AreEqual(0, count);
    }

    [Test]
    public void RowUsed()
    {
        var wb = new XLWorkbook();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 2).SetValue("Test");
        ws.Cell(1, 3).SetValue("Test");

        IXLRangeRow fromRow = ws.Row(1).RowUsed();
        Assert.AreEqual("B1:C1", fromRow.RangeAddress.ToStringRelative());

        IXLRangeRow fromRange = ws.Range("A1:E1").FirstRow().RowUsed();
        Assert.AreEqual("B1:C1", fromRange.RangeAddress.ToStringRelative());
    }

    [Test]
    public void RowsUsedWithDataValidation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().SetValue("Hello world!");
        ws.Range("A1:A100").CreateDataValidation().WholeNumber.EqualTo(1);

        var range = ws.Column(1).AsRange();

        Assert.AreEqual(100, range.RowsUsed(XLCellsUsedOptions.DataValidation).Count());
        Assert.AreEqual(100, range.RowsUsed(XLCellsUsedOptions.All).Count());
    }

    [Test]
    public void RowsUsedWithConditionalFormatting()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().SetValue("Hello world!");
        ws.Range("A1:A100").AddConditionalFormat().WhenStartsWith("Hell").Fill.SetBackgroundColor(XLColor.Red).Font.SetFontColor(XLColor.White);

        var range = ws.Column(1).AsRange();

        Assert.AreEqual(100, range.RowsUsed(XLCellsUsedOptions.ConditionalFormats).Count());
        Assert.AreEqual(100, range.RowsUsed(XLCellsUsedOptions.All).Count());
    }

    [Test]
    public void UngroupFromAll()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet("Sheet1");
        ws.Rows(1, 2).Group();
        ws.Rows(1, 2).Ungroup(true);

        Assert.AreEqual(0, ws.Row(1).OutlineLevel);
        Assert.AreEqual(0, ws.Row(2).OutlineLevel);
    }

    [Test]
    public void NegativeRowNumberIsInvalid()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1") as XLWorksheet;

        var row = new XLRow(ws, -1);

        Assert.IsFalse(row.RangeAddress.IsValid);
    }

    [Test]
    public void DeleteRowOnWorksheetWithComment()
    {
        var ws = new XLWorkbook().AddWorksheet();
        ws.Cell(4, 1).GetComment().AddText("test");
        ws.Column(1).Width = 100;
        Assert.DoesNotThrow(() => ws.Row(1).Delete());
    }

    [Test]
    public void AssignWorksheetRowHeightWhenAllRowsChanged()
    {
        var ws = new XLWorkbook().AddWorksheet();
        var rows = ws.Rows();

        rows.Height = 30;

        Assert.AreEqual(30, ws.Row(11).Height, XLHelper.Epsilon);
        Assert.AreEqual(30, ws.RowHeight, XLHelper.Epsilon);
    }

    [Test]
    public void PreserveWorksheetRowHeightWhenNotAllRowsChanged()
    {
        var ws = new XLWorkbook().AddWorksheet();
        var defaultRowHeight = ws.RowHeight;
        var rows = ws.Rows(1, XLHelper.MaxRowNumber);

        rows.Height = 30;

        Assert.AreEqual(30, ws.Row(11).Height, XLHelper.Epsilon);
        Assert.AreEqual(defaultRowHeight, ws.RowHeight, XLHelper.Epsilon);
    }

    [Test]
    public void PreserveWorksheetRowHeightWhenUsedRowsChanged()
    {
        var ws = new XLWorkbook().AddWorksheet();
        ws.Cells("A1:E5").Value = "Not empty";
        var defaultRowHeight = ws.RowHeight;
        var rows = ws.RowsUsed(XLCellsUsedOptions.Contents);

        rows.Height = 30;

        Assert.AreEqual(30, ws.Row(3).Height, XLHelper.Epsilon);
        Assert.AreEqual(defaultRowHeight, ws.Row(11).Height, XLHelper.Epsilon);
        Assert.AreEqual(defaultRowHeight, ws.RowHeight, XLHelper.Epsilon);
    }

    [Test]
    public void LoadingDataOnlyRows_DoesNotCreateXLRowObjects()
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
        Assert.That(loadedWs.Internals.RowsCollection, Is.Empty);

        // But cell data should still be accessible
        Assert.That(loadedWs.Cell("A1").GetString(), Is.EqualTo("Hello"));
        Assert.That(loadedWs.Cell("A2").GetString(), Is.EqualTo("World"));
        Assert.That(loadedWs.Cell("A3").GetValue<int>(), Is.EqualTo(42));
    }

    [Test]
    public void LoadingRowsWithCustomHeight_CreatesXLRowObjects()
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
        Assert.That(loadedWs.Internals.RowsCollection, Has.Count.EqualTo(1));
        Assert.That(loadedWs.Internals.RowsCollection.ContainsKey(2), Is.True);
        Assert.That(loadedWs.Internals.RowsCollection[2].Height, Is.EqualTo(30).Within(XLHelper.Epsilon));

        // All cell data should still be accessible
        Assert.That(loadedWs.Cell("A1").GetString(), Is.EqualTo("Normal row"));
        Assert.That(loadedWs.Cell("A2").GetString(), Is.EqualTo("Custom height row"));
        Assert.That(loadedWs.Cell("A3").GetString(), Is.EqualTo("Normal row"));
    }

    [Test]
    public void LoadingRowsWithHiddenFlag_CreatesXLRowObjects()
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

        Assert.That(loadedWs.Internals.RowsCollection.ContainsKey(2), Is.True);
        Assert.That(loadedWs.Row(2).IsHidden, Is.True);
    }

    [Test]
    public void LoadAndSaveRoundTrip_DataOnlyRows_PreservesData()
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
        Assert.That(rws.Cell("A1").GetString(), Is.EqualTo("Row 1"));
        Assert.That(rws.Cell("A100").GetString(), Is.EqualTo("Row 100"));
        Assert.That(rws.Cell("B50").GetValue<int>(), Is.EqualTo(500));
        Assert.That(rws.Row(50).Height, Is.EqualTo(25).Within(XLHelper.Epsilon));
    }
}
