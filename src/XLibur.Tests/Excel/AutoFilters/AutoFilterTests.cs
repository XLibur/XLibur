using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using XLibur.Excel;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.AutoFilters;

public class AutoFilterTests
{
    [Test]
    public async Task AutoFilterExpandsWithTable()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.FirstCell().SetValue("Categories")
            .CellBelow().SetValue("1")
            .CellBelow().SetValue("2");

        var table = ws.RangeUsed()!.CreateTable();

        var listOfArr = new List<int>
        {
            3,
            4,
            5,
            6
        };

        table.DataRange!.InsertRowsBelow(listOfArr.Count - table.DataRange.RowCount());
        table.DataRange.FirstCell().InsertData(listOfArr);

        await Assert.That(table.AutoFilter.Range.RangeAddress.ToStringRelative()).IsEqualTo("A1:A5");
        await Assert.That(table.AutoFilter.VisibleRows.Count()).IsEqualTo(5);
    }

    [Test]
    public async Task AutoFilterSortWhenNotInFirstRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Cell(3, 3).SetValue("Names")
            .CellBelow().SetValue("Manuel")
            .CellBelow().SetValue("Carlos")
            .CellBelow().SetValue("Dominic");
        ws.RangeUsed()!.SetAutoFilter().Sort();
        await Assert.That(ws.Cell(4, 3).GetText()).IsEqualTo("Carlos");
    }

    [Test]
    public async Task CanClearAutoFilter()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("AutoFilter");
        ws.Cell("A1").Value = "Names";
        ws.Cell("A2").Value = "John";
        ws.Cell("A3").Value = "Hank";
        ws.Cell("A4").Value = "Dagny";

        ws.AutoFilter.Clear(); // We should be able to clear a filter even if it hasn't been set.
        await Assert.That(!ws.AutoFilter.IsEnabled).IsTrue();

        ws.RangeUsed()!.SetAutoFilter();
        await Assert.That(ws.AutoFilter.IsEnabled).IsTrue();

        ws.AutoFilter.Clear();
        await Assert.That(!ws.AutoFilter.IsEnabled).IsTrue();
    }

    [Test]
    public async Task CanClearAutoFilter2()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("AutoFilter");
        ws.Cell("A1").Value = "Names";
        ws.Cell("A2").Value = "John";
        ws.Cell("A3").Value = "Hank";
        ws.Cell("A4").Value = "Dagny";

        ws.SetAutoFilter(false);
        await Assert.That(!ws.AutoFilter.IsEnabled).IsTrue();

        ws.RangeUsed()!.SetAutoFilter();
        await Assert.That(ws.AutoFilter.IsEnabled).IsTrue();

        ws.RangeUsed()!.SetAutoFilter(false);
        await Assert.That(!ws.AutoFilter.IsEnabled).IsTrue();
    }

    [Test]
    public async Task CanCopyAutoFilterToNewSheetOnNewWorkbook()
    {
        using var ms1 = new MemoryStream();
        using var ms2 = new MemoryStream();
        using (var wb1 = new XLWorkbook())
        using (var wb2 = new XLWorkbook())
        {
            var ws = wb1.Worksheets.Add("AutoFilter");
            ws.Cell("A1").Value = "Names";
            ws.Cell("A2").Value = "John";
            ws.Cell("A3").Value = "Hank";
            ws.Cell("A4").Value = "Dagny";

            ws.RangeUsed()!.SetAutoFilter();

            wb1.SaveAs(ms1);

            ws.CopyTo(wb2, ws.Name);
            wb2.SaveAs(ms2);
        }

        using (var wb2 = new XLWorkbook(ms2))
        {
            await Assert.That(wb2.Worksheets.First().AutoFilter.IsEnabled).IsTrue();
        }
    }

    [Test]
    public async Task CannotAddAutoFilterOverExistingTable()
    {
        using var wb = new XLWorkbook();

        var data = Enumerable.Range(1, 10).Select(i => new
        {
            Index = i,
            String = $"String {i}"
        });

        var ws = wb.AddWorksheet();
        ws.FirstCell().InsertTable(data);

        await Assert.That(() => ws.RangeUsed()!.SetAutoFilter()).Throws<InvalidOperationException>();
    }

    [Test]
    [Arguments("A1:A4")]
    [Arguments("A1:B4")]
    [Arguments("A1:C4")]
    public async Task AutoFilterRangeRemainsValidOnInsertColumn(string rangeAddress)
    {
        // Arrange
        using var ms1 = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("AutoFilter");
        ws.Cell("A1").Value = "Ids";
        ws.Cell("B1").Value = "Names";
        ws.Cell("B2").Value = "John";
        ws.Cell("B3").Value = "Hank";
        ws.Cell("B4").Value = "Dagny";
        ws.Cell("C1").Value = "Phones";

        ws.Range("B1:B4").SetAutoFilter(true);

        // Act
        var range = ws.Range(rangeAddress);
        range.InsertColumnsBefore(1);

        // Assert
        await Assert.That(ws.AutoFilter.Range.RangeAddress.IsValid).IsTrue();
    }

    [Test]
    public async Task AutoFilterVisibleRows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Cell(3, 3).SetValue("Names")
            .CellBelow().SetValue("Manuel")
            .CellBelow().SetValue("Carlos")
            .CellBelow().SetValue("Dominic");

        var autoFilter = ws.RangeUsed()!.SetAutoFilter();

        autoFilter.Column(1).AddFilter("Carlos");

        await Assert.That(ws.Cell(5, 3).GetText()).IsEqualTo("Carlos");
        await Assert.That(autoFilter.VisibleRows.Count()).IsEqualTo(2);
        await Assert.That(autoFilter.VisibleRows.First().WorksheetRow().RowNumber()).IsEqualTo(3);
        await Assert.That(autoFilter.VisibleRows.Last().WorksheetRow().RowNumber()).IsEqualTo(5);
    }

    [Test]
    public async Task ReapplyAutoFilter()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Cell(3, 3).SetValue("Names")
            .CellBelow().SetValue("Manuel")
            .CellBelow().SetValue("Carlos")
            .CellBelow().SetValue("Dominic")
            .CellBelow().SetValue("Jose");

        var autoFilter = ws.RangeUsed()!
            .SetAutoFilter();

        autoFilter.Column(1).AddFilter("Carlos");

        await Assert.That(autoFilter.HiddenRows.Count()).IsEqualTo(3);

        // Unhide the rows so that the table is out of sync with the filter
        autoFilter.HiddenRows.ForEach(r => r.WorksheetRow().Unhide());
        await Assert.That(autoFilter.HiddenRows.Any()).IsFalse();

        autoFilter.Reapply();
        await Assert.That(autoFilter.HiddenRows.Count()).IsEqualTo(3);
    }

    [Test]
    public async Task CanLoadAutoFilterWithThousandsSeparator()
    {
        var backupCulture = Thread.CurrentThread.CurrentCulture;

        try
        {
            // Set thread culture to French, which should format numbers using a space as thousand's separator
            var culture = CultureInfo.CreateSpecificCulture("fr-FR");

            // The value in the sheet that will be compared with autofilter value is a number
            // `10000`. That number will be formatted using culture to `10 000.00` thanks to
            // modified properties of culture - period instead of a comma for decimal separator
            // and space as group separator. The formatted number will thus match with the
            // filter value.
            culture.NumberFormat.NumberDecimalSeparator = ".";
            culture.NumberFormat.NumberGroupSeparator = " ";

            Thread.CurrentThread.CurrentCulture = culture;

            using (var stream =
                   TestHelper.GetStreamFromResource(
                       TestHelper.GetResourcePath(@"Other\AutoFilter\AutoFilterWithThousandsSeparator.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();

                // Regular filter compares values as strings, doesn't convert to XLCellValue,
                // so the value is read from the file as a text despite looking like a number.
                await Assert.That(((XLAutoFilter)ws.AutoFilter).Column(1).Single().Value).IsEqualTo("10 000.00");
                await Assert.That(ws.AutoFilter.VisibleRows.Count()).IsEqualTo(2);

                ws.AutoFilter.Reapply();
                await Assert.That(ws.AutoFilter.VisibleRows.Count()).IsEqualTo(2);
            }

            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");

            using (var stream =
                   TestHelper.GetStreamFromResource(
                       TestHelper.GetResourcePath(@"Other\AutoFilter\AutoFilterWithThousandsSeparator.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                await Assert.That(((XLAutoFilter)ws.AutoFilter).Column(1).Single().Value).IsEqualTo("10 000.00");

                var unused = ws.AutoFilter.VisibleRows.Select(r => r.FirstCell().Value).ToList();
                await Assert.That(ws.AutoFilter.VisibleRows.Count()).IsEqualTo(2);

                ws.AutoFilter.Reapply();
                await Assert.That(ws.AutoFilter.VisibleRows.Count()).IsEqualTo(1);
            }
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = backupCulture;
        }
    }

    [Test]
    public async Task Issue1917NotContainsFilter()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Test");
            ws.Cell(1, 1).SetValue("StringCol");

            for (var i = 0; i < 5; i++)
            {
                ws.Cell(i + 2, 1).SetValue($"String{i}");
            }

            var autoFilter = ws.RangeUsed()!
                .SetAutoFilter();

            autoFilter.Column(1).NotContains("String3");
            await Assert.That(autoFilter.HiddenRows.Count()).IsEqualTo(1);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.Worksheet("Test");
            var autoFilter = ws.AutoFilter;

            autoFilter.Reapply();
            await Assert.That(autoFilter.HiddenRows.Count()).IsEqualTo(1);
        }
    }

    [Test]
    [Arguments("ends")]
    [Arguments("begins")]
    [Arguments("equal")]
    [Arguments("contains")]
    public async Task NotStringFilter(string type)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Test");
            ws.Cell(1, 1).SetValue("StringCol");

            for (var i = 0; i < 5; i++)
            {
                ws.Cell(i + 2, 1).SetValue($"{i}-String{i}");
            }

            ws.Columns().AdjustToContents();
            var autoFilter = ws.RangeUsed()!
                .SetAutoFilter();

            switch (type)
            {
                case "ends":
                    autoFilter.Column(1).NotEndsWith("3");
                    break;
                case "begins":
                    autoFilter.Column(1).NotBeginsWith("3");
                    break;
                case "equal":
                    autoFilter.Column(1).NotEqualTo("3-String3");
                    break;
                case "contains":
                    autoFilter.Column(1).NotContains("3-");
                    break;
            }

            await Assert.That(autoFilter.HiddenRows.Count()).IsEqualTo(1);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.Worksheet("Test");
            var autoFilter = ws.AutoFilter;

            autoFilter.Reapply();
            await Assert.That(autoFilter.HiddenRows.Count()).IsEqualTo(1);
        }
    }

    [Test]
    public async Task AutoFilterReapplyShouldNotThrowNullReferenceError()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("Test");

        await Assert.That(() => { sheet.AutoFilter.Reapply(); }).ThrowsNothing();
    }

    [Test]
    public async Task ReapplyExpandsRangeToIncludeNewDataRows()
    {
        // Bug #2812: When new rows are added below the autofilter range and
        // Reapply() is called, the range should expand to include the new data
        // and the filter should be applied to those rows too.
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Cell("A1").SetValue("Number");
        ws.Cell("B1").SetValue("Name");
        ws.Cell("A2").SetValue(34);
        ws.Cell("B2").SetValue("Alice");
        ws.Cell("A3").SetValue(50);
        ws.Cell("B3").SetValue("Bob");

        // Set autofilter on A1:B3 and filter to show only value 34
        var autoFilter = ws.Range("A1:B3").SetAutoFilter();
        autoFilter.Column(1).AddFilter("34");

        // Verify initial state: only row with 34 visible (+ header)
        await Assert.That(autoFilter.HiddenRows.Count()).IsEqualTo(1);
        await Assert.That(autoFilter.Range.RangeAddress.ToStringRelative()).IsEqualTo("A1:B3");

        // Add a new row beyond the filter range
        ws.Cell("A4").SetValue(35);
        ws.Cell("B4").SetValue("Charlie");

        // Reapply should expand range and hide the new row (35 != 34)
        autoFilter.Reapply();

        await Assert.That(autoFilter.Range.RangeAddress.ToStringRelative()).IsEqualTo("A1:B4");
        await Assert.That(autoFilter.HiddenRows.Count()).IsEqualTo(2);
        // Visible = header row + the row matching filter "34"
        await Assert.That(autoFilter.VisibleRows.Count()).IsEqualTo(2);
    }

    [Test]
    public async Task ReapplyDoesNotExpandRangeAcrossGap()
    {
        // Range should only expand to contiguous data, not jump over empty rows.
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Cell("A1").SetValue("Number");
        ws.Cell("A2").SetValue(34);
        ws.Cell("A3").SetValue(50);

        var autoFilter = ws.Range("A1:A3").SetAutoFilter();
        autoFilter.Column(1).AddFilter("34");

        // Add data with a gap (row 4 empty, row 5 has data)
        ws.Cell("A5").SetValue(99);

        autoFilter.Reapply();

        // Range should NOT expand past the empty row
        await Assert.That(autoFilter.Range.RangeAddress.ToStringRelative()).IsEqualTo("A1:A3");
    }

    [Test]
    public async Task ReapplyExpandsRangeAndFiltersFromLoadedFile()
    {
        // Test with the actual bug report file
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"Other\AutoFilter\autofilter_bug_2812.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();

        var originalRange = ws.AutoFilter.Range.RangeAddress.ToStringRelative();

        // Add a new row of data below the filter range
        var lastRow = ws.AutoFilter.Range.RangeAddress.LastAddress.RowNumber;
        var newRow = lastRow + 1;
        ws.Cell(newRow, 1).SetValue(35);
        ws.Cell(newRow, 2).SetValue("test");
        ws.Cell(newRow, 3).SetValue("test2");

        ws.AutoFilter.Reapply();

        // Range should have expanded
        var newRange = ws.AutoFilter.Range.RangeAddress.ToStringRelative();
        await Assert.That(newRange).IsNotEqualTo(originalRange);

        // The last row of the range should now include the new data
        await Assert.That(ws.AutoFilter.Range.RangeAddress.LastAddress.RowNumber).IsEqualTo(newRow);
    }

    [Test]
    public async Task SaveAutoFilterWithClearedColumnDoesNotThrow()
    {
        // When a filter column is added and then cleared, it remains in the
        // internal dictionary with FilterType.None. Saving should skip it
        // rather than throwing NotSupportedException.
        using var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Test");

        ws.Cell("A1").SetValue("Header");
        ws.Cell("A2").SetValue("Value1");
        ws.Cell("A3").SetValue("Value2");

        var autoFilter = ws.RangeUsed()!.SetAutoFilter();
        autoFilter.Column(1).AddFilter("Value1");
        autoFilter.Column(1).Clear();

        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }
}
