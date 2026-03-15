using XLibur.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLibur.Tests.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Tests;

[TestFixture]
public class XLPivotTableTests
{
    [Test]
    public void PivotTables()
    {
        Assert.DoesNotThrow(() =>
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("PastrySalesData");
            var table = ws.Table("PastrySalesData");
            var ptSheet = wb.Worksheets.Add("BlankPivotTable");
            ptSheet.PivotTables.Add("pvt", ptSheet.Cell(1, 1), table);

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);
        });
    }

    [Test]
    public void TestPivotTableVersioningAttributes()
    {
        // Pivot cache definitions in input file has created and refreshed version attributes = 3
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\VersioningAttributes\inputfile.xlsx"));
        TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook(stream);

            var data = wb.Worksheet("Data");

            var pt = data.RangeUsed().CreatePivotTable(wb.AddWorksheet("pvt2").FirstCell(), "pvt2");

            pt.ColumnLabels.Add("Sex");
            pt.RowLabels.Add("FullName");
            pt.Values.Add("Id", "Count of Id").SetSummaryFormula(XLPivotSummary.Count);

            return wb;
            // Pivot cache definitions in output file has created and refreshed version attributes = 5
        }, @"Other\PivotTableReferenceFiles\VersioningAttributes\outputfile.xlsx");
    }

    [Test]
    public void PivotTableOptionsSaveTest()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("PastrySalesData");
        var table = ws.Table("PastrySalesData");
        var ptSheet = wb.Worksheets.Add("BlankPivotTable");
        var pt = ptSheet.PivotTables.Add("pvtOptionsTest", ptSheet.Cell(1, 1), table);

        pt.ColumnHeaderCaption = "clmn header";
        pt.RowHeaderCaption = "row header";

        pt.AutofitColumns = true;
        pt.PreserveCellFormatting = false;
        pt.ShowGrandTotalsColumns = true;
        pt.ShowGrandTotalsRows = true;
        pt.UseCustomListsForSorting = false;
        pt.ShowExpandCollapseButtons = false;
        pt.ShowContextualTooltips = false;
        pt.DisplayCaptionsAndDropdowns = false;
        pt.RepeatRowLabels = true;
        pt.PivotCache.SaveSourceData = false;
        pt.EnableShowDetails = false;
        pt.ShowColumnHeaders = false;
        pt.ShowRowHeaders = false;

        pt.MergeAndCenterWithLabels = true; // MergeItem
        pt.RowLabelIndent = 12; // Indent
        pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown; // PageOverThenDown
        pt.FilterFieldsPageWrap = 14; // PageWrap
        pt.ErrorValueReplacement = "error test"; // ErrorCaption
        pt.EmptyCellReplacement = "empty test"; // MissingCaption

        pt.FilteredItemsInSubtotals = true; // Subtotal filtered page items
        pt.AllowMultipleFilters = false; // MultipleFieldFilters

        pt.ShowPropertiesInTooltips = false;
        pt.ClassicPivotTableLayout = true;
        pt.ShowEmptyItemsOnRows = true;
        pt.ShowEmptyItemsOnColumns = true;
        pt.DisplayItemLabels = false;
        pt.SortFieldsAtoZ = true;

        pt.PrintExpandCollapsedButtons = true;
        pt.PrintTitles = true;

        pt.PivotCache.RefreshDataOnOpen = false;
        pt.PivotCache.ItemsToRetainPerField = XLItemsToRetain.Max;
        pt.EnableCellEditing = true;
        pt.ShowValuesRow = true;
        pt.ShowRowStripes = true;
        pt.ShowColumnStripes = true;
        pt.Theme = XLPivotTableTheme.PivotStyleDark13;

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;

        using var wbassert = new XLWorkbook(ms);
        var wsassert = wbassert.Worksheet("BlankPivotTable");
        var ptassert = wsassert.PivotTable("pvtOptionsTest");
        Assert.AreNotEqual(null, ptassert, "name save failure");
        Assert.AreEqual("clmn header", ptassert.ColumnHeaderCaption, "ColumnHeaderCaption save failure");
        Assert.AreEqual("row header", ptassert.RowHeaderCaption, "RowHeaderCaption save failure");
        Assert.AreEqual(true, ptassert.MergeAndCenterWithLabels, "MergeAndCenterWithLabels save failure");
        Assert.AreEqual(12, ptassert.RowLabelIndent, "RowLabelIndent save failure");
        Assert.AreEqual(XLFilterAreaOrder.OverThenDown, ptassert.FilterAreaOrder, "FilterAreaOrder save failure");
        Assert.AreEqual(14, ptassert.FilterFieldsPageWrap, "FilterFieldsPageWrap save failure");
        Assert.AreEqual("error test", ptassert.ErrorValueReplacement, "ErrorValueReplacement save failure");
        Assert.AreEqual("empty test", ptassert.EmptyCellReplacement, "EmptyCellReplacement save failure");
        Assert.AreEqual(true, ptassert.AutofitColumns, "AutofitColumns save failure");
        Assert.AreEqual(false, ptassert.PreserveCellFormatting, "PreserveCellFormatting save failure");
        Assert.AreEqual(true, ptassert.ShowGrandTotalsRows, "ShowGrandTotalsRows save failure");
        Assert.AreEqual(true, ptassert.ShowGrandTotalsColumns, "ShowGrandTotalsColumns save failure");
        Assert.AreEqual(true, ptassert.FilteredItemsInSubtotals, "FilteredItemsInSubtotals save failure");
        Assert.AreEqual(false, ptassert.AllowMultipleFilters, "AllowMultipleFilters save failure");
        Assert.AreEqual(false, ptassert.UseCustomListsForSorting, "UseCustomListsForSorting save failure");
        Assert.AreEqual(false, ptassert.ShowExpandCollapseButtons, "ShowExpandCollapseButtons save failure");
        Assert.AreEqual(false, ptassert.ShowContextualTooltips, "ShowContextualTooltips save failure");
        Assert.AreEqual(false, ptassert.ShowPropertiesInTooltips, "ShowPropertiesInTooltips save failure");
        Assert.AreEqual(false, ptassert.DisplayCaptionsAndDropdowns, "DisplayCaptionsAndDropdowns save failure");
        Assert.AreEqual(true, ptassert.ClassicPivotTableLayout, "ClassicPivotTableLayout save failure");
        Assert.AreEqual(true, ptassert.ShowEmptyItemsOnRows, "ShowEmptyItemsOnRows save failure");
        Assert.AreEqual(true, ptassert.ShowEmptyItemsOnColumns, "ShowEmptyItemsOnColumns save failure");
        Assert.AreEqual(false, ptassert.DisplayItemLabels, "DisplayItemLabels save failure");
        Assert.AreEqual(true, ptassert.SortFieldsAtoZ, "SortFieldsAtoZ save failure");
        Assert.AreEqual(true, ptassert.PrintExpandCollapsedButtons, "PrintExpandCollapsedButtons save failure");
        Assert.AreEqual(true, ptassert.RepeatRowLabels, "RepeatRowLabels save failure");
        Assert.AreEqual(true, ptassert.PrintTitles, "PrintTitles save failure");
        Assert.AreEqual(false, ptassert.PivotCache.SaveSourceData, "SaveSourceData save failure");
        Assert.AreEqual(false, ptassert.EnableShowDetails, "EnableShowDetails save failure");
        Assert.AreEqual(false, ptassert.PivotCache.RefreshDataOnOpen, "RefreshDataOnOpen save failure");
        Assert.AreEqual(XLItemsToRetain.Max, ptassert.PivotCache.ItemsToRetainPerField, "ItemsToRetainPerField save failure");
        Assert.AreEqual(true, ptassert.EnableCellEditing, "EnableCellEditing save failure");
        Assert.AreEqual(XLPivotTableTheme.PivotStyleDark13, ptassert.Theme, "Theme save failure");
        Assert.AreEqual(true, ptassert.ShowValuesRow, "ShowValuesRow save failure");
        Assert.AreEqual(false, ptassert.ShowRowHeaders, "ShowRowHeaders save failure");
        Assert.AreEqual(false, ptassert.ShowColumnHeaders, "ShowColumnHeaders save failure");
        Assert.AreEqual(true, ptassert.ShowRowStripes, "ShowRowStripes save failure");
        Assert.AreEqual(true, ptassert.ShowColumnStripes, "ShowColumnStripes save failure");
    }

    [TestCase(true)]
    [TestCase(false)]
    public void PivotFieldOptionsSaveTest(bool withDefaults)
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("PastrySalesData");
        var table = ws.Table("PastrySalesData");

        var ptSheet = wb.Worksheets.Add("pvtFieldOptionsTest");
        var pt = ptSheet.PivotTables.Add("pvtFieldOptionsTest", ptSheet.Cell(1, 1), table);

        var field = pt.RowLabels.Add("Name")
            .SetSubtotalCaption("Test caption")
            .SetCustomName("Test name");
        SetFieldOptions(field, withDefaults);

        pt.ColumnLabels.Add("Month");
        pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

        using var ms = new MemoryStream();
        wb.SaveAs(ms, true);

        ms.Position = 0;

        using var wbassert = new XLWorkbook(ms);
        var wsassert = wbassert.Worksheet("pvtFieldOptionsTest");
        var ptassert = wsassert.PivotTable("pvtFieldOptionsTest");
        var pfassert = ptassert.RowLabels.Get("Name");
        Assert.AreNotEqual(null, pfassert, "name save failure");
        Assert.AreEqual("Test caption", pfassert.SubtotalCaption, "SubtotalCaption save failure");
        Assert.AreEqual("Test name", pfassert.CustomName, "CustomName save failure");
        AssertFieldOptions(pfassert, withDefaults);
    }

    [Test]
    [Ignore("PT styles will be fixed in a different PR")]
    public void PivotTableStyleFormatsTest()
    {
        /*
                    using (var ms = new MemoryStream())
                    {
                        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
                        using (var wbSource = new XLWorkbook(stream))
                        using (var wbDestination = new XLWorkbook())
                        {
                            var ws = wbSource.Worksheet("PastrySalesData");
                            wbDestination.AddWorksheet(ws);
                            ws = wbDestination.Worksheet("PastrySalesData");

                            var table = ws.Table("PastrySalesData");
                            var ptSheet = wbDestination.Worksheets.Add("PivotTableStyleFormats");
                            var pt = ptSheet.PivotTables.Add("pvtStyleFormats", ptSheet.Cell(1, 1), table);
                            pt.Layout = XLPivotLayout.Tabular;

                            pt.SetSubtotals(XLPivotSubtotals.AtBottom);

                            var monthPivotField = pt.ColumnLabels.Add("Month");

                            var namePivotField = pt.RowLabels.Add("Name")
                                .SetSubtotalCaption("Test caption")
                                .SetCustomName("Test name")
                                .AddSubtotal(XLSubtotalFunction.Sum);

                            ptSheet.SetTabActive();

                            var numberOfOrdersPivotValue = pt.Values.Add("NumberOfOrders")
                                .SetSummaryFormula(XLPivotSummary.Sum);

                            var qualityPivotValue = pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);

                            pt.StyleFormats.RowGrandTotalFormats.ForElement(XLPivotStyleFormatElement.All).Style.Font.FontColor = XLColor.VenetianRed;

                            namePivotField.StyleFormats.Subtotal.Style.Fill.BackgroundColor = XLColor.Blue;
                            monthPivotField.StyleFormats.Label.Style.Fill.BackgroundColor = XLColor.Amber;
                            monthPivotField.StyleFormats.Header.Style.Font.FontColor = XLColor.Yellow;
                            namePivotField.StyleFormats.DataValuesFormat
                                .AndWith(monthPivotField, v => v.IsText && v.GetText() == "May")
                                .ForValueField(numberOfOrdersPivotValue)
                                .Style.Font.FontColor = XLColor.Green;

                            wbDestination.SaveAs(ms);
                        }

                        ms.Seek(0, SeekOrigin.Begin);

                        using (var wb = new XLWorkbook(ms))
                        {
                            var ws = wb.Worksheet("PivotTableStyleFormats");
                            var pt = ws.PivotTable("pvtStyleFormats").CastTo<XLPivotTable>();

                            Assert.AreEqual(0, pt.StyleFormats.ColumnGrandTotalFormats.Count());

                            Assert.NotNull(pt.StyleFormats.RowGrandTotalFormats);
                            Assert.AreEqual(1, pt.StyleFormats.RowGrandTotalFormats.Count());
                            Assert.AreEqual(XLPivotStyleFormatElement.All, pt.StyleFormats.RowGrandTotalFormats.First().AppliesTo);
                            Assert.AreEqual(XLColor.VenetianRed, pt.StyleFormats.RowGrandTotalFormats.ForElement(XLPivotStyleFormatElement.All).Style.Font.FontColor);

                            var namePivotField = pt.RowLabels.Get("Name");
                            var monthPivotField = pt.ColumnLabels.Get("Month");
                            var numberOfOrdersPivotValue = pt.Values.Get("NumberOfOrders");

                            Assert.AreEqual(XLStyle.Default, namePivotField.StyleFormats.Label.Style);
                            Assert.AreEqual(XLColor.Blue, namePivotField.StyleFormats.Subtotal.Style.Fill.BackgroundColor);

                            Assert.AreEqual(XLStyle.Default, monthPivotField.StyleFormats.Subtotal.Style);
                            Assert.AreEqual(XLColor.Amber, monthPivotField.StyleFormats.Label.Style.Fill.BackgroundColor);
                            Assert.AreEqual(XLColor.Yellow, monthPivotField.StyleFormats.Header.Style.Font.FontColor);

                            var nameDataValuesFormat = namePivotField.StyleFormats.DataValuesFormat as XLPivotValueStyleFormat;
                            Assert.AreEqual(2, nameDataValuesFormat.FieldReferences.Count());

                            Assert.AreEqual(monthPivotField, nameDataValuesFormat.FieldReferences.First().CastTo<PivotLabelFieldReference>().PivotField);

                            Assert.AreEqual(numberOfOrdersPivotValue.CustomName, nameDataValuesFormat.FieldReferences.Last().CastTo<PivotValueFieldReference>().Value);

                            wb.Save();
                        }
                    }
        */
    }

    [Test]
    public void CopyPivotTableTests()
    {
        using var ms = new MemoryStream();
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws1 = wb.Worksheet("pvt1");
        var pt1 = ws1.PivotTables.First() as XLPivotTable;

        Assert.Throws<InvalidOperationException>(() => pt1.CopyTo(pt1.TargetCell));

        var pt2 = pt1.CopyTo(ws1.Cell("AB100")) as XLPivotTable;

        AssertPivotTablesAreEqual(pt1, pt2, compareName: false);

        var ws2 = wb.AddWorksheet("Copy Of pvt1");
        AssertPivotTablesAreEqual(pt1, pt1.CopyTo(ws2.FirstCell()) as XLPivotTable, compareName: true);

        using var wb2 = new XLWorkbook();
        wb.Worksheet("PastrySalesData").CopyTo(wb2);

        AssertPivotTablesAreEqual(pt1, pt1.CopyTo(wb2.AddWorksheet("pvt").FirstCell()) as XLPivotTable, compareName: true);
    }

    private void AssertPivotTablesAreEqual(XLPivotTable original, XLPivotTable copy, Boolean compareName)
    {
        Assert.AreEqual(compareName, original.Name.Equals(copy.Name));

        var comparer = new PivotTableComparer(compareName: compareName, compareRelId: false, compareTargetCellAddress: false);
        Assert.IsTrue(comparer.Equals(original, copy));
    }

    private class Pastry
    {
        public Pastry(string name, int? code, int numberOfOrders, double quality, string month, DateTime? bakeDate)
        {
            Name = name;
            Code = code;
            NumberOfOrders = numberOfOrders;
            Quality = quality;
            Month = month;
            BakeDate = bakeDate;
        }

        public string Name { get; set; }
        public int? Code { get; }
        public int NumberOfOrders { get; set; }
        public double Quality { get; set; }
        public string Month { get; set; }
        public DateTime? BakeDate { get; set; }
    }

    [Test]
    public void SharedItemsWithVariousDataTypesInTableColumn()
    {
        // Load an excel that contains a table which has various combinations of types in columns.
        // The pivot cache definition contain various flags in shared items for each field and the
        // test checks the flags in cache are set correctly (they are determined in cache writer).
        Assert.DoesNotThrow(() => TestHelper.LoadSaveAndCompare(
            @"Other\PivotTableReferenceFiles\VariousDataTypesInTableColumns\input.xlsx",
            @"Other\PivotTableReferenceFiles\VariousDataTypesInTableColumns\output.xlsx"));
    }

    [Test]
    public void BlankPivotTableField()
    {
        using var ms = new MemoryStream();
        TestHelper.CreateAndCompare(() =>
        {
            // Based on .\XLibur\XLibur.Examples\PivotTables\PivotTables.cs
            // But with empty column for Month
            var pastries = new List<Pastry>
            {
                new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
                new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
                new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
                new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
                new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
                new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
                new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
                new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
                new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
                new Pastry("Danish", 104, 394, -20.24, "", null),
                new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
                new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

                // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
                new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
                new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

                // For ContainsBlank and integer rows/columns test
                new Pastry("Scone", null, 255, 18.4, "", null),
            };

            var wb = new XLWorkbook();

            var sheet = wb.Worksheets.Add("PastrySalesData");
            // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
            var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
            sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Columns().AdjustToContents();

            for (var i = 1; i <= 5; i++)
            {
                // Add a new sheet for our pivot table
                var ptSheet = wb.Worksheets.Add("pvt" + i);

                // Create the pivot table, using the data from the "PastrySalesData" table
                var pt = ptSheet.PivotTables.Add("pvt" + i, ptSheet.Cell(1, 1), table);

                if (i == 1 || i == 4 || i == 5)
                    pt.ColumnLabels.Add("Name");
                else if (i == 2 || i == 3)
                    pt.RowLabels.Add("Name");

                if (i == 1 || i == 3)
                    pt.RowLabels.Add("Month");
                else if (i == 2 || i == 4)
                    pt.ColumnLabels.Add("Month");
                else if (i == 5)
                    pt.RowLabels.Add("BakeDate");

                // The values in our table will come from the "NumberOfOrders" field
                // The default calculation setting is a total of each row/column
                pt.Values.Add("NumberOfOrders", "NumberOfOrdersPercentageOfBearclaw")
                    .ShowAsPercentageFrom("Name").And("Bearclaw")
                    .NumberFormat.Format = "0%";

                ptSheet.Columns().AdjustToContents();
            }

            return wb;
        }, @"Other\PivotTableReferenceFiles\BlankPivotTableField\BlankPivotTableField.xlsx");
    }

    [Test]
    public void SourceSheetWithWhitespace()
    {
        // Check that pivot source reference for a sheet name with whitespaces
        // is not saved to the file with escaped quotes, issue #955.
        TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook();

            // Worksheet name contains whitespaces that shouldn't be quoted in the file.
            var sheet = wb.Worksheets.Add("Pastry Sales Data");
            var range = sheet.Cell(1, 1).InsertData(new object[]
            {
                ("Name", "Sold count"),
                ("Pie", 7),
                ("Cake", 10),
                ("Pie", 2),
            });

            // Add a new sheet for our pivot table
            var ptSheet = wb.Worksheets.Add("pvt");
            var pt = ptSheet.PivotTables.Add("pvt", ptSheet.Cell(1, 1), range);
            pt.RowLabels.Add("Name");
            pt.Values.Add("Sold count");

            return wb;
        }, @"Other\PivotTableReferenceFiles\SourceSheetWithWhitespace\outputfile.xlsx");
    }

    [Test]
    public void PivotTableWithNoneTheme()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\PivotTableWithNoneTheme\inputfile.xlsx"));
        using var ms = new MemoryStream();
        TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook(stream);
            wb.SaveAs(ms);
            return wb;
        }, @"Other\PivotTableReferenceFiles\PivotTableWithNoneTheme\outputfile.xlsx");
    }

    [Test]
    public void MaintainPivotTableLabelsOrder()
    {
        var pastries = new List<Pastry>
        {
            new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
            new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
            new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
            new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
            new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
            new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
            new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
            new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
            new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
            new Pastry("Danish", 104, 394, -20.24, "", null),
            new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
            new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

            // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
            new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
            new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
            new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

            // For ContainsBlank and integer rows/columns test
            new Pastry("Scone", null, 255, 18.4, "", null),
        };

        using (var ms = new MemoryStream())
        {
            // Page fields
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("PastrySalesData");
                // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Columns().AdjustToContents();

                IXLWorksheet ptSheet;
                IXLPivotTable pt;

                // Add a new sheet for our pivot table
                ptSheet = wb.Worksheets.Add("pvt");

                // Create the pivot table, using the data from the "PastrySalesData" table
                pt = ptSheet.PivotTables.Add("PastryPivot", ptSheet.Cell(1, 1), table);

                pt.ReportFilters.Add("Month");
                pt.ReportFilters.Add("Name");

                pt.RowLabels.Add("BakeDate");
                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var pageFields = wb.Worksheets.SelectMany(ws => ws.PivotTables)
                    .First()
                    .ReportFilters
                    .ToArray();

                Assert.AreEqual("Month", pageFields[0].SourceName);
                Assert.AreEqual("Name", pageFields[1].SourceName);
            }
        }

        using (var ms = new MemoryStream())
        {
            // Column labels
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("PastrySalesData");
                // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Columns().AdjustToContents();

                IXLWorksheet ptSheet;
                IXLPivotTable pt;

                // Add a new sheet for our pivot table
                ptSheet = wb.Worksheets.Add("pvt");

                // Create the pivot table, using the data from the "PastrySalesData" table
                pt = ptSheet.PivotTables.Add("PastryPivot", ptSheet.Cell(1, 1), table);

                pt.ColumnLabels.Add("Month");
                pt.ColumnLabels.Add("Name");

                pt.RowLabels.Add("BakeDate");
                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var columnLabels = wb.Worksheets.SelectMany(ws => ws.PivotTables)
                    .First()
                    .ColumnLabels
                    .ToArray();

                Assert.AreEqual("Month", columnLabels[0].SourceName);
                Assert.AreEqual("Name", columnLabels[1].SourceName);
            }
        }

        using (var ms = new MemoryStream())
        {
            // Row labels
            using (var wb = new XLWorkbook())
            {
                var sheet = wb.Worksheets.Add("PastrySalesData");
                // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Columns().AdjustToContents();

                IXLWorksheet ptSheet;
                IXLPivotTable pt;

                // Add a new sheet for our pivot table
                ptSheet = wb.Worksheets.Add("pvt");

                // Create the pivot table, using the data from the "PastrySalesData" table
                pt = ptSheet.PivotTables.Add("PastryPivot", ptSheet.Cell(1, 1), table);

                pt.RowLabels.Add("Month");
                pt.RowLabels.Add("Name");
                pt.RowLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);

                pt.ColumnLabels.Add("BakeDate");
                pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var rowLabels = wb.Worksheets.SelectMany(ws => ws.PivotTables)
                    .First()
                    .RowLabels
                    .ToArray();

                Assert.AreEqual("Month", rowLabels[0].SourceName);
                Assert.AreEqual("Name", rowLabels[1].SourceName);
                Assert.AreEqual("{{Values}}", rowLabels[2].SourceName);
            }
        }
    }

    [Test]
    public void MaintainPivotTableIntegrityOnMultipleSaves()
    {
        var pastries = new List<Pastry>
        {
            new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
            new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
            new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
            new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
            new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
            new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
            new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
            new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
            new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
            new Pastry("Danish", 104, 394, -20.24, "", null),
            new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
            new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

            // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
            new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
            new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
            new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

            // For ContainsBlank and integer rows/columns test
            new Pastry("Scone", null, 255, 18.4, "", null),
        };

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("PastrySalesData");
            var table = ws.FirstCell().InsertTable(pastries, "PastrySalesData", true);

            var pvtSheet = wb.Worksheets.Add("pvt");
            var pvt = table.CreatePivotTable(pvtSheet.FirstCell(), "PastryPvt");

            pvt.ColumnLabels.Add("Month");
            pvt.RowLabels.Add("Name");
            pvt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

            //Deliberately try to save twice
            wb.SaveAs(ms);
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            Assert.AreEqual(1, wb.Worksheets.SelectMany(ws => ws.PivotTables).Count());
        }
    }

    [Test]
    public void TwoPivotWithOneSourceTest()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\TwoPivotTablesWithSingleSource\input.xlsx"));
        TestHelper.CreateAndCompare(() =>
        {
            var wb = new XLWorkbook(stream);
            var srcRange = wb.Range("Sheet1!$B$2:$H$207");

            var pivotSource = wb.PivotCaches.Add(srcRange);

            foreach (var pt in wb.Worksheets.SelectMany(ws => ws.PivotTables))
            {
                pt.PivotCache = pivotSource;
            }

            return wb;
        }, @"Other\PivotTableReferenceFiles\TwoPivotTablesWithSingleSource\output.xlsx");
    }

    [Test]
    public void PivotSubtotalsLoadingTest()
    {
        // Make sure that if the original file has *subtotals*, the subtotals are
        // turned on even after loading into XLibur and then saving the document.
        Assert.DoesNotThrow(() => TestHelper.LoadSaveAndCompare(
            @"Other\PivotTableReferenceFiles\PivotSubtotalsSource\input.xlsx",
            @"Other\PivotTableReferenceFiles\PivotSubtotalsSource\output.xlsx"));
    }

    [Test]
    public void ClearPivotTableRenderedRange()
    {
        // https://github.com/XLibur/XLibur/pull/856
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\ClearPivotTableRenderedRangeWhenLoading\inputfile.xlsx"));
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheet("Sheet1");
            Assert.IsTrue(ws.Cell("B1").IsEmpty());
            Assert.IsTrue(ws.Cell("C2").IsEmpty());
            Assert.IsTrue(ws.Cell("D5").IsEmpty());
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet("Sheet1");
            Assert.IsTrue(ws.Cell("B1").IsEmpty());
            Assert.IsTrue(ws.Cell("C2").IsEmpty());
            Assert.IsTrue(ws.Cell("D5").IsEmpty());
        }
    }

    [Test]
    public void Add_all_pivot_tables_for_same_range_use_same_pivot_cache()
    {
        // Two different pivot tables created from same range use same pivot cache
        // and don't create a separate pivot cache for each pivot table.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.FirstCell().InsertData(new object[]
        {
            ("Name", "Count"),
            ("Pie", 14),
        });

        var rangePivot1 = ws.PivotTables.Add("rangePivot1", ws.Cell("D1"), range);
        var rangePivot2 = ws.PivotTables.Add("rangePivot2", ws.Cell("D20"), range);

        Assert.AreNotSame(rangePivot1, rangePivot2);
        Assert.AreSame(rangePivot1.PivotCache, rangePivot2.PivotCache);
    }

    [Test]
    public void Add_all_pivot_tables_for_same_table_use_same_pivot_cache()
    {
        // Two different pivot tables created from same table use same pivot cache
        // and don't create a separate pivot cache for each pivot table.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var table = ws.FirstCell().InsertTable(new object[]
        {
            ("Name", "Count"),
            ("Pie", 14),
        });

        var tablePivot1 = ws.PivotTables.Add("tablePivot1", ws.Cell("J1"), table);
        var tablePivot2 = ws.PivotTables.Add("tablePivot2", ws.Cell("J20"), table);

        Assert.AreNotSame(tablePivot1, tablePivot2);
        Assert.AreSame(tablePivot1.PivotCache, tablePivot2.PivotCache);
    }

    [Test]
    public void Add_pivot_tables_will_use_table_as_source_if_range_matches_table_area()
    {
        // When a pivot table is created, the `Add` method tries to first
        // find a table with same area as the requested range. If it finds one,
        // the cache will be created from the table and not a range. That is the
        // Excel behavior and generally makes sense.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().InsertTable(new object[]
        {
            ("Name", "Count"),
            ("Pie", 14),
        }, "Test_table");

        // A range that matches the size of an area
        var matchingRange = ws.Range("A1:B3");

        var tablePivot1 = ws.PivotTables.Add("tablePivot1", ws.Cell("J1"), matchingRange);

        var cacheSource = (XLPivotSourceReference)((XLPivotCache)tablePivot1.PivotCache).Source;
        Assert.True(cacheSource.UsesName);
        Assert.AreEqual("Test_table", cacheSource.Name);
    }

    [Test]
    public void Load_and_save_pivot_table_with_cache_records_but_missing_source_data()
    {
        // Test file contains a pivot table created from a normal table in
        // a sheet that was already deleted. The file contains cache records,
        // but the table and original sheet are gone. It's possible to load
        // and save such a pivot table.
        // Opening the saved file in Excel throws an error 'Reference isn't valid'
        // on load, because of `RefreshOnLoad` flag. That flag is enabled by default for
        // newly created pivot caches (XLibur relies on Excel to rebuild the table and fix it).
        // At this time, there is no content, only shape, because we don't have an engine
        // to determine correct layout and values. Change RefreshDataOnOpen to 0 and change
        // PT in Excel to see the values (aka gimp on Excel PT engine).
        Assert.DoesNotThrow(() => TestHelper.LoadSaveAndCompare(
            @"Other\PivotTableReferenceFiles\PivotTableWithoutSourceData-input.xlsx",
            @"Other\PivotTableReferenceFiles\PivotTableWithoutSourceData-output.xlsx"));
    }

    [Test]
    [Description("https://github.com/ClosedXML/ClosedXML/issues/2219")]
    public void Pivot_field_item_hidden_flags_survive_round_trip()
    {
        // Pivot table has filters applied through hidden items (h="1") on row fields.
        // Previously, the cache writer always set refreshOnLoad=true, causing Excel
        // to rebuild the pivot table on open and lose all applied filters.
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\Pivotfilters_lost_2219.xlsx"));
        using var wb = new XLWorkbook(stream);

        var pt = (XLPivotTable)wb.Worksheets.First().PivotTables.First();

        // Field 2 (Modell) has items with h="1" (hidden) — these represent the applied filter.
        var modellField = pt.PivotFields[2];
        var hiddenItems = modellField.Items.Where(i => i.Hidden).ToList();
        Assert.That(hiddenItems.Count, Is.GreaterThan(0), "Precondition: field should have hidden items");

        // Save and reload
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        ms.Position = 0;
        using var wb2 = new XLWorkbook(ms);
        var pt2 = (XLPivotTable)wb2.Worksheets.First().PivotTables.First();
        var modellField2 = pt2.PivotFields[2];

        // Hidden items must be preserved
        var hiddenItems2 = modellField2.Items.Where(i => i.Hidden).ToList();
        Assert.That(hiddenItems2.Count, Is.EqualTo(hiddenItems.Count),
            "Hidden items (filters) should survive round-trip");

        // RefreshOnLoad should not be forced to true
        Assert.That(pt2.PivotCache.RefreshDataOnOpen, Is.False,
            "RefreshDataOnOpen should preserve original value (false)");
    }

    [Test]
    public void Skips_chartsheets_during_pivot_table_loading()
    {
        // Pivot table loading code looks for pivot tables on each sheet, but it shouldn't
        // crash when sheet is a chartsheet or other type of sheet. The referenced test file
        // contains chartsheet and a pivot table to ensure that loading code won't crash.
        TestHelper.LoadAndAssert(wb =>
        {
            // Check that existing pivot table is loaded.
            Assert.True(wb.Worksheet("pivot").PivotTables.Contains("Pastries"));
        }, @"Other\PivotTableReferenceFiles\ChartsheetAndPivotTable.xlsx");
    }

    #region IXLPivotTable properties

    #region TargetCell

    [Test]
    public void Property_TargetCell_sets_value_of_the_top_left_corner_of_pivot_table()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var data = ws.Cell("A1").InsertData(new object[]
        {
            ("Name", "City", "Flavor", "Sales"),
            ("Cake", "Tokyo", "Vanilla", 7),
        });
        var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), data);
        pt.ReportFilters.Add("City");

        // Even when we added filter and a gap row, the target cell is still E1
        Assert.AreEqual("E1", pt.TargetCell.Address.ToString());
        Assert.AreEqual("E3", ((XLPivotTable)pt).Area.FirstPoint.ToString());

        pt.TargetCell = ws.Cell("E2");
        Assert.AreEqual("E2", pt.TargetCell.Address.ToString());
        Assert.AreEqual("E4", ((XLPivotTable)pt).Area.FirstPoint.ToString());
    }

    #endregion

    #region FilterAreaOrder

    [TestCase(XLFilterAreaOrder.DownThenOver, "E5")]
    [TestCase(XLFilterAreaOrder.OverThenDown, "E3")]
    public void Property_FilterAreaOrder_determines_direction_in_which_are_filter_fields_laid_out(XLFilterAreaOrder order, string tableAddress)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var data = ws.Cell("A1").InsertData(new object[]
        {
            ("Name", "City", "Flavor", "Sales"),
            ("Cake", "Tokyo", "Vanilla", 7),
        });

        var pt = ws.PivotTables.Add("pt", ws.Cell("E1"), data);
        pt.FilterAreaOrder = order;

        pt.ReportFilters.Add("Name");
        pt.ReportFilters.Add("City");
        pt.ReportFilters.Add("Flavor");

        // Indirect detection of filter fields layout: The address of pivot table are is
        // determined by filter area order.
        Assert.AreEqual(tableAddress, ((XLPivotTable)pt).Area.ToString());
    }

    #endregion

    #region Layout

    [TestCase(XLPivotLayout.Outline, "Property_layout_sets_layout_of_pivot_table_and_all_fields-outline.xlsx")]
    [TestCase(XLPivotLayout.Tabular, "Property_layout_sets_layout_of_pivot_table_and_all_fields-tabular.xlsx")]
    [TestCase(XLPivotLayout.Compact, "Property_layout_sets_layout_of_pivot_table_and_all_fields-compact.xlsx")]
    public void Property_layout_sets_layout_of_pivot_table_and_all_fields(XLPivotLayout layout, string testFile)
    {
        // The pivot table also contains unused field Currency. It is there, because tabular
        // layout doesn't display header fields properly (i.e. one header per axis field),
        // unless all (even fields that are not on any axis) have the same field layout.
        TestHelper.CreateAndCompare(wb =>
            {
                var dataSheet = wb.AddWorksheet();
                var dataRange = dataSheet.Cell("A1").InsertData(new object[]
                {
                    ("Name", "Size", "Month", "Season", "Price", "Currency"),
                    ("Cake", "Small", "Jan", "Winter", 9, "EUR"),
                    ("Pie", "Small", "Jan", "Winter", 7, "EUR"),
                    ("Cake", "Large", "Feb", "Summer", 3, "CZK"),
                });

                var ptSheet = wb.AddWorksheet().SetTabActive();
                ptSheet.Column("A").Width = 15;
                var pt = dataRange.CreatePivotTable(ptSheet.Cell("A1"), "pivot table");

                // Add at least two fields to each axis to make each layout distinctive.
                pt.RowLabels.Add("Name");
                pt.RowLabels.Add("Size");
                pt.ColumnLabels.Add("Month");
                pt.ColumnLabels.Add("Season");
                pt.Values.Add("Price");

                pt.Layout = layout;
            }, $@"Other\PivotTable\TableProps\{testFile}");
    }

    #endregion

    #endregion

    private static void SetFieldOptions(IXLPivotField field, bool withDefaults)
    {
        field.SubtotalsAtTop = !withDefaults;
        field.ShowBlankItems = !withDefaults;
        field.Outline = !withDefaults;
        field.Compact = !withDefaults;
        field.Collapsed = withDefaults;
        field.InsertBlankLines = withDefaults;
        field.RepeatItemLabels = withDefaults;
        field.InsertPageBreaks = withDefaults;
        field.IncludeNewItemsInFilter = withDefaults;
    }

    private static void AssertFieldOptions(IXLPivotField field, bool withDefaults)
    {
        Assert.AreEqual(!withDefaults, field.SubtotalsAtTop, "SubtotalsAtTop save failure");
        Assert.AreEqual(!withDefaults, field.ShowBlankItems, "ShowBlankItems save failure");
        Assert.AreEqual(!withDefaults, field.Outline, "Outline save failure");
        Assert.AreEqual(!withDefaults, field.Compact, "Compact save failure");
        Assert.AreEqual(withDefaults, field.Collapsed, "Collapsed save failure");
        Assert.AreEqual(withDefaults, field.InsertBlankLines, "InsertBlankLines save failure");
        Assert.AreEqual(withDefaults, field.RepeatItemLabels, "RepeatItemLabels save failure");
        Assert.AreEqual(withDefaults, field.InsertPageBreaks, "InsertPageBreaks save failure");
        Assert.AreEqual(withDefaults, field.IncludeNewItemsInFilter, "IncludeNewItemsInFilter save failure");
    }

    [Test]
    [Description("Loading pivots with custom theme should not throw (ClosedXML#1429)")]
    public void PivotTableWithCustomTheme_CanLoadAndSave()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Lion\PivotTables\CustomPivotTheme.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        Assert.That(ms.Length, Is.GreaterThan(0));
    }

    [Test]
    [Description("Loading pivots with styles should not throw (ClosedXML#1429)")]
    public void PivotTableWithStyles_CanLoadAndSave()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Lion\PivotTables\PivotWithStyles.xlsx"));
        using var wb = new XLWorkbook(stream);
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        Assert.That(ms.Length, Is.GreaterThan(0));
    }

    [Test]
    [Description("Loading an Excel-created workbook with a pivot table calculated field should not throw IncorrectElementsCount")]
    public void PivotTable_with_calculated_field_can_be_loaded()
    {
        // The file has 3 database fields (datum, weekdag, verbruik) and 1 calculated field (Field1 = verbruik*2).
        // Before the fix, ReadRecords compared record.ChildElements.Count (3) against FieldCount (4),
        // causing IncorrectElementsCount. The fix uses DatabaseFieldCount (3) instead.
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTable\pivottable_customfield.xlsx"));
        using var wb = new XLWorkbook(stream);

        // Verify calculated field formula survived load
        var pivotCache = wb.PivotCachesInternal.Cast<XLPivotCache>().First();
        Assert.That(pivotCache.FieldCount, Is.EqualTo(4));
        Assert.That(pivotCache.DatabaseFieldCount, Is.EqualTo(3));
        Assert.That(pivotCache.GetCalculatedFieldFormula(3), Is.EqualTo("verbruik*2"));

        // Round-trip: save and reload
        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        using var doc = SpreadsheetDocument.Open(ms, false);
        var cachePart = doc.WorkbookPart!.GetPartsOfType<PivotTableCacheDefinitionPart>().First();
        var cacheFields = cachePart.PivotCacheDefinition.CacheFields!.Elements<CacheField>().ToList();
        Assert.That(cacheFields.Count, Is.EqualTo(4));

        var calcField = cacheFields[3];
        Assert.That(calcField.Name?.Value, Is.EqualTo("Field1"));
        Assert.That(calcField.Formula?.Value, Is.EqualTo("verbruik*2"));
        Assert.That(calcField.DatabaseField?.Value, Is.False);
    }

    [Test]
    [Description("Pivot table with calculated field should round-trip without losing the formula (ClosedXML#885)")]
    public void PivotTableWithCalculatedField_RoundTrips()
    {
        // Create an xlsx with a pivot table that has a calculated field using raw OpenXML SDK.
        using var inputStream = new MemoryStream();
        CreateWorkbookWithCalculatedField(inputStream);
        inputStream.Position = 0;

        // Load through XLibur and re-save.
        using var wb = new XLWorkbook(inputStream);
        using var outputStream = new MemoryStream();
        wb.SaveAs(outputStream);

        // Verify the calculated field survived the round-trip.
        outputStream.Position = 0;
        using var doc = SpreadsheetDocument.Open(outputStream, false);
        var cachePart = doc.WorkbookPart!.GetPartsOfType<PivotTableCacheDefinitionPart>().First();
        var cacheFields = cachePart.PivotCacheDefinition.CacheFields!.Elements<CacheField>().ToList();

        // Should have 3 data fields + 1 calculated field = 4 total
        Assert.That(cacheFields.Count, Is.EqualTo(4), "Expected 3 source fields + 1 calculated field");

        var calculatedField = cacheFields.Last();
        Assert.That(calculatedField.Name?.Value, Is.EqualTo("Profit"), "Calculated field name should be 'Profit'");
        Assert.That(calculatedField.Formula?.Value, Is.EqualTo("Revenue - Cost"), "Calculated field formula should be preserved");
        Assert.That(calculatedField.DatabaseField?.Value, Is.False, "Calculated field DatabaseField should be false");
    }

    /// <summary>
    /// Creates a minimal xlsx with a pivot table that includes a calculated field "Profit = Revenue - Cost".
    /// </summary>
    private static void CreateWorkbookWithCalculatedField(Stream stream)
    {
        using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        // Workbook part
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        // Data worksheet
        var dataSheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var dataSheetPartId = workbookPart.GetIdOfPart(dataSheetPart);
        sheets.Append(new Sheet { Id = dataSheetPartId, SheetId = 1, Name = "Data" });

        var sheetData = new SheetData();
        // Header row
        sheetData.Append(CreateRow(1, "Name", "Revenue", "Cost"));
        // Data rows
        sheetData.Append(CreateNumericRow(2, "Cookies", 500, 200));
        sheetData.Append(CreateNumericRow(3, "Cake", 800, 350));
        sheetData.Append(CreateNumericRow(4, "Pie", 300, 100));
        dataSheetPart.Worksheet = new Worksheet(sheetData);

        // Pivot table worksheet
        var pivotSheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var pivotSheetPartId = workbookPart.GetIdOfPart(pivotSheetPart);
        sheets.Append(new Sheet { Id = pivotSheetPartId, SheetId = 2, Name = "PivotSheet" });
        pivotSheetPart.Worksheet = new Worksheet(new SheetData());

        // Pivot cache definition
        var cachePart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        var cachePartId = workbookPart.GetIdOfPart(cachePart);

        var cacheDefinition = new PivotCacheDefinition
        {
            Id = "rId1",
            RefreshOnLoad = true,
            CreatedVersion = 5,
            RefreshedVersion = 5,
        };
        cacheDefinition.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var cacheSource = new CacheSource { Type = SourceValues.Worksheet };
        cacheSource.Append(new WorksheetSource { Sheet = "Data", Reference = "A1:C4" });
        cacheDefinition.Append(cacheSource);

        var cacheFields = new CacheFields();
        // Field 0: Name (string)
        var nameField = new CacheField { Name = "Name" };
        var nameShared = new SharedItems { ContainsSemiMixedTypes = false, ContainsString = true, ContainsNumber = false, Count = 3 };
        nameShared.Append(new StringItem { Val = "Cookies" });
        nameShared.Append(new StringItem { Val = "Cake" });
        nameShared.Append(new StringItem { Val = "Pie" });
        nameField.SharedItems = nameShared;
        cacheFields.Append(nameField);

        // Field 1: Revenue (number)
        var revenueField = new CacheField { Name = "Revenue" };
        var revenueShared = new SharedItems { ContainsSemiMixedTypes = false, ContainsString = false, ContainsNumber = true, MinValue = 300, MaxValue = 800, Count = 3 };
        revenueShared.Append(new NumberItem { Val = 500 });
        revenueShared.Append(new NumberItem { Val = 800 });
        revenueShared.Append(new NumberItem { Val = 300 });
        revenueField.SharedItems = revenueShared;
        cacheFields.Append(revenueField);

        // Field 2: Cost (number)
        var costField = new CacheField { Name = "Cost" };
        var costShared = new SharedItems { ContainsSemiMixedTypes = false, ContainsString = false, ContainsNumber = true, MinValue = 100, MaxValue = 350, Count = 3 };
        costShared.Append(new NumberItem { Val = 200 });
        costShared.Append(new NumberItem { Val = 350 });
        costShared.Append(new NumberItem { Val = 100 });
        costField.SharedItems = costShared;
        cacheFields.Append(costField);

        // Field 3: Profit (calculated field) - Formula = "Revenue - Cost"
        var profitField = new CacheField { Name = "Profit", Formula = "Revenue - Cost", DatabaseField = false };
        cacheFields.Append(profitField);

        cacheFields.Count = 4;
        cacheDefinition.Append(cacheFields);

        // Cache records
        var recordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
        var records = new PivotCacheRecords { Count = 3 };
        // Records for 3 source fields only (calculated fields don't have records)
        records.Append(new PivotCacheRecord(new FieldItem { Val = 0 }, new NumberItem { Val = 500 }, new NumberItem { Val = 200 }));
        records.Append(new PivotCacheRecord(new FieldItem { Val = 1 }, new NumberItem { Val = 800 }, new NumberItem { Val = 350 }));
        records.Append(new PivotCacheRecord(new FieldItem { Val = 2 }, new NumberItem { Val = 300 }, new NumberItem { Val = 100 }));
        recordsPart.PivotCacheRecords = records;

        cachePart.PivotCacheDefinition = cacheDefinition;

        // Register the cache in the workbook
        var pivotCaches = new PivotCaches();
        pivotCaches.Append(new PivotCache { CacheId = 0, Id = cachePartId });
        workbookPart.Workbook.Append(pivotCaches);

        // Pivot table part
        var pivotTablePart = pivotSheetPart.AddNewPart<PivotTablePart>();

        // Link the pivot table part to the cache definition part
        var pivotTableCacheRelId = pivotTablePart.CreateRelationshipToPart(cachePart);

        var pivotTableDef = new PivotTableDefinition
        {
            Name = "PivotTable1",
            CacheId = 0,
            DataCaption = "Values",
            CreatedVersion = 5,
            UpdatedVersion = 5,
        };

        var location = new Location { Reference = "A1:B5", FirstHeaderRow = 1, FirstDataRow = 2, FirstDataColumn = 1 };
        pivotTableDef.Append(location);

        // Pivot fields: one per cache field
        var pivotFields = new PivotFields { Count = 4 };

        // Field 0: Name - axis row
        var pf0 = new PivotField { Axis = PivotTableAxisValues.AxisRow, ShowAll = false };
        var items0 = new Items { Count = 4 };
        items0.Append(new Item { Index = 0 });
        items0.Append(new Item { Index = 1 });
        items0.Append(new Item { Index = 2 });
        items0.Append(new Item { ItemType = ItemValues.Default });
        pf0.Append(items0);
        pivotFields.Append(pf0);

        // Field 1: Revenue - not on any axis
        pivotFields.Append(new PivotField { ShowAll = false });

        // Field 2: Cost - not on any axis
        pivotFields.Append(new PivotField { ShowAll = false });

        // Field 3: Profit - calculated field, used as data field
        pivotFields.Append(new PivotField
        {
            DataField = true,
            ShowAll = false,
            DefaultSubtotal = false,
            DragToRow = false,
            DragToColumn = false,
            DragToPage = false,
        });

        pivotTableDef.Append(pivotFields);

        // Row fields
        var rowFields = new RowFields { Count = 1 };
        rowFields.Append(new Field { Index = 0 });
        pivotTableDef.Append(rowFields);

        // Row items
        var rowItems = new RowItems { Count = 4 };
        rowItems.Append(new RowItem(new MemberPropertyIndex { Val = 0 }));
        rowItems.Append(new RowItem(new MemberPropertyIndex { Val = 1 }));
        rowItems.Append(new RowItem(new MemberPropertyIndex { Val = 2 }));
        rowItems.Append(new RowItem(new MemberPropertyIndex()) { ItemType = ItemValues.Grand });
        pivotTableDef.Append(rowItems);

        // Column items
        var colItems = new ColumnItems { Count = 1 };
        colItems.Append(new RowItem(new MemberPropertyIndex()) { ItemType = ItemValues.Grand });
        pivotTableDef.Append(colItems);

        // Data fields (the calculated field "Profit")
        var dataFields = new DataFields { Count = 1 };
        dataFields.Append(new DataField { Name = "Sum of Profit", Field = 3 });
        pivotTableDef.Append(dataFields);

        // Pivot table style
        pivotTableDef.Append(new PivotTableStyle { Name = "PivotStyleLight16", ShowRowHeaders = true, ShowColumnHeaders = true });

        pivotTablePart.PivotTableDefinition = pivotTableDef;
    }

    private static Row CreateRow(uint rowIndex, params string[] values)
    {
        var row = new Row { RowIndex = rowIndex };
        for (var i = 0; i < values.Length; i++)
        {
            row.Append(new Cell
            {
                CellReference = $"{(char)('A' + i)}{rowIndex}",
                DataType = CellValues.String,
                CellValue = new CellValue(values[i])
            });
        }
        return row;
    }

    private static Row CreateNumericRow(uint rowIndex, string name, double val1, double val2)
    {
        var row = new Row { RowIndex = rowIndex };
        row.Append(new Cell { CellReference = $"A{rowIndex}", DataType = CellValues.String, CellValue = new CellValue(name) });
        row.Append(new Cell { CellReference = $"B{rowIndex}", DataType = CellValues.Number, CellValue = new CellValue(val1) });
        row.Append(new Cell { CellReference = $"C{rowIndex}", DataType = CellValues.Number, CellValue = new CellValue(val2) });
        return row;
    }

    [Test]
    public void Deleting_sheet_with_pivot_table_does_not_throw_on_save()
    {
        // https://github.com/ClosedXML/ClosedXML/issues/2737
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx"));
        using var wb = new XLWorkbook(stream);

        // The file has a sheet "PivotTable1" that contains a pivot table
        wb.Worksheet("PivotTable1").Delete();

        using var ms = new MemoryStream();
        Assert.DoesNotThrow(() => wb.SaveAs(ms));
    }
}
