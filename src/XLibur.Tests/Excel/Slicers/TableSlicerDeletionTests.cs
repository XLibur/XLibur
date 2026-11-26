using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;
using XLibur.Excel.Tables;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Slicers;

/// <summary>
/// A table slicer goes when its table or its column goes, and follows its column when the column
/// moves (#971).
/// </summary>
/// <remarks>
/// In the fixture the Data sheet (sheet1.xml, drawing1.xml) holds <c>SalesTable</c> (id 1, columns
/// Region, Date, Amount with ids 1, 2, 3) and the table slicer <c>Slicer_Region</c>. The Pivot sheet
/// holds a pivot slicer, <c>Slicer_Region1</c>, which none of these edits may touch.
/// </remarks>
public class TableSlicerDeletionTests
{
    private const string Fixture = @"TryToLoad\SlicersOnPivotAndTable.xlsx";

    [Test]
    public async Task Deleting_a_table_and_adding_another_does_not_move_its_slicer()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var data = wb.Worksheet("Data");
            data.Tables.Remove("SalesTable");
            data.Cell("H1").Value = "X";
            data.Cell("I1").Value = "Y";
            data.Cell("H2").Value = 1;
            data.Cell("I2").Value = 2;
            data.Range("H1:I2").CreateTable("Other");
            wb.SaveAs(saved);
        }

        await AssertTableSlicerGone(saved);
    }

    [Test]
    public async Task Deleting_the_sliced_table_field_takes_its_slicer_with_it()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Data").Table("SalesTable").Field("Region").Delete();
            wb.SaveAs(saved);
        }

        await AssertTableSlicerGone(saved);
    }

    [Test]
    public async Task Deleting_the_sliced_worksheet_column_takes_its_slicer_with_it()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Data").Column("A").Delete();
            wb.SaveAs(saved);
        }

        await AssertTableSlicerGone(saved);
    }

    [Test]
    public async Task A_new_table_slicer_follows_its_column_when_a_column_to_its_left_is_deleted()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var data = (XLWorksheet)wb.Worksheet("Data");
            var table = (XLTable)data.Table("SalesTable");
            data.SlicersInternal.AddTableSlicer(table, "Amount");
            table.Field("Date").Delete();
            wb.SaveAs(saved);
        }

        // Amount keeps tableColumn id 3 after Date goes; its position drops from 3 to 2.
        await Assert.That(saved.ReadPart("xl/tables/table1.xml")).Contains("id=\"3\" name=\"Amount\"");
        await Assert.That(CacheXml(saved)).Contains("tableSlicerCache tableId=\"1\" column=\"3\"");
        await AssertSchemaValid(saved);
    }

    [Test]
    public async Task A_loaded_table_slicer_survives_a_column_to_its_left_being_deleted()
    {
        // Give the fixture a slicer on Amount and round-trip it, so the slicer under test is a
        // loaded one rather than one created this session.
        using var withAmountSlicer = new MemoryStream();
        using (var wb = Load())
        {
            var data = (XLWorksheet)wb.Worksheet("Data");
            data.SlicersInternal.AddTableSlicer((XLTable)data.Table("SalesTable"), "Amount");
            wb.SaveAs(withAmountSlicer);
        }

        using var saved = new MemoryStream();
        withAmountSlicer.Position = 0;
        using (var wb = new XLWorkbook(withAmountSlicer))
        {
            wb.Worksheet("Data").Column("B").Delete();
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var slicers = reloaded.Worksheet("Data").Slicers;
        await Assert.That(slicers.Select(s => s.SourceFieldName)).IsEquivalentTo(new[] { "Region", "Amount" });
        await Assert.That(slicers.All(s => s.Table?.Name == "SalesTable")).IsTrue();
        await AssertSchemaValid(saved);
    }

    [Test]
    public async Task A_new_table_is_not_given_the_id_of_a_slicer_cache_whose_table_is_missing()
    {
        // A file whose table slicer cache names a table id no table has: the slicer did not bind on
        // load, so nothing removes it, and the next free id would hand it to the new table.
        using var source = TestHelper.OpenResource(Fixture)
            .RewritePart("xl/slicerCaches/slicerCache2.xml", xml => xml.Replace("tableId=\"1\"", "tableId=\"2\""));

        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            var data = wb.Worksheet("Data");
            data.Cell("H1").Value = "X";
            data.Cell("H2").Value = 1;
            data.Range("H1:H2").CreateTable("Other");
            wb.SaveAs(saved);
        }

        await Assert.That(CacheXml(saved)).Contains("tableSlicerCache tableId=\"2\"");

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(((XLTable)reloaded.Worksheet("Data").Table("Other")).Id).IsEqualTo((uint?)3);
        await Assert.That(reloaded.Worksheet("Data").Slicers.Single().Table).IsNull();
    }

    #region Helpers

    private static XLWorkbook Load() => TestHelper.LoadWorkbook(Fixture);

    private static string CacheXml(MemoryStream package) =>
        string.Concat(package.PartsUnder("xl/slicerCaches/").Select(package.ReadPart));

    /// <summary>
    /// The table slicer and each of its six pieces are gone, and the pivot slicer is still whole.
    /// </summary>
    private static async Task AssertTableSlicerGone(MemoryStream saved)
    {
        await Assert.That(CacheXml(saved)).DoesNotContain("tableSlicerCache");
        await Assert.That(CacheXml(saved)).Contains("name=\"Slicer_Region1\"");
        await Assert.That(saved.PartsUnder("xl/slicers/").Length).IsEqualTo(1);

        var workbookXml = saved.WorkbookXml();
        await Assert.That(workbookXml).DoesNotContain("name=\"Slicer_Region\">");
        await Assert.That(workbookXml).Contains("name=\"Slicer_Region1\">");
        await Assert.That(workbookXml).DoesNotContain("46BE6895-7355-4a93-B00E-2C351335B9C9")
            .Because("The x15:slicerCaches registry held only the table slicer's cache.");

        await Assert.That(saved.Sheet1Xml()).DoesNotContain("slicerList");
        if (saved.TryReadPart("xl/drawings/drawing1.xml", out var drawingXml))
            await Assert.That(drawingXml).DoesNotContain("/slicer");

        saved.Position = 0;
        using (var reloaded = new XLWorkbook(saved))
        {
            await Assert.That(reloaded.Worksheet("Data").Slicers.Count).IsEqualTo(0);
            await Assert.That(reloaded.Worksheet("Pivot").Slicers.Single().Name).IsEqualTo("Region 1");
        }

        await AssertSchemaValid(saved);
    }

    private static async Task AssertSchemaValid(MemoryStream package)
    {
        package.Position = 0;
        using var doc = SpreadsheetDocument.Open(package, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2010)
            .Validate(doc)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToList();

        await Assert.That(string.Join(Environment.NewLine, errors)).IsEmpty();
    }

    #endregion
}
