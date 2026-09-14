using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// A sheet is deleted or renamed through one door, the worksheet collection, and every holder of text
/// that names the sheet hears about it (spec 55).
/// </summary>
public class SheetLifecycleTests
{
    private const string BlockedOnTask3 =
        "Spec 55 task 3 (names at every scope) turns this green. It is blocked on Excel-authored "
        + "fixtures the owner has not made yet: delete-before.xlsx/delete-after.xlsx and "
        + "refdelete-before.xlsx/refdelete-after.xlsx in Resource/Other/SheetLifecycle.";

    private const string GreenInTask2 = "Spec 55 task 2 (the door) turns this green.";

    /// <summary>
    /// D53: <c>wb.Worksheets.Delete</c> skipped <c>IsDeleted</c>, the name fix-up and the calc-engine
    /// purge, so a dependent kept the deleted sheet's value and saved it as its cached value.
    /// </summary>
    [Test]
    [Skip(GreenInTask2)]
    public async Task D53_the_collection_delete_does_everything_the_sheet_delete_does()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var sheet1 = wb.AddWorksheet("Sheet1");
            var sheet2 = wb.AddWorksheet("Sheet2");
            sheet1.Cell("A1").Value = 5;
            sheet2.Cell("A1").FormulaA1 = "Sheet1!A1*2";
            wb.DefinedNames.Add("W", "Sheet1!$A$1");
            await Assert.That(sheet2.Cell("A1").Value).IsEqualTo(10);

            wb.Worksheets.Delete("Sheet1");

            await Assert.That(((XLWorksheet)sheet1).IsDeleted).IsTrue();
            await Assert.That(sheet2.Cell("A1").Value).IsEqualTo(XLError.CellReference);
            await Assert.That(wb.DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet("Sheet2").Cell("A1").Value).IsEqualTo(XLError.CellReference);
    }

    /// <summary>
    /// D54: a sheet-scoped name kept pointing at a deleted sheet, through a save and a reload, while
    /// the workbook-scoped control became <c>#REF!</c>.
    /// </summary>
    [Test]
    [Skip(BlockedOnTask3)]
    public async Task D54_a_sheet_scoped_name_on_another_sheet_loses_a_deleted_sheet()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1");
            var sheet2 = wb.AddWorksheet("Sheet2");
            sheet2.DefinedNames.Add("N", "Sheet1!$A$1");
            wb.DefinedNames.Add("W", "Sheet1!$A$1");

            wb.Worksheet("Sheet1").Delete();

            await Assert.That(sheet2.DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
            await Assert.That(wb.DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet("Sheet2").DefinedNames.Single().RefersTo).IsEqualTo("#REF!");
    }

    /// <summary>
    /// D55: a rename skipped a sheet-qualified name and a 3D reference in a defined name, while the
    /// control <c>P</c> was renamed.
    /// </summary>
    [Test]
    [Skip(BlockedOnTask3)]
    public async Task D55_a_rename_reaches_sheet_qualified_names_and_3D_references_in_defined_names()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Sheet2");
        wb.AddWorksheet("Sheet3");
        sheet1.DefinedNames.Add("Local", "Sheet1!$B$1");
        wb.DefinedNames.Add("Q", "Sheet1!Local");
        wb.DefinedNames.Add("ThreeD", "SUM(Sheet1:Sheet3!$A$1)");
        wb.DefinedNames.Add("P", "Sheet1!$A$1");

        sheet1.Name = "Data";

        await Assert.That(RefersTo(wb, "Q")).IsEqualTo("Data!Local");
        await Assert.That(RefersTo(wb, "ThreeD")).IsEqualTo("SUM(Data:Sheet3!$A$1)");
        await Assert.That(RefersTo(wb, "P")).IsEqualTo("Data!$A$1");
    }

    private static string RefersTo(XLWorkbook wb, string name)
        => wb.DefinedNames.Single(n => n.Name == name).RefersTo;
}
