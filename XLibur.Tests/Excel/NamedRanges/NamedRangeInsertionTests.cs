using System.IO;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.NamedRanges;

public class NamedRangeInsertionTests
{
    private static string RefersTo(XLWorkbook wb, string name)
    {
        return wb.DefinedNames.First(dn => dn.Name == name).RefersTo;
    }

    /// <summary>
    /// When rows are inserted inside a named range by shifting cells down,
    /// the named range should expand to include the new rows,
    /// matching Excel's behavior.
    /// </summary>
    [Test]
    public async Task InsertingRowsInsideNamedRange_ExpandsRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // Fill A1:B8 with data
        for (var r = 1; r <= 8; r++)
        {
            ws.Cell(r, 1).Value = $"A{r}";
            ws.Cell(r, 2).Value = $"B{r}";
        }

        // Create named range for A1:B8
        ws.Range("A1:B8").AddToNamed("Region", XLScope.Workbook);

        // Insert 6 rows at row 3 (inside the named range), shifting cells down
        ws.Row(3).InsertRowsAbove(6);

        // Named range should expand from A1:B8 to A1:B14 (8 original + 6 inserted)
        var definedName = wb.DefinedNames.First(dn => dn.Name == "Region");
        await Assert.That(definedName.RefersTo).IsEqualTo("Sheet1!$A$1:$B$14").Because("Named range should expand when rows are inserted inside it");
    }

    /// <summary>
    /// When rows are inserted at the bottom boundary of a named range,
    /// the named range should expand.
    /// </summary>
    [Test]
    public async Task InsertingRowsAtBottomOfNamedRange_ExpandsRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var r = 1; r <= 8; r++)
        {
            ws.Cell(r, 1).Value = $"A{r}";
            ws.Cell(r, 2).Value = $"B{r}";
        }

        ws.Range("A1:B8").AddToNamed("Region", XLScope.Workbook);

        // Insert 3 rows at the last row of the range (row 8)
        ws.Row(8).InsertRowsAbove(3);

        var definedName = wb.DefinedNames.First(dn => dn.Name == "Region");
        await Assert.That(definedName.RefersTo).IsEqualTo("Sheet1!$A$1:$B$11").Because("Named range should expand when rows are inserted at its bottom boundary");
    }

    /// <summary>
    /// When rows are inserted above the named range (before its first row),
    /// the named range should shift down but NOT expand.
    /// </summary>
    [Test]
    public async Task InsertingRowsAboveNamedRange_ShiftsRangeDown()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var r = 1; r <= 8; r++)
        {
            ws.Cell(r, 1).Value = $"A{r}";
            ws.Cell(r, 2).Value = $"B{r}";
        }

        ws.Range("A1:B8").AddToNamed("Region", XLScope.Workbook);

        // Insert 2 rows above the named range
        ws.Row(1).InsertRowsAbove(2);

        // Named range should shift from A1:B8 to A3:B10 (same size, shifted down)
        var definedName = wb.DefinedNames.First(dn => dn.Name == "Region");
        await Assert.That(definedName.RefersTo).IsEqualTo("Sheet1!$A$3:$B$10").Because("Named range should shift down when rows are inserted above it");
    }

    /// <summary>
    /// When rows are inserted below the named range (after its last row),
    /// the named range should NOT change.
    /// </summary>
    [Test]
    public async Task InsertingRowsBelowNamedRange_DoesNotChangeRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var r = 1; r <= 8; r++)
        {
            ws.Cell(r, 1).Value = $"A{r}";
            ws.Cell(r, 2).Value = $"B{r}";
        }

        ws.Range("A1:B8").AddToNamed("Region", XLScope.Workbook);

        // Insert rows below the named range
        ws.Row(9).InsertRowsAbove(3);

        var definedName = wb.DefinedNames.First(dn => dn.Name == "Region");
        await Assert.That(definedName.RefersTo).IsEqualTo("Sheet1!$A$1:$B$8").Because("Named range should not change when rows are inserted below it");
    }

    /// <summary>
    /// Verifies the named range expansion survives a save/reload roundtrip.
    /// </summary>
    [Test]
    public async Task InsertingRowsInsideNamedRange_ExpandsRange_SurvivesRoundtrip()
    {
        using var ms = new MemoryStream();

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            for (var r = 1; r <= 8; r++)
            {
                ws.Cell(r, 1).Value = $"A{r}";
                ws.Cell(r, 2).Value = $"B{r}";
            }

            ws.Range("A1:B8").AddToNamed("Region", XLScope.Workbook);

            // Insert 6 rows at row 3, shifting cells down
            ws.Row(3).InsertRowsAbove(6);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb2 = new XLWorkbook(ms))
        {
            var definedName = wb2.DefinedNames.First(dn => dn.Name == "Region");
            await Assert.That(definedName.RefersTo).IsEqualTo("Sheet1!$A$1:$B$14").Because("Named range expansion should survive save/reload");

            // Verify the range actually resolves to correct cells
            var ranges = definedName.Ranges;
            await Assert.That(ranges.Count).IsEqualTo(1);
            var range = ranges.First();
            await Assert.That(range.RangeAddress.ToString()).IsEqualTo("$A$1:$B$14");
        }
    }

    /// <summary>
    /// A row-only reference expands the same way a cell range does: rows 3:5 with two rows inserted at
    /// row 4 becomes 3:7. The row-only branch used to shift both endpoints regardless of where the insert
    /// landed, giving 5:7 -- a reference that had moved off the rows it covered.
    /// </summary>
    [Test]
    public async Task InsertingRowsInsideRowOnlyNamedRange_ExpandsRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Rows", "Sheet1!$3:$5");

        ws.Row(4).InsertRowsAbove(2);

        await Assert.That(RefersTo(wb, "Rows")).IsEqualTo("Sheet1!$3:$7");
    }

    /// <summary>
    /// Inserting above a row-only reference shifts it down without expanding, as for a cell range.
    /// </summary>
    [Test]
    public async Task InsertingRowsAboveRowOnlyNamedRange_ShiftsRangeDown()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Rows", "Sheet1!$3:$5");

        ws.Row(1).InsertRowsAbove(2);

        await Assert.That(RefersTo(wb, "Rows")).IsEqualTo("Sheet1!$5:$7");
    }

    /// <summary>
    /// Column-only counterpart: columns C:E with two columns inserted before D becomes C:G.
    /// </summary>
    [Test]
    public async Task InsertingColumnsInsideColumnOnlyNamedRange_ExpandsRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Cols", "Sheet1!$C:$E");

        ws.Column(4).InsertColumnsBefore(2);

        await Assert.That(RefersTo(wb, "Cols")).IsEqualTo("Sheet1!$C:$G");
    }

    /// <summary>
    /// Inserting to the left of a column-only reference shifts it right without expanding.
    /// </summary>
    [Test]
    public async Task InsertingColumnsLeftOfColumnOnlyNamedRange_ShiftsRangeRight()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Cols", "Sheet1!$C:$E");

        ws.Column(1).InsertColumnsBefore(2);

        await Assert.That(RefersTo(wb, "Cols")).IsEqualTo("Sheet1!$E:$G");
    }

    /// <summary>
    /// Worksheet-scoped named ranges should also expand when rows are inserted.
    /// </summary>
    [Test]
    public async Task InsertingRowsInsideWorksheetScopedNamedRange_ExpandsRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var r = 1; r <= 8; r++)
        {
            ws.Cell(r, 1).Value = $"A{r}";
            ws.Cell(r, 2).Value = $"B{r}";
        }

        ws.Range("A1:B8").AddToNamed("Region", XLScope.Worksheet);

        ws.Row(3).InsertRowsAbove(6);

        var definedName = ws.DefinedNames.First(dn => dn.Name == "Region");
        await Assert.That(definedName.RefersTo).IsEqualTo("Sheet1!$A$1:$B$14").Because("Worksheet-scoped named range should expand when rows are inserted inside it");
    }
}
