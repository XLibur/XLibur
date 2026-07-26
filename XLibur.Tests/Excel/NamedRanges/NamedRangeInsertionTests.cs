using System.IO;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.NamedRanges;

public class NamedRangeInsertionTests
{
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
