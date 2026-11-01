using System;
using System.IO;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class ArrayFormulaTests
{
    [Test]
    public async Task ArrayFormulaIsSaved()
    {
        await TestHelper.CreateAndCompare(wb =>
        {
            var ws = wb.AddWorksheet();
            ws.Range("A1:B2").FormulaArrayA1 = "1+2";
        }, @"Other\Formulas\ArrayFormula.xlsx");
    }

    [Test]
    public async Task ArrayFormulaCanBeLoaded()
    {
        await TestHelper.LoadAndAssert(async wb =>
        {
            var ws = wb.Worksheets.First();

            foreach (var arrayFormulaCell in ws.Range("A1:B2").Cells())
            {
                await Assert.That(arrayFormulaCell.FormulaA1).IsEqualTo("1+2");
                await Assert.That(arrayFormulaCell.FormulaReference!.ToStringRelative()).IsEqualTo("A1:B2");
            }

            var outsideCell = ws.Cell("A3");
            await Assert.That(outsideCell.FormulaA1).IsEmpty();
            await Assert.That(outsideCell.FormulaReference).IsNull();
        }, @"Other\Formulas\ArrayFormula.xlsx");
    }

    [Test]
    public async Task CanBeOnlyForOneCell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var oneCell = ws.Cell("B3");

        oneCell.AsRange().FormulaArrayA1 = "2+5";

        await Assert.That(oneCell.HasArrayFormula).IsTrue();
        await Assert.That(oneCell.FormulaA1).IsEqualTo("2+5");
        await Assert.That(oneCell.FormulaReference!.ToStringRelative()).IsEqualTo("B3:B3");
    }

    [Test]
    [Arguments("B2:C3")]
    [Arguments("B2:C4")]
    [Arguments("A1:D7")]
    public async Task SettingValueToContainingRangeClearsArrayFormula(string containingRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var arrayFormulaRange = ws.Range("B2:C3");
        arrayFormulaRange.FormulaArrayA1 = "5";

        ws.Range(containingRange).Value = Blank.Value;

        foreach (var cell in arrayFormulaRange.Cells())
        {
            await Assert.That(cell.Value).IsEqualTo(Blank.Value);
            await Assert.That(cell.HasArrayFormula).IsFalse();
            await Assert.That(cell.FormulaA1).IsEmpty();
            await Assert.That(cell.FormulaReference).IsNull();
        }
    }

    [Test]
    [Arguments("B2:D3")]
    [Arguments("A1:E4")]
    public async Task SettingFormulaToContainingRangeClearsOriginalArrayFormula(string overlapRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("B2:D3").FormulaArrayA1 = "1";

        await Assert.That(() => ws.Range(overlapRange).FormulaArrayA1 = "2").ThrowsNothing();
    }

    [Test]
    [Arguments("B2:B2")]
    [Arguments("B2:B3")]
    [Arguments("A1:C3")]
    [Arguments("D2:F3")]
    [Arguments("C:C")]
    [Arguments("2:2")]
    public async Task ArrayFormulaCantPartiallyOverlapWithAnotherArrayFormula(string partialOverlapRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("B2:D3").FormulaArrayA1 = "1";

        var ex = await Assert.That(() => ws.Range(partialOverlapRange).FormulaArrayA1 = "2").Throws<InvalidOperationException>();
        await Assert.That(ex!.Message).IsEqualTo("Can't create array function that partially covers another array function.");
    }

    [Test]
    [Arguments("A1:B2")]
    [Arguments("A2")]
    public async Task ArrayFormulaCantOverlapWithMergedRange(string partialOverlapRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:A2").Merge();

        var ex = await Assert.That(() => ws.Range(partialOverlapRange).FormulaArrayA1 = "1").Throws<InvalidOperationException>();
        await Assert.That(ex!.Message).IsEqualTo("Can't create array function over a merged range.");
    }

    [Test]
    [Arguments("A1:B2")]
    [Arguments("A1:C1")]
    public async Task ArrayFormulaCantOverlapWithTable(string formulaRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Name";
        ws.Cell("A2").Value = 5;
        ws.Range("A1:A2").CreateTable();

        var ex = await Assert.That(() => ws.Range(formulaRange).FormulaArrayA1 = "1").Throws<InvalidOperationException>();
        await Assert.That(ex!.Message).IsEqualTo("Can't create array function over a table.");
    }

    [Test]
    public async Task SettingArrayFormulaInvalidatesCells()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("A2").NeedsRecalculation).IsFalse();

        ws.Range("A1:A2").FormulaArrayA1 = "ABS(-3)";

        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A2").NeedsRecalculation).IsTrue();
    }

    [Test]
    public async Task ReferencingItselfIsCircularError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "A2";
        ws.Range("A2").FormulaArrayA1 = "A1";

        var ex = await Assert.That(() => _ = ws.Cell("A2").Value).Throws<InvalidOperationException>();
        await Assert.That(ex!.Message).IsEqualTo("Formula in a cell '$Sheet1'!$A1 is part of a cycle.");
    }

    [Test]
    public async Task ArrayFormulaCachedValues_WrittenToXml()
    {
        // Verify that cached values for array formula cells (both master and child)
        // are written to the XML even when EvaluateFormulasBeforeSaving is false.
        using var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:A3").FormulaArrayA1 = "TRANSPOSE({10,20,30})";

        // Evaluate all cells so cached values are populated
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(10.0);
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(20.0);
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(30.0);

        wb.SaveAs(ms, validate: false);

        // Extract and check the XML content
        var bytes = ms.ToArray();
        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(bytes), System.IO.Compression.ZipArchiveMode.Read);
        var sheetEntry = zip.Entries.First(e => e.FullName.Contains("sheet1.xml", StringComparison.OrdinalIgnoreCase));
        using var sr = new StreamReader(sheetEntry.Open());
        var sheetXml = sr.ReadToEnd();

        // All three cells should have their distinct cached values in the XML.
        // Previously, only the master cell (A1) would have a value, and child cells
        // (A2, A3) would be empty because cached values were only written when
        // EvaluateFormulasBeforeSaving was true.
        await Assert.That(sheetXml).Contains("<x:v>10</x:v>").Because("Master cell A1 value missing from XML");
        await Assert.That(sheetXml).Contains("<x:v>20</x:v>").Because("Child cell A2 value missing from XML");
        await Assert.That(sheetXml).Contains("<x:v>30</x:v>").Because("Child cell A3 value missing from XML");
    }

    [Test]
    public async Task NormalFormulaCachedValues_PreservedOnRoundTrip()
    {
        // Verify that non-array formula cells also preserve cached values
        // without requiring EvaluateFormulasBeforeSaving.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 10;
            ws.Cell("B1").FormulaA1 = "A1*2";
            // Evaluate to populate cached value
            ws.Cell("B1").Value.ToString();

            wb.SaveAs(ms, false);
        }

        ms.Position = 0;

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.Cell("B1").CachedValue).IsEqualTo(20.0);
        }
    }

    [Test]
    public async Task InsertingRowsInAnotherSheetKeepsArrayFormulaIntact()
    {
        // Regression: inserting rows/columns anywhere in the workbook used to route every
        // formula cell through the FormulaA1 setter, rebuilding a *normal* formula per cell.
        // For an array formula (which shares one instance across its whole range) this split a
        // single spilled array (e.g. =UNIQUE(...)) into N implicit-intersection =@UNIQUE(...)
        // cells, even when the insert happened on an unrelated sheet.
        using var wb = new XLWorkbook();
        var dataSheet = wb.AddWorksheet("Data");
        var arraySheet = wb.AddWorksheet("Calc");
        arraySheet.Range("A1:A3").FormulaArrayA1 = "TRANSPOSE({10,20,30})";

        dataSheet.Row(1).InsertRowsAbove(5);

        foreach (var cell in arraySheet.Range("A1:A3").Cells())
        {
            await Assert.That(cell.HasArrayFormula).IsTrue().Because($"{cell.Address} lost its array formula");
            await Assert.That(cell.FormulaReference!.ToStringRelative()).IsEqualTo("A1:A3");
        }
    }

    [Test]
    public async Task InsertingRowsAboveShiftsArrayFormulaRange()
    {
        // A same-sheet insert above the array must relocate the array's spill range so the
        // master cell is still identifiable (otherwise the formula vanishes on save).
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("B3:B5").FormulaArrayA1 = "TRANSPOSE({1,2,3})";

        ws.Row(1).InsertRowsAbove(2);

        foreach (var cell in ws.Range("B5:B7").Cells())
        {
            await Assert.That(cell.HasArrayFormula).IsTrue().Because($"{cell.Address} lost its array formula");
            await Assert.That(cell.FormulaReference!.ToStringRelative()).IsEqualTo("B5:B7");
        }
    }

    [Test]
    public async Task InsertingColumnsBeforeShiftsArrayFormulaRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("B3:D3").FormulaArrayA1 = "{1,2,3}";

        ws.Column(1).InsertColumnsBefore(1);

        foreach (var cell in ws.Range("C3:E3").Cells())
        {
            await Assert.That(cell.HasArrayFormula).IsTrue().Because($"{cell.Address} lost its array formula");
            await Assert.That(cell.FormulaReference!.ToStringRelative()).IsEqualTo("C3:E3");
        }
    }

    [Test]
    public async Task ArrayFormulaSurvivesInsertOnSaveAsSingleFormula()
    {
        // End-to-end: after an unrelated insert, the saved sheet must still contain exactly one
        // array formula element (on the master cell), not one normal formula per cell.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var dataSheet = wb.AddWorksheet("Data");
            var arraySheet = wb.AddWorksheet("Calc");
            arraySheet.Range("A1:A3").FormulaArrayA1 = "TRANSPOSE({10,20,30})";

            dataSheet.Row(1).InsertRowsAbove(3);

            wb.SaveAs(ms, validate: false);
        }

        var bytes = ms.ToArray();
        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(bytes), System.IO.Compression.ZipArchiveMode.Read);
        var sheetEntry = zip.Entries.First(e => e.FullName.Contains("sheet2.xml", StringComparison.OrdinalIgnoreCase));
        using var sr = new StreamReader(sheetEntry.Open());
        var sheetXml = sr.ReadToEnd();

        // Exactly one array-formula element, referencing the whole spill range.
        var arrayCount = sheetXml.Split("t=\"array\"").Length - 1;
        await Assert.That(arrayCount).IsEqualTo(1).Because("Array formula was split into multiple per-cell formulas");
        await Assert.That(sheetXml).Contains("ref=\"A1:A3\"");
    }

    [Test]
    public async Task DynamicArrayFormulaKeepsDynamicFlagWhenShifted()
    {
        // A dynamic array is stored as a normal formula with the dynamic-array flag set.
        // When a shift changes the referenced cells, the formula must stay dynamic so the
        // saved cell keeps its cm metadata link and Excel does not apply implicit
        // intersection (=@UNIQUE(...)).
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("A3").Value = 2;
            ws.Cell("C1").SetDynamicFormulaA1("UNIQUE(A1:A3)");

            ws.Row(1).InsertRowsAbove(1); // C1 -> C2, references A1:A3 -> A2:A4

            wb.SaveAs(ms, validate: false);
        }

        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(ms.ToArray()), System.IO.Compression.ZipArchiveMode.Read);

        var sheetEntry = zip.Entries.First(e => e.FullName.Contains("sheet1.xml", StringComparison.OrdinalIgnoreCase));
        using var sheetReader = new StreamReader(sheetEntry.Open());
        var sheetXml = sheetReader.ReadToEnd();

        // The shifted formula still carries the dynamic-array cell-metadata link.
        await Assert.That(sheetXml).Contains("_xlfn.UNIQUE(A2:A4)").Because("Reference shift did not apply");
        await Assert.That(sheetXml).Contains("cm=\"1\"").Because("Dynamic-array cell metadata (cm) was lost on shift");

        // The dynamic-array metadata part is present.
        var metadataEntry = zip.Entries.First(e => e.FullName.Contains("metadata", StringComparison.OrdinalIgnoreCase));
        using var metadataReader = new StreamReader(metadataEntry.Open());
        var metadataXml = metadataReader.ReadToEnd();
        await Assert.That(metadataXml).Contains("fDynamic=\"1\"").Because("Dynamic-array metadata missing");
    }

    [Test]
    public async Task DynamicArrayFormulaSavesAsArrayFormulaWithSpillRef()
    {
        // A spilled dynamic array serialises as an array formula whose ref is the spill
        // footprint, on the anchor cell, plus the cm dynamic-array metadata link.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");
            wb.RecalculateAllFormulas(); // spills A1:A3

            wb.SaveAs(ms, validate: false);
        }

        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(ms.ToArray()), System.IO.Compression.ZipArchiveMode.Read);
        var sheetEntry = zip.Entries.First(e => e.FullName.Contains("sheet1.xml", StringComparison.OrdinalIgnoreCase));
        using var sheetReader = new StreamReader(sheetEntry.Open());
        var sheetXml = sheetReader.ReadToEnd();

        await Assert.That(sheetXml).Contains("t=\"array\"").Because("Dynamic array must serialise as an array formula");
        await Assert.That(sheetXml).Contains("ref=\"A1:A3\"").Because("Spill footprint must be written as the array ref");
        await Assert.That(sheetXml).Contains("cm=\"1\"").Because("Dynamic-array cell metadata (cm) missing");
    }

    [Test]
    public async Task DynamicArrayFormulaRoundTripsAndReSpills()
    {
        // Full round-trip: save a spilled dynamic array, load it back, and confirm it is
        // reconstructed as a dynamic array (not a legacy CSE array) that re-spills correctly.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("D1").Value = 5;
            ws.Cell("D2").Value = 6;
            ws.Cell("D3").Value = 7;
            ws.Cell("A1").SetDynamicFormulaA1("UNIQUE(D1:D3)");
            wb.RecalculateAllFormulas(); // spills A1:A3 = {5;6;7}

            wb.SaveAs(ms, validate: false);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet("S");
            var anchor = (XLCell)ws.Cell("A1");

            await Assert.That(anchor.HasFormula).IsTrue();
            await Assert.That(anchor.Formula!.IsDynamicArray).IsTrue().Because("Loaded formula must be a dynamic array, not a CSE array");
            await Assert.That(((XLCell)ws.Cell("A2")).HasFormula).IsFalse().Because("Spilled cell must load formula-less");

            // The spill re-evaluates and fills the footprint from the cached child values
            // without a #SPILL! collision.
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(5);
            await Assert.That(ws.Cell("A3").Value).IsEqualTo(7);

            // Changing a source re-spills the loaded formula.
            ws.Cell("D3").Value = 8;
            wb.RecalculateAllFormulas();
            await Assert.That(ws.Cell("A3").Value).IsEqualTo(8);
        }
    }

    [Test]
    public async Task DynamicArrayAndCseArrayLoadDistinctly()
    {
        // A dynamic array (cm metadata) and a legacy CSE array (no cm) both serialise with
        // t="array"; on load only the one whose cm references XLDAPR becomes dynamic.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(2)"); // dynamic, A1:A2
            ws.Range("C1:C2").FormulaArrayA1 = "TRANSPOSE({10,20})"; // CSE array, C1:C2
            wb.RecalculateAllFormulas();

            wb.SaveAs(ms, validate: false);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet("S");

            await Assert.That(((XLCell)ws.Cell("A1")).Formula!.IsDynamicArray).IsTrue().Because("SEQUENCE must load as a dynamic array");
            await Assert.That(((XLCell)ws.Cell("A2")).HasFormula).IsFalse().Because("Dynamic spill cell is formula-less");

            await Assert.That(((XLCell)ws.Cell("C1")).Formula!.IsDynamicArray).IsFalse().Because("CSE array must not load as dynamic");
            await Assert.That(ws.Cell("C1").HasArrayFormula).IsTrue().Because("CSE array keeps its array formula");
            await Assert.That(ws.Cell("C2").HasArrayFormula).IsTrue().Because("CSE array child keeps the shared array formula");
        }
    }

    [Test]
    public async Task DynamicArraySpillErrorRoundTrips()
    {
        // A blocked dynamic array (#SPILL! anchor) round-trips: the error value survives save/load
        // (exercising the XLError.SpillRange save/parse path) and stays blocked after reload.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A2").Value = "block";
            ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");
            wb.RecalculateAllFormulas();
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.SpillRange);

            wb.SaveAs(ms, validate: false);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheet("S");
            await Assert.That(ws.Cell("A2").Value).IsEqualTo("block");
            await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.SpillRange);
        }
    }

    [Test]
    public async Task DeletingRowsThroughArrayDoesNotCorruptRange()
    {
        // Deleting rows that overlap an array used to push the stored range past row 1, producing
        // an out-of-bounds coordinate (e.g. A0:A2) via the unchecked Point constructor.
        // Excel forbids editing part of an array; XLibur must at least keep a valid range and a
        // saveable workbook rather than silently corrupting it.
        using var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Range("A2:A4").FormulaArrayA1 = "TRANSPOSE({1,2,3})";

        await Assert.That(() => ws.Rows(1, 2).Delete()).ThrowsNothing();

        foreach (var cell in ws.CellsUsed(c => c.HasArrayFormula))
        {
            var reference = cell.FormulaReference!;
            await Assert.That(reference.FirstAddress.RowNumber).IsGreaterThanOrEqualTo(1).Because($"{cell.Address} array range has an out-of-bounds row: {reference.ToStringRelative()}");
            await Assert.That(reference.FirstAddress.ColumnNumber).IsGreaterThanOrEqualTo(1).Because($"{cell.Address} array range has an out-of-bounds column: {reference.ToStringRelative()}");
        }

        await Assert.That(() => wb.SaveAs(ms, validate: false)).ThrowsNothing();
    }

    [Test]
    public async Task DeletingColumnsThroughArrayDoesNotCorruptRange()
    {
        using var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Range("B1:D1").FormulaArrayA1 = "{1,2,3}";

        await Assert.That(() => ws.Columns(1, 2).Delete()).ThrowsNothing(); // delete A:B, overlaps the array's left edge

        foreach (var cell in ws.CellsUsed(c => c.HasArrayFormula))
        {
            var reference = cell.FormulaReference!;
            await Assert.That(reference.FirstAddress.ColumnNumber).IsGreaterThanOrEqualTo(1).Because($"{cell.Address} array range has an out-of-bounds column: {reference.ToStringRelative()}");
            await Assert.That(reference.FirstAddress.RowNumber).IsGreaterThanOrEqualTo(1).Because($"{cell.Address} array range has an out-of-bounds row: {reference.ToStringRelative()}");
        }

        await Assert.That(() => wb.SaveAs(ms, validate: false)).ThrowsNothing();
    }
}
