using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The array-shaping functions: VSTACK, HSTACK, TOROW, TOCOL, WRAPROWS, WRAPCOLS, CHOOSEROWS,
/// CHOOSECOLS, TAKE, DROP and EXPAND. They are pure array-to-array transforms, so most of these
/// tests enter the formula as an array formula over the target range and read back the grid.
/// </summary>
[SetCulture("en-US")]
public class ArrayShapingTests
{
    private static IXLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return wb.AddWorksheet("Sheet1");
    }

    /// <summary>Fill A1:C2 with 1..6 read left to right, top to bottom.</summary>
    private static void SeedGrid(IXLWorksheet ws)
    {
        var value = 1;
        for (var row = 1; row <= 2; row++)
        {
            for (var column = 1; column <= 3; column++)
                ws.Cell(row, column).Value = value++;
        }
    }

    [Test]
    public async Task VStack_AppendsArraysOneBelowAnother()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("E1:F4").FormulaArrayA1 = "VSTACK({1,2;3,4}, {5,6;7,8})";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("F2").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("E3").Value).IsEqualTo(5d);
            await Assert.That((double)ws.Cell("F4").Value).IsEqualTo(8d);
        }
    }

    [Test]
    public async Task VStack_PadsNarrowerArgumentsWithNoValueAvailable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("E1:F3").FormulaArrayA1 = "VSTACK({1,2}, {3}, {4,5})";

            await Assert.That((double)ws.Cell("E2").Value).IsEqualTo(3d);
            await Assert.That(ws.Cell("F2").Value).IsEqualTo(XLError.NoValueAvailable);
            await Assert.That((double)ws.Cell("F3").Value).IsEqualTo(5d);
        }
    }

    [Test]
    public async Task HStack_AppendsArraysSideBySide()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("E1:H2").FormulaArrayA1 = "HSTACK({1;3}, {2;4}, {5;6}, {7;8})";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(2d);
            await Assert.That((double)ws.Cell("H2").Value).IsEqualTo(8d);
        }
    }

    [Test]
    public async Task HStack_PadsShorterArgumentsWithNoValueAvailable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("E1:F2").FormulaArrayA1 = "HSTACK({1;2}, {3})";

            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(3d);
            await Assert.That(ws.Cell("F2").Value).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task ToRow_ReadsTheGridLeftToRightByDefault()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:J1").FormulaArrayA1 = "TOROW(A1:C2)";

            for (var i = 0; i < 6; i++)
                await Assert.That((double)ws.Cell(1, 5 + i).Value).IsEqualTo(i + 1d);
        }
    }

    [Test]
    public async Task ToRow_ScansByColumnWhenAsked()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:J1").FormulaArrayA1 = "TOROW(A1:C2, 0, TRUE)";

            // Column order over 1 2 3 / 4 5 6 is 1 4 2 5 3 6.
            var expected = new[] { 1d, 4d, 2d, 5d, 3d, 6d };
            for (var i = 0; i < expected.Length; i++)
                await Assert.That((double)ws.Cell(1, 5 + i).Value).IsEqualTo(expected[i]);
        }
    }

    [Test]
    public async Task ToCol_ReadsTheGridIntoASingleColumn()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:E6").FormulaArrayA1 = "TOCOL(A1:C2)";

            for (var i = 0; i < 6; i++)
                await Assert.That((double)ws.Cell(1 + i, 5).Value).IsEqualTo(i + 1d);
        }
    }

    [Test]
    public async Task ToCol_IgnoresBlanksAndErrorsOnRequest()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            // A2 left blank.
            ws.Cell("A3").FormulaA1 = "1/0";
            ws.Cell("A4").Value = 2;

            ws.Range("C1:C4").FormulaArrayA1 = "TOCOL(A1:A4, 1)"; // Ignore blanks.
            ws.Range("D1:D4").FormulaArrayA1 = "TOCOL(A1:A4, 2)"; // Ignore errors.
            ws.Range("E1:E4").FormulaArrayA1 = "TOCOL(A1:A4, 3)"; // Ignore both.

            await Assert.That(ws.Cell("C2").Value).IsEqualTo(XLError.DivisionByZero);
            await Assert.That((double)ws.Cell("D2").Value).IsEqualTo(0d); // The blank survives as an empty value.
            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("E2").Value).IsEqualTo(2d);
            await Assert.That(ws.Cell("E3").Value).IsEqualTo(XLError.NoValueAvailable); // Past the end of the result.
        }
    }

    [Test]
    public async Task ToCol_WithEverythingIgnoredReturnsAnError()
    {
        // Excel reports #CALC! here; XLibur has no such error value and reports #VALUE!.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "1/0";
            ws.Cell("C1").FormulaA1 = "TOCOL(A1:A2, 3)";

            await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.IncompatibleValue);
        }
    }

    [Test]
    public async Task WrapRows_CutsAVectorIntoRows()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("E1:F3").FormulaArrayA1 = "WRAPROWS({1,2,3,4,5}, 2)";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(2d);
            await Assert.That((double)ws.Cell("E3").Value).IsEqualTo(5d);
            await Assert.That(ws.Cell("F3").Value).IsEqualTo(XLError.NoValueAvailable); // The short last row.
        }
    }

    [Test]
    public async Task WrapRows_PadsWithTheGivenValue()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("E1:F3").FormulaArrayA1 = "WRAPROWS({1,2,3,4,5}, 2, \"-\")";

            await Assert.That(ws.Cell("F3").Value).IsEqualTo("-");
        }
    }

    [Test]
    public async Task WrapCols_CutsAVectorIntoColumns()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("E1:G2").FormulaArrayA1 = "WRAPCOLS({1,2,3,4,5}, 2)";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("E2").Value).IsEqualTo(2d);
            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(3d);
            await Assert.That((double)ws.Cell("G1").Value).IsEqualTo(5d);
            await Assert.That(ws.Cell("G2").Value).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task Wrap_RejectsARectangleAndANonPositiveCount()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Cell("E1").FormulaA1 = "WRAPROWS(A1:C2, 2)";
            ws.Cell("E2").FormulaA1 = "WRAPROWS({1,2,3}, 0)";

            await Assert.That(ws.Cell("E1").Value).IsEqualTo(XLError.IncompatibleValue);
            await Assert.That(ws.Cell("E2").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task ChooseRows_PicksRowsInTheOrderAsked()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:G3").FormulaArrayA1 = "CHOOSEROWS(A1:C2, 2, 1, 2)";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("E2").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("G3").Value).IsEqualTo(6d); // Repeats are allowed.
        }
    }

    [Test]
    public async Task ChooseRows_CountsBackFromTheEndForNegativeIndices()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:G1").FormulaArrayA1 = "CHOOSEROWS(A1:C2, -1)";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(4d);
        }
    }

    [Test]
    public async Task ChooseCols_PicksColumns()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:F2").FormulaArrayA1 = "CHOOSECOLS(A1:C2, 3, 1)";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(3d);
            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("E2").Value).IsEqualTo(6d);
        }
    }

    [Test]
    public async Task ChooseRows_AcceptsAnArrayOfIndices()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:G2").FormulaArrayA1 = "CHOOSEROWS(A1:C2, {2,1})";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("E2").Value).IsEqualTo(1d);
        }
    }

    [Test]
    [Arguments("CHOOSEROWS(A1:C2, 0)")] // Rows are one-based.
    [Arguments("CHOOSEROWS(A1:C2, 3)")]
    [Arguments("CHOOSEROWS(A1:C2, -3)")]
    [Arguments("CHOOSECOLS(A1:C2, 4)")]
    public async Task Choose_OutOfRangeIndicesReturnIncompatibleValue(string formula)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Cell("E1").FormulaA1 = formula;

            await Assert.That(ws.Cell("E1").Value).IsEqualTo(XLError.IncompatibleValue);
        }
    }

    [Test]
    public async Task Take_KeepsFromTheStartAndTheEnd()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:F1").FormulaArrayA1 = "TAKE(A1:C2, 1, 2)"; // First row, first two columns.
            ws.Range("E3:F3").FormulaArrayA1 = "TAKE(A1:C2, -1, -2)"; // Last row, last two columns.

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(2d);
            await Assert.That((double)ws.Cell("E3").Value).IsEqualTo(5d);
            await Assert.That((double)ws.Cell("F3").Value).IsEqualTo(6d);
        }
    }

    [Test]
    public async Task Take_LeavesAnAxisAloneWhenItsCountIsOmitted()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:G1").FormulaArrayA1 = "TAKE(A1:C2, 1)"; // All three columns kept.
            ws.Range("E3:E4").FormulaArrayA1 = "TAKE(A1:C2, , 1)"; // Both rows kept.

            await Assert.That((double)ws.Cell("G1").Value).IsEqualTo(3d);
            await Assert.That((double)ws.Cell("E3").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("E4").Value).IsEqualTo(4d);
        }
    }

    [Test]
    public async Task Drop_DiscardsFromTheStartAndTheEnd()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:G1").FormulaArrayA1 = "DROP(A1:C2, 1)"; // Drop the first row.
            ws.Range("E3:F4").FormulaArrayA1 = "DROP(A1:C2, , -1)"; // Drop the last column.

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("E3").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("F4").Value).IsEqualTo(5d);
        }
    }

    [Test]
    public async Task TakeAndDrop_AreComplementary()
    {
        // Dropping the first row leaves what taking the last row leaves.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:G1").FormulaArrayA1 = "DROP(A1:C2, 1)";
            ws.Range("E2:G2").FormulaArrayA1 = "TAKE(A1:C2, -1)";

            for (var column = 5; column <= 7; column++)
                await Assert.That((double)ws.Cell(1, column).Value).IsEqualTo((double)ws.Cell(2, column).Value);
        }
    }

    [Test]
    public async Task Drop_ThatLeavesNothingReturnsAnError()
    {
        // Excel reports #CALC! here; XLibur has no such error value and reports #VALUE!.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Cell("E1").FormulaA1 = "DROP(A1:C2, 2)";
            ws.Cell("E2").FormulaA1 = "TAKE(A1:C2, 0)";

            await Assert.That(ws.Cell("E1").Value).IsEqualTo(XLError.IncompatibleValue);
            await Assert.That(ws.Cell("E2").Value).IsEqualTo(XLError.IncompatibleValue);
        }
    }

    [Test]
    public async Task Expand_GrowsAnArrayAndPadsTheNewCells()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("E1:G3").FormulaArrayA1 = "EXPAND({1,2;3,4}, 3, 3, \"-\")";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("F2").Value).IsEqualTo(4d);
            await Assert.That(ws.Cell("G1").Value).IsEqualTo("-");
            await Assert.That(ws.Cell("E3").Value).IsEqualTo("-");
        }
    }

    [Test]
    public async Task Expand_DefaultsToNoValueAvailableAndLeavesOmittedAxesAlone()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("E1:F3").FormulaArrayA1 = "EXPAND({1,2;3,4}, 3)";

            await Assert.That(ws.Cell("E3").Value).IsEqualTo(XLError.NoValueAvailable);
            await Assert.That((double)ws.Cell("F2").Value).IsEqualTo(4d);
        }
    }

    [Test]
    public async Task Expand_RefusesToShrink()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("E1").FormulaA1 = "EXPAND({1,2;3,4}, 1)";

            await Assert.That(ws.Cell("E1").Value).IsEqualTo(XLError.IncompatibleValue);
        }
    }

    [Test]
    public async Task ShapingFunctionsSpillIntoTheGrid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Cell("E1").SetDynamicFormulaA1("TOCOL(A1:C2)");

            for (var i = 0; i < 6; i++)
                await Assert.That((double)ws.Cell(1 + i, 5).Value).IsEqualTo(i + 1d);
        }
    }

    [Test]
    public async Task Shaping_RejectsArgumentsItCannotRead()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Cell("E1").FormulaA1 = "TOCOL(A1:C2, 4)"; // Only ignore modes 0..3 exist.
            ws.Cell("E2").FormulaA1 = "TOCOL(A1:C2, -1)";
            ws.Cell("E3").FormulaA1 = "WRAPROWS({1,2,3}, \"x\")"; // Not a count.
            ws.Cell("E4").FormulaA1 = "CHOOSEROWS(A1:C2, \"x\")";
            ws.Cell("E5").FormulaA1 = "TAKE(A1:C2, \"x\")";
            ws.Cell("E6").FormulaA1 = "EXPAND(A1:C2, \"x\")";
            ws.Cell("E7").FormulaA1 = "TOCOL(5)"; // A bare scalar is not an array.

            foreach (var address in new[] { "E1", "E2", "E7" })
                await Assert.That(ws.Cell(address).Value).IsEqualTo(XLError.IncompatibleValue);

            foreach (var address in new[] { "E3", "E4", "E5", "E6" })
                await Assert.That(ws.Cell(address).Value).IsEqualTo(XLError.IncompatibleValue);
        }
    }

    [Test]
    public async Task ToCol_ScansByColumnWhileIgnoringValues()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").FormulaA1 = "1/0";
            // A2 left blank.
            ws.Cell("B2").Value = 4;

            // Column order over the block is 1, blank, error, 4; ignoring both leaves 1 and 4.
            ws.Range("D1:D2").FormulaArrayA1 = "TOCOL(A1:B2, 3, TRUE)";

            await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("D2").Value).IsEqualTo(4d);
        }
    }

    [Test]
    public async Task Take_AskingForEverythingReturnsTheArrayUnchanged()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:G2").FormulaArrayA1 = "TAKE(A1:C2, 2, 3)";
            ws.Range("E4:G5").FormulaArrayA1 = "DROP(A1:C2, 0, 0)";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("G2").Value).IsEqualTo(6d);
            await Assert.That((double)ws.Cell("E4").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("G5").Value).IsEqualTo(6d);
        }
    }

    [Test]
    public async Task Expand_RefusesASizeBeyondTheSheet()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Cell("E1").FormulaA1 = "EXPAND(A1:C2, 2000000)";
            ws.Cell("E2").FormulaA1 = "EXPAND(A1:C2, 3, 20000)";

            await Assert.That(ws.Cell("E1").Value).IsEqualTo(XLError.NumberInvalid);
            await Assert.That(ws.Cell("E2").Value).IsEqualTo(XLError.NumberInvalid);
        }
    }

    [Test]
    public async Task ChooseCols_AcceptsAnArrayOfIndicesAndRejectsBadOnes()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Range("E1:F2").FormulaArrayA1 = "CHOOSECOLS(A1:C2, {3,-3})";
            ws.Cell("E4").FormulaA1 = "CHOOSECOLS(A1:C2, {1,9})";

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(3d);
            await Assert.That((double)ws.Cell("F1").Value).IsEqualTo(1d); // -3 is the first of three.
            await Assert.That(ws.Cell("E4").Value).IsEqualTo(XLError.IncompatibleValue);
        }
    }

    [Test]
    public async Task Stack_RejectsAScalarArgument()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Cell("E1").FormulaA1 = "VSTACK(A1:C2, 5)";

            await Assert.That(ws.Cell("E1").Value).IsEqualTo(XLError.IncompatibleValue);
        }
    }

    [Test]
    public async Task ShapingFunctionsCompose()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            // Flatten, wrap back into rows of three, and the original grid comes out again.
            ws.Cell("E1").SetDynamicFormulaA1("WRAPROWS(TOROW(A1:C2), 3)");

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("G1").Value).IsEqualTo(3d);
            await Assert.That((double)ws.Cell("E2").Value).IsEqualTo(4d);
            await Assert.That((double)ws.Cell("G2").Value).IsEqualTo(6d);
        }
    }

    [Test]
    public async Task StackingSpillsIntoTheGrid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            SeedGrid(ws);
            ws.Cell("E1").SetDynamicFormulaA1("VSTACK(A1:C2, A1:C2)");

            await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(1d);
            await Assert.That((double)ws.Cell("G4").Value).IsEqualTo(6d);
        }
    }
}
