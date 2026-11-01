using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class ArrayFormulaCalculationTests
{
    [Test]
    public async Task ScalarResultOfArrayFormulaIsCopiedAcrossCellGroup()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("C2:D4");

        range.FormulaArrayA1 = "ABS(-1)";

        foreach (var arrayFormulaCell in range.Cells())
        {
            await Assert.That(arrayFormulaCell.Value).IsEqualTo(1);
        }
    }

    [Test]
    public async Task SameShapeResultCausesEachCellOfCellGroupToUseCorrespondingValue()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:A2");

        range.FormulaArrayA1 = "TRANSPOSE({1,2})";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(2);
    }

    [Test]
    public async Task OnlyLeftmostValuesAreUsedWhenCellGroupHasFewerColumnsThanValue()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:C1");

        range.FormulaArrayA1 = "{1,2,3,4,5}";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(3);
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(Blank.Value);
    }

    [Test]
    public async Task OnlyTopmostValuesAreUsedWhenCellGroupHasFewerRowsThanValue()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:A3");

        range.FormulaArrayA1 = "{1;2;3;4;5}";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(3);
        await Assert.That(ws.Cell("A4").Value).IsEqualTo(Blank.Value);
    }

    [Test]
    public async Task SingleColumnValueIsClonedAcrossCellGroup()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:C3");

        range.FormulaArrayA1 = "{1;2}";

        for (var column = 1; column <= 3; column++)
        {
            await Assert.That(ws.Cell(1, column).Value).IsEqualTo(1);
            await Assert.That(ws.Cell(2, column).Value).IsEqualTo(2);
            await Assert.That(ws.Cell(3, column).Value).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task SingleRowValueIsClonedAcrossCellGroup()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:C3");

        range.FormulaArrayA1 = "{1,2}";

        for (var row = 1; row <= 3; row++)
        {
            await Assert.That(ws.Cell(row, 1).Value).IsEqualTo(1);
            await Assert.That(ws.Cell(row, 2).Value).IsEqualTo(2);
            await Assert.That(ws.Cell(row, 3).Value).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task ExcessColumnsAndRowsOfCellGroupTakeOnNoValueAvailable()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:C3");

        range.FormulaArrayA1 = "{1,2;3,4}";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(3);
        await Assert.That(ws.Cell("B2").Value).IsEqualTo(4);
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(ws.Cell("B3").Value).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(ws.Cell("C3").Value).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    public async Task Array_argument_for_scalar_function_in_array_formula_uses_only_first_value_of_array()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("B1:B3").FormulaArrayA1 = "SIGN({-1,2,0})";

        // The result is the row {-1,1,0}; a one-row result is cloned down a one-column cell group,
        // so every cell shows its first value.
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(-1);
        await Assert.That(ws.Cell("B2").Value).IsEqualTo(-1);
        await Assert.That(ws.Cell("B3").Value).IsEqualTo(-1);
    }

    [Test]
    public async Task Scalar_function_is_evaluated_for_each_element_of_an_array_argument()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:C1").FormulaArrayA1 = "SIGN({-1,2,0})";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(-1);
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(0);
    }

    [Test]
    public async Task Scalar_function_is_evaluated_for_each_cell_of_a_range_argument()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = -1;
        ws.Cell("A2").Value = -2;
        ws.Range("B1:B2").FormulaArrayA1 = "ABS(A1:A2)";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("B2").Value).IsEqualTo(2);
    }

    [Test]
    public async Task Text_function_is_evaluated_for_each_element_of_an_array_argument()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:B1").FormulaArrayA1 = "LEN({\"ab\",\"c\"})";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
    }

    [Test]
    public async Task Scalar_arguments_of_different_shapes_are_broadcast_per_element()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:B2").FormulaArrayA1 = "POWER({2,3},{1;2})";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(3);
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(4);
        await Assert.That(ws.Cell("B2").Value).IsEqualTo(9);
    }
}
