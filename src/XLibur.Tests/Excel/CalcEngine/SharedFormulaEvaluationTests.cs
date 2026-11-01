using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;
using XLibur.Excel.IO;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// A cell of a shared formula is evaluated from the one R1C1 tree of its group, not from its own A1
/// text (#686). The loader wrote that A1 text from the same R1C1 text, so both must give the same
/// value in every cell, including where a reference moves off the sheet and the A1 text says
/// <c>#REF!</c>.
/// </summary>
public class SharedFormulaEvaluationTests
{
    private const int LastRow = XLHelper.MaxRowNumber;
    private const int LastColumn = XLHelper.MaxColumnNumber;

    public static IEnumerable<object[]> Groups()
    {
        int[] top = [1, 2, 3];
        int[] bottom = [LastRow - 2, LastRow - 1, LastRow];

        // A relative column off the left and the right edge.
        yield return ["RC[-1]*2", 3, top];
        yield return ["RC[-1]*2", 1, top];
        yield return ["RC[1]*2", LastColumn, top];

        // An area whose first end moves off the top, and whose second end moves off the bottom.
        yield return ["SUM(R[-1]C[-2]:R[1]C[-1])", 3, top];
        yield return ["SUM(R[-1]C[-2]:R[1]C[-1])", 3, bottom];

        // Whole columns and whole rows.
        yield return ["SUM(C[-1])", 3, top];
        yield return ["SUM(C[-1])", 1, top];
        yield return ["SUM(R[1])", 3, top];
        yield return ["SUM(R[1])", 3, bottom];

        // Mixed axes, and no relative axis at all.
        yield return ["R1C[-1]+R[1]C1", 3, top];
        yield return ["R1C[-1]+R[1]C1", 3, bottom];
        yield return ["R1C2*2", 3, top];

        // Another sheet, and a bang reference.
        yield return ["Data!R[-1]C*2", 2, top];
        yield return ["!RC[-1]+1", 3, top];
        yield return ["!RC[-1]+1", 1, top];

        // Operators on references: a range between two references, and an intersection.
        yield return ["SUM(RC[-2]:RC[-1],R[1]C[-1])", 3, bottom];
        yield return ["SUM(R1C1:R3C2 RC[-2]:RC[-1])", 3, top];
    }

    [Test]
    [MethodDataSource(nameof(Groups))]
    public async Task Each_cell_of_a_shared_formula_evaluates_as_its_A1_text_does(string r1c1, int column, int[] rows)
    {
        using var wb = WorkbookWithValues(column, rows);
        var sheet = (XLWorksheet)wb.Worksheet("Sheet1");
        var cells = SetShared(sheet, r1c1, column, rows);

        // Read every cell first: a formula that reads another formula is then clean when the A1
        // text is evaluated on its own below.
        var fromGroup = new List<XLCellValue>();
        foreach (var cell in cells)
            fromGroup.Add(cell.Value);

        for (var i = 0; i < cells.Count; i++)
        {
            var cell = cells[i];
            var fromText = wb.CalcEngine.EvaluateFormula(cell.FormulaA1, wb, sheet, cell.SheetPoint).ToCellValue();
            await Assert.That(fromGroup[i]).IsEqualTo(fromText).Because($"{cell.Address} holds {cell.FormulaA1}");

            // The group's tree was read, and the cell's own A1 text was never parsed.
            await Assert.That(cell.Formula!.HasAst).IsFalse();
        }
    }

    /// <summary>
    /// A cell called as a function, such as <c>A1(100)</c>, is <c>#REF!</c>, except in the cell where
    /// the reference lands on a name such as <c>LOG10</c>, which is that function. The R1C1 tree
    /// cannot tell the two apart, so the parser refuses it and each cell reads its own A1 text.
    /// </summary>
    [Test]
    public async Task A_cell_called_as_a_function_is_read_from_the_A1_text_of_each_cell()
    {
        // Column LOG is 8509, so from column 8510 of row 11 the reference R[-1]C[-1] is LOG10.
        using var wb = WorkbookWithValues(8510, [11, 12]);
        var sheet = (XLWorksheet)wb.Worksheet("Sheet1");
        var cells = SetShared(sheet, "R[-1]C[-1](100)", 8510, [11, 12]);

        await Assert.That(cells[0].FormulaA1).IsEqualTo("LOG10(100)");
        await Assert.That(cells[0].Value).IsEqualTo(2);
        await Assert.That(cells[1].Value).IsEqualTo(XLError.CellReference);

        await Assert.That(cells[0].Formula!.TryGetShared(cells[0].SheetPoint, out var group)).IsTrue();
        await Assert.That(group!.TryGetAst(wb.CalcEngine, out _)).IsFalse();
    }

    /// <summary>
    /// The recalculation pass evaluates through the group's tree as well. Each cell reads the one above,
    /// so a read of the last falls back to the pass, which calculates the column from the top.
    /// </summary>
    [Test]
    public async Task A_recalculation_evaluates_each_cell_from_the_group()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("C1").Value = 1;
        var cells = SetShared((XLWorksheet)ws, "R[-1]C+1", 3, [2, 3, 4, 5]);

        await Assert.That(cells[^1].Value).IsEqualTo(5);

        ws.Cell("C1").Value = 10;

        await Assert.That(cells[^1].Value).IsEqualTo(14);
        foreach (var cell in cells)
            await Assert.That(cell.Formula!.HasAst).IsFalse();
    }

    /// <summary>
    /// Every cell of a shared formula evaluates the one tree of its group: nothing is parsed per cell.
    /// </summary>
    [Test]
    public async Task A_loaded_shared_formula_is_evaluated_from_one_tree()
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            for (var row = 1; row <= 3; row++)
            {
                ws.Cell(row, 1).Value = row;
                ws.Cell(row, 2).FormulaA1 = $"A{row}*10";
            }

            wb.SaveAs(package);
        }

        package.RewriteSheet1(xml => xml
            .Replace("<x:f>A1*10</x:f>", "<x:f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*10</x:f>", StringComparison.Ordinal)
            .Replace("<x:f>A2*10</x:f>", "<x:f t=\"shared\" si=\"0\" />", StringComparison.Ordinal)
            .Replace("<x:f>A3*10</x:f>", "<x:f t=\"shared\" si=\"0\" />", StringComparison.Ordinal));

        package.Position = 0;
        using var loaded = new XLWorkbook(package);
        var sheet = loaded.Worksheet("Sheet1");

        Formula? groupAst = null;
        for (var row = 1; row <= 3; row++)
        {
            var cell = (XLCell)sheet.Cell(row, 2);
            await Assert.That(cell.FormulaA1).IsEqualTo($"A{row}*10");
            await Assert.That(cell.Value).IsEqualTo(row * 10);
            await Assert.That(cell.Formula!.HasAst).IsFalse();

            var ast = cell.Formula!.GetAst(loaded.CalcEngine, cell.SheetPoint);
            groupAst ??= ast;
            await Assert.That(ast).IsSameReferenceAs(groupAst);
        }
    }

    /// <summary>
    /// A row inserted above moves a cell of a shared formula that reads another sheet, and its A1 text
    /// stays the same. The group's tree would read the row below at the new cell, so the cell reads
    /// its own text.
    /// </summary>
    [Test]
    public async Task A_moved_cell_of_a_shared_formula_reads_its_own_A1_text()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var data = wb.AddWorksheet("Data");
        data.Cell("A2").Value = 7;
        data.Cell("A3").Value = 9;
        var cells = SetShared((XLWorksheet)ws, "Data!RC[-2]*2", 3, [2, 3]);
        await Assert.That(cells[0].Value).IsEqualTo(14);

        ws.Row(1).InsertRowsAbove(1);

        var moved = (XLCell)ws.Cell("C3");
        await Assert.That(moved.FormulaA1).IsEqualTo("Data!A2*2");
        await Assert.That(moved.Formula!.TryGetShared(moved.SheetPoint, out _)).IsFalse();
        await Assert.That(moved.Value).IsEqualTo(14);
    }

    /// <summary>
    /// Put one shared formula in <paramref name="rows"/> of <paramref name="column"/>, as the loader
    /// does: each cell gets the A1 text that the loader writes, and the group.
    /// </summary>
    private static List<XLCell> SetShared(XLWorksheet sheet, string r1c1, int column, int[] rows)
    {
        var shared = new WorksheetSheetDataReader.SharedFormula(r1c1);
        var cells = new List<XLCell>();
        foreach (var row in rows)
        {
            var point = new Point(row, column);
            var formula = XLCellFormula.NormalA1(shared.ToA1(point));
            formula.SetShared(shared.Group, point);

            var cell = (XLCell)sheet.Cell(row, column);
            cell.Formula = formula;
            cells.Add(cell);
        }

        return cells;
    }

    /// <summary>
    /// A workbook with a number in every cell within two rows and two columns of the formula cells,
    /// except the formula cells themselves, and in the first rows and columns of the sheet.
    /// </summary>
    private static XLWorkbook WorkbookWithValues(int column, int[] rows)
    {
        var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet("Sheet1");
        var data = wb.AddWorksheet("Data");
        var formulaCells = new HashSet<Point>();
        foreach (var row in rows)
            formulaCells.Add(new Point(row, column));

        foreach (var row in rows)
        {
            for (var r = Math.Max(1, row - 2); r <= Math.Min(LastRow, row + 2); r++)
            {
                for (var c = Math.Max(1, column - 2); c <= Math.Min(LastColumn, column + 2); c++)
                {
                    if (!formulaCells.Contains(new Point(r, c)))
                        sheet.Cell(r, c).Value = r * 7 + c;

                    data.Cell(r, c).Value = r * 11 + c;
                }
            }
        }

        for (var i = 1; i <= 3; i++)
        {
            sheet.Cell(i, 1).Value = i * 3;
            sheet.Cell(1, i).Value = i * 5;
        }

        return wb;
    }
}
