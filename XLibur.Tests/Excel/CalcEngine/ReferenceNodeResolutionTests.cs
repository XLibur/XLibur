using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Covers how a <c>ReferenceNode</c> resolves to a range during evaluation. The node builds
/// the address from the area the parser produced and memoises the result, so these assert the
/// shapes that construction has to get right and the cases where a memo must not be reused.
/// </summary>
public class ReferenceNodeResolutionTests
{
    [Test]
    public async Task RelativeRange_ResolvesToAllCellsInArea()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 2;
        ws.Cell("A2").Value = 4;
        ws.Cell("B2").Value = 8;
        ws.Cell("D1").FormulaA1 = "SUM(A1:B2)";

        await Assert.That(ws.Cell("D1").Value).IsEqualTo(15);
    }

    [Test]
    public async Task AbsoluteRange_ResolvesToTheSameCellsAsTheRelativeForm()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 2;
        ws.Cell("A2").Value = 4;
        ws.Cell("B2").Value = 8;
        ws.Cell("D1").FormulaA1 = "SUM($A$1:$B$2)";

        await Assert.That(ws.Cell("D1").Value).IsEqualTo(15);
    }

    [Test]
    public async Task MixedAbsoluteAndRelativeEndpoints_ResolveToTheWholeArea()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 2;
        ws.Cell("A2").Value = 4;
        ws.Cell("B2").Value = 8;
        ws.Cell("D1").FormulaA1 = "SUM($A1:B$2)";

        await Assert.That(ws.Cell("D1").Value).IsEqualTo(15);
    }

    /// <summary>
    /// A column reference has no row axis, so the missing axis has to span the whole sheet
    /// rather than collapse to a single row.
    /// </summary>
    [Test]
    public async Task ColumnReference_SpansEveryRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 1;
        ws.Cell("A500").Value = 2;
        ws.Cell("B1").Value = 4;
        ws.Cell("C1").FormulaA1 = "SUM(A:B)";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(7);
    }

    [Test]
    public async Task RowReference_SpansEveryColumn()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 1;
        ws.Cell("ZZ1").Value = 2;
        ws.Cell("A2").Value = 4;
        ws.Cell("A4").FormulaA1 = "SUM(1:2)";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(7);
    }

    /// <summary>
    /// The parser hands back the endpoints in the order they were written, which need not be
    /// top-left then bottom-right. The address has to be normalized before use.
    /// </summary>
    [Test]
    public async Task ReversedEndpoints_ResolveToTheSameAreaAsTheNormalizedForm()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 2;
        ws.Cell("A2").Value = 4;
        ws.Cell("B2").Value = 8;
        ws.Cell("D1").FormulaA1 = "SUM(B2:A1)";

        await Assert.That(ws.Cell("D1").Value).IsEqualTo(15);
    }

    [Test]
    public async Task CrossSheetReference_ResolvesAgainstThePrefixedSheet()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        var main = wb.AddWorksheet("Main");
        data.Cell("A1").Value = 3;
        data.Cell("A2").Value = 4;
        main.Cell("A1").FormulaA1 = "SUM(Data!A1:A2)";

        await Assert.That(main.Cell("A1").Value).IsEqualTo(7);
    }

    [Test]
    public async Task CrossSheetReferenceToMissingSheet_IsRefError()
    {
        using var wb = new XLWorkbook();
        var main = wb.AddWorksheet("Main");
        main.Cell("A1").FormulaA1 = "SUM(Missing!A1:A2)";

        await Assert.That(main.Cell("A1").Value).IsEqualTo(XLError.CellReference);
    }

    /// <summary>
    /// The same formula text is shared by every cell that holds it, so one AST — and one
    /// <c>ReferenceNode</c> — serves them all. Each cell must still see the value of the
    /// range as evaluated for it.
    /// </summary>
    [Test]
    public async Task SharedFormulaText_ResolvesTheSameRangeForEveryCell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 5;
        ws.Cell("A2").Value = 6;
        ws.Cell("C1").FormulaA1 = "SUM($A$1:$A$2)";
        ws.Cell("C2").FormulaA1 = "SUM($A$1:$A$2)";
        ws.Cell("C3").FormulaA1 = "SUM($A$1:$A$2)";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(11);
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(11);
        await Assert.That(ws.Cell("C3").Value).IsEqualTo(11);
    }

    /// <summary>
    /// A resolved reference is memoised per node. Replacing the referenced sheet with a fresh
    /// one of the same name gives the prefix a different worksheet to resolve to, and the
    /// memo must not answer for the old one.
    /// </summary>
    [Test]
    public async Task ReplacingTheReferencedSheet_ResolvesAgainstTheNewSheet()
    {
        using var wb = new XLWorkbook();
        var main = wb.AddWorksheet("Main");
        var data = wb.AddWorksheet("Data");
        data.Cell("A1").Value = 1;
        main.Cell("A1").FormulaA1 = "Data!A1";

        await Assert.That(main.Cell("A1").Value).IsEqualTo(1);

        data.Delete();
        var replacement = wb.AddWorksheet("Data");
        replacement.Cell("A1").Value = 99;

        wb.RecalculateAllFormulas();

        await Assert.That(main.Cell("A1").Value).IsEqualTo(99);
    }

    /// <summary>
    /// Re-evaluating the same formula has to reflect the current contents of the referenced
    /// range, not the contents it had when the reference was first resolved.
    /// </summary>
    [Test]
    public async Task ReevaluatingAfterAnEdit_SeesTheUpdatedValues()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("C1").FormulaA1 = "SUM(A1:A2)";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(3);

        ws.Cell("A2").Value = 40;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(41);
    }
}
