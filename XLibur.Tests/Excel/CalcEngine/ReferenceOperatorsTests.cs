using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class ReferenceOperatorsTests
{
    #region Implicit intersection

    [Test]
    public async Task ImplicitIntersection_DoesNotAffectSingleCellReference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("D5").FormulaA1 = "ABS(B3:B3)";

        await Assert.That(ws.Cell("D5").Value).IsEqualTo(1);
    }

    [Test]
    public async Task ImplicitIntersection_TakesReferenceFromHorizontalLine()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("D3").FormulaA1 = "ABS(B1:B10)";

        await Assert.That(ws.Cell("D3").Value).IsEqualTo(1);
    }

    [Test]
    public async Task ImplicitIntersection_TakesReferenceFromVerticalLine()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("B5").FormulaA1 = "ABS(A3:Z3)";

        await Assert.That(ws.Cell("B5").Value).IsEqualTo(1);
    }

    [Test]
    public async Task ImplicitIntersection_TakesReferenceEvenFromIntersectionEvenFromDifferentSheet()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        sheet1.Cell("B3").Value = -1;

        var sheet2 = wb.AddWorksheet("Sheet2");
        sheet2.Cell("D3").FormulaA1 = "ABS(Sheet1!B1:B10)";

        await Assert.That(sheet2.Cell("D3").Value).IsEqualTo(1);
    }

    [Test]
    public async Task ImplicitIntersection_WithoutIntersectionResultsInValueError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("D5").FormulaA1 = "ABS(B1:B4)";

        await Assert.That(ws.Cell("D5").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task ImplicitIntersection_CanWorkOnlyWithOneArea()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("D3").FormulaA1 = "ABS((B1:B2,B3:B5))"; // A continous range made of two areas

        await Assert.That(ws.Cell("D3").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task ImplicitIntersection_IntersectionMustHaveSpanOfOneCell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        var horizontalIntersectionCell = ws.Cell("D3");
        horizontalIntersectionCell.FormulaA1 = "ABS(A1:B5)";
        await Assert.That(horizontalIntersectionCell.Value).IsEqualTo(XLError.IncompatibleValue);

        var verticalIntersectionCell = ws.Cell("B5");
        verticalIntersectionCell.FormulaA1 = "ABS(A3:C4)";
        await Assert.That(verticalIntersectionCell.Value).IsEqualTo(XLError.IncompatibleValue);
    }

    #endregion

    #region Implicit intersection operator @

    // The explicit operator. `@` binds looser than `:` and the space, so `@A1:A4` is `@(A1:A4)` and
    // `D3:@A1:A2` is `D3:(@(A1:A2))`. A range gives the cell on the formula's row or column, an array
    // its top-left element, and anything else itself.

    [Test]
    public async Task ImplicitIntersectionOperator_TakesTheCellOnTheFormulasRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = 7;
        ws.Cell("C3").FormulaA1 = "@B1:B10";

        await Assert.That(ws.Cell("C3").Value).IsEqualTo(7);
    }

    [Test]
    public async Task ImplicitIntersectionOperator_TakesTheCellInTheFormulasColumn()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = 7;
        ws.Cell("B5").FormulaA1 = "@A3:Z3";

        await Assert.That(ws.Cell("B5").Value).IsEqualTo(7);
    }

    [Test]
    public async Task ImplicitIntersectionOperator_WorksInAFunctionArgument()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 20;
        ws.Cell("A3").Value = 300;
        ws.Cell("A4").Value = 4000;
        ws.Cell("C2").FormulaA1 = "SUM(@A1:A4)";
        ws.Cell("C3").FormulaA1 = "IF(@A1:A4=300,\"yes\",\"no\")";

        await Assert.That(ws.Cell("C2").Value).IsEqualTo(20);
        await Assert.That(ws.Cell("C3").Value).IsEqualTo("yes");
    }

    /// <summary>
    /// The operator gives a reference, not a value, so it can be the operand of the range operator
    /// and the argument of a function that needs a reference.
    /// </summary>
    [Test]
    public async Task ImplicitIntersectionOperator_GivesAReference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        for (var row = 1; row <= 3; row++)
        {
            for (var column = 1; column <= 4; column++)
                ws.Cell(row, column).Value = (row - 1) * 4 + column;
        }

        // D3:(@(A1:A2)) on row 2 is D3:A2, which is A2:D3: 5 + 6 + ... + 12.
        ws.Cell("F2").FormulaA1 = "SUM(D3:@A1:A2)";
        ws.Cell("F4").FormulaA1 = "ROW(@A1:A10)";

        await Assert.That(ws.Cell("F2").Value).IsEqualTo(68);
        await Assert.That(ws.Cell("F4").Value).IsEqualTo(4);
    }

    [Test]
    public async Task ImplicitIntersectionOperator_TakesTheTopLeftElementOfAnArray()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        wb.DefinedNames.Add("Grid", "{10,20;30,40}");
        ws.Cell("C3").FormulaA1 = "@Grid";

        await Assert.That(ws.Cell("C3").Value).IsEqualTo(10);
    }

    [Test]
    public async Task ImplicitIntersectionOperator_OfARangeTheFormulaDoesNotSpanIsAnError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D5").FormulaA1 = "@B1:B4";
        ws.Cell("D6").FormulaA1 = "@A1:C2";

        await Assert.That(ws.Cell("D5").Value).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Cell("D6").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task ImplicitIntersectionOperator_OfSeveralAreasIsAnError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("Both", "Sheet1!$A$1:$A$5,Sheet1!$B$1:$B$5");
        ws.Cell("D3").FormulaA1 = "@Both";

        await Assert.That(ws.Cell("D3").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    /// <summary>A single cell is its own intersection, so it needs no formula address.</summary>
    [Test]
    public async Task ImplicitIntersectionOperator_LeavesASingleCellAlone()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 5;

        await Assert.That(ws.Evaluate("@A1")).IsEqualTo(5);
    }

    [Test]
    public async Task ImplicitIntersectionOperator_TakesTheCellOfASpillOnTheFormulasRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");
        ws.Cell("B2").FormulaA1 = "@A1#";

        await Assert.That(ws.Cell("B2").Value).IsEqualTo(2);
    }

    #endregion

    #region Intersection operator (space)

    // A space between two references gives "a reference to cells common to the two references"
    // (Microsoft, "Calculation operators and precedence in Excel"). No cell in common is #NULL!.

    /// <summary>Fills A1:D8 with its row number times 10 plus its column number, so every cell differs.</summary>
    private static void SeedCells(IXLWorksheet ws)
    {
        for (var row = 1; row <= 8; row++)
        {
            for (var column = 1; column <= 4; column++)
                ws.Cell(row, column).Value = row * 10 + column;
        }
    }

    [Test]
    public async Task IntersectionOperator_GivesTheCellsCommonToBothReferences()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        SeedCells(ws);

        // Microsoft's example: B7:D7 and C6:C8 share C7.
        ws.Cell("F1").FormulaA1 = "SUM(B7:D7 C6:C8)";
        ws.Cell("F2").FormulaA1 = "B7:D7 C6:C8";
        // A1:C3 and B2:D4 share B2:C3: 22 + 23 + 32 + 33.
        ws.Cell("F3").FormulaA1 = "SUM(A1:C3 B2:D4)";

        await Assert.That(ws.Cell("F1").Value).IsEqualTo(73);
        await Assert.That(ws.Cell("F2").Value).IsEqualTo(73);
        await Assert.That(ws.Cell("F3").Value).IsEqualTo(110);
    }

    [Test]
    public async Task IntersectionOperator_WithNoCellInCommonIsANullError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("F1").FormulaA1 = "SUM(A1:A2 C1:C2)";

        await Assert.That(ws.Cell("F1").Value).IsEqualTo(XLError.NullValue);
    }

    [Test]
    public async Task IntersectionOperator_OfReferencesOnDifferentSheetsIsAnError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.AddWorksheet("Sheet2");
        ws.Cell("F1").FormulaA1 = "SUM(Sheet1!A1:B2 Sheet2!A1:B2)";

        await Assert.That(ws.Cell("F1").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    /// <summary>
    /// With several areas on a side, the cells common to both references are the union of what
    /// each left area has in common with each right area.
    /// </summary>
    [Test]
    public async Task IntersectionOperator_OfSeveralAreasKeepsEveryCommonCell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        SeedCells(ws);

        // Both sides cover A1:D2, split differently: all eight cells are common.
        ws.Cell("F1").FormulaA1 = "SUM((A1:B2,C1:D2) (A1:D1,A2:D2))";
        // A2 and C2 are common: 21 + 23.
        ws.Cell("F2").FormulaA1 = "SUM((A1:A3,C1:C3) A2:D2)";

        await Assert.That(ws.Cell("F1").Value).IsEqualTo(11 + 12 + 13 + 14 + 21 + 22 + 23 + 24);
        await Assert.That(ws.Cell("F2").Value).IsEqualTo(44);
    }

    /// <summary>
    /// <c>@</c> binds looser than the space but applies to the operand that follows it, so this is
    /// <c>B1:C3 ∩ @(C1:C9)</c>, which in row 2 is <c>B1:C3 ∩ C2</c>.
    /// </summary>
    [Test]
    public async Task IntersectionOperator_WorksWithTheImplicitIntersectionOperator()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        SeedCells(ws);
        ws.Cell("F2").FormulaA1 = "SUM(B1:C3 @C1:C9)";

        await Assert.That(ws.Cell("F2").Value).IsEqualTo(23);
    }

    [Test]
    public async Task IntersectionOperator_OfAnOperandThatIsNotAReferenceIsAnError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        wb.DefinedNames.Add("Grid", "{10,20;30,40}");
        ws.Cell("F1").FormulaA1 = "SUM(Grid A1:B2)";

        await Assert.That(ws.Cell("F1").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task IntersectionOperator_PassesAnErrorOperandThrough()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        wb.DefinedNames.Add("Broken", "#REF!");
        ws.Cell("F1").FormulaA1 = "SUM(Broken A1:B2)";

        await Assert.That(ws.Cell("F1").Value).IsEqualTo(XLError.CellReference);
    }

    [Test]
    public async Task IntersectionOperator_RecalculatesWhenACommonCellChanges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C7").Value = 1;
        ws.Cell("F1").FormulaA1 = "SUM(B7:D7 C6:C8)";
        await Assert.That(ws.Cell("F1").Value).IsEqualTo(1);

        ws.Cell("C7").Value = 5;

        await Assert.That(ws.Cell("F1").Value).IsEqualTo(5);
    }

    #endregion

    #region Cell-content scalar reduction (spec 37's AnyValue.TryReduceToScalar ladder)

    // A formula whose own top-level result is a reference — no function in between — is reduced to
    // the cell's stored value by the same ladder a function's scalar argument uses. These mirror the
    // "Implicit intersection" cases above one-for-one, but exercise the ladder through
    // XLCalcEngine.ToCellContentValue directly (a bare reference) rather than through a function
    // argument (ABS(reference), pre-reduced by the function-definition intersection pass). Spec 37
    // moves the ladder onto AnyValue.TryReduceToScalar and makes ToCellContentValue its first caller;
    // these pin the existing correct behaviour before that move, so a regression would show here.

    [Test]
    public async Task CellContent_SingleCellReferenceReadsTheCellUnchanged()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("D5").FormulaA1 = "B3:B3";

        await Assert.That(ws.Cell("D5").Value).IsEqualTo(-1);
    }

    [Test]
    public async Task CellContent_VerticalLineIntersectsAtTheFormulasOwnRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("D3").FormulaA1 = "B1:B10";

        await Assert.That(ws.Cell("D3").Value).IsEqualTo(-1);
    }

    [Test]
    public async Task CellContent_HorizontalLineIntersectsAtTheFormulasOwnColumn()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("B5").FormulaA1 = "A3:Z3";

        await Assert.That(ws.Cell("B5").Value).IsEqualTo(-1);
    }

    [Test]
    public async Task CellContent_BlankCellReadsAsZero()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D5").FormulaA1 = "B3:B3";

        // A blank cell's stored scalar is 0, same as reading it any other way (e.g. =B3).
        await Assert.That(ws.Cell("D5").Value).IsEqualTo(0);
    }

    [Test]
    public async Task CellContent_WithoutIntersectionResultsInValueError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("D5").FormulaA1 = "B1:B4";

        await Assert.That(ws.Cell("D5").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task CellContent_CanIntersectOnlyASingleArea()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = -1;
        ws.Cell("D3").FormulaA1 = "(B1:B2,B3:B5)"; // A continuous range made of two areas.

        await Assert.That(ws.Cell("D3").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task CellContent_ARectangleHasNoIntersectionToTake()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D3").FormulaA1 = "A1:B5";

        await Assert.That(ws.Cell("D3").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    #endregion

    #region Reference range operator

    [Test]
    [Arguments("A1:B2", 4)]
    [Arguments("A1:B5:C3", 3 * 5)]
    [Arguments("A1:C3:B5", 3 * 5)]
    [Arguments("A1:C3:B2", 3 * 3)]
    [Arguments("Sheet1!A1:B2", 4)]
    [Arguments("Sheet1!A1:Sheet1!B2", 4)]
    [Arguments("Sheet1!A1:Sheet1!B2", 4)]
    [Arguments("A1:Sheet1!B2", 4)]
    [Arguments("Sheet1!B2:C5:Sheet1!D3", 12)]
    [Arguments("(Sheet1!A1,A5):B5", 10)]
    [Arguments("B5:(Sheet1!A1,A5)", 10)]
    public async Task Range_UnifiesReferencesIntoSingleAreas(string referenceFormula, int expectedCellCount)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cells("A1:Z100").Value = 1;

        var referenceCells = ws.Evaluate($"SUM({referenceFormula})");
        await Assert.That(referenceCells).IsEqualTo(expectedCellCount);
    }

    [Test]
    [Arguments("Sheet1!A1:C5")]
    [Arguments("Sheet1!A1:B3:C5")]
    [Arguments("Sheet1!A1:B3:C4:Sheet1!B5:C5")]
    public async Task Range_LeftSideDeterminesSheetIfRightOmitted(string formula)
    {
        using var wb = new XLWorkbook();
        var firstSheet = wb.AddWorksheet("Sheet1");
        firstSheet.Cells("A1:C5").Value = 1;
        var secondSheet = wb.AddWorksheet("Sheet2");
        secondSheet.Cell("A1").FormulaA1 = $"=SUM({formula})";

        await Assert.That(secondSheet.Cell("A1").Value).IsEqualTo(15);
    }

    [Test]
    [Arguments("Current!A1:Other!B2")]
    [Arguments("A1:Other!B2")]
    [Arguments("A1:(Other!B2,C3)")]
    [Arguments("Other!A1:(Other!B2,C3)")] // C3 is taken from current worksheet since multiple areas on rhs
    [Arguments("(Other!A1,A5):Other!B2")] // A5 is taken from current worksheet since multiple areas on lhs
    [Arguments("(Current!A1):Other!B2")]
    // [TestCase("Other!A5:(B5)")] This causes #VALUE! in Excel, but it shouldn't. It's likely there is a "Fast parser for simple sheet areas" and "Full path" for complicated operands and they behave inconsistenly
    public async Task Range_UnificationAcrossSheetsResultsInValueError(string referenceFormula)
    {
        using var wb = new XLWorkbook();
        var formulaSheet = wb.AddWorksheet("Current");
        wb.AddWorksheet("Other");

        await Assert.That(formulaSheet.Evaluate($"SUM({referenceFormula})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("A1:IF(TRUE,1,)")]
    [Arguments("IF(TRUE,1,):A1")]
    [Arguments("IF(TRUE,\"text\"):A1")]
    [Arguments("IF(TRUE,FALSE):A1")]
    public async Task Range_OnlyReferencesCanBeRange(string referenceFormula)
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet();

        await Assert.That(sheet.Evaluate($"SUM({referenceFormula})")).IsEqualTo(XLError.IncompatibleValue);
    }

    #endregion

    #region Reference union

    [Test]
    [Arguments("A1,A2", 2)]
    [Arguments("A1:A3,B1", 4)]
    [Arguments("A1,B1:B3", 4)]
    [Arguments("Other!A1,Current!A1", 11)]
    [Arguments("A1,Other!A1", 11)]
    [Arguments("B2:D3,B2:D3", 12)] // Full overlap
    [Arguments("A1:B3,B1:C3", 12)] // Partial overlap
    [Arguments("Current!A1:B3,Other!B1:C3", 66)]
    [Arguments("A1,Other!A1,Current!A1", 10 + 1 + 1)]
    [Arguments("A1:B2,Other!A1:B2,B2:C3,Other!E5:Other!F6", 4 + 40 + 4 + 40)]
    public async Task Union_CanJoinAnyTwoRanges(string formula, int expectedSum)
    {
        using var wb = new XLWorkbook();
        var currentSheet = wb.AddWorksheet("Current");
        currentSheet.Cells("A1:F10").Value = 1;
        var otherSheet = wb.AddWorksheet("Other");
        otherSheet.Cells("A1:F10").Value = 10;

        // Not extra braces, so the comma is interpreted as union and not an extra argument
        var value = currentSheet.Evaluate($"SUM(({formula}))");

        await Assert.That(value).IsEqualTo(expectedSum);
    }

    #endregion
}
