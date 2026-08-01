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
