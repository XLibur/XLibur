using System;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Misc;

public class FormulaTests
{
    [Test]
    public async Task CopyFormula()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").FormulaA1 = "B1";
        ws.Cell("A1").CopyTo("A2");
        await Assert.That(ws.Cell("A2").FormulaA1).IsEqualTo("B2");
    }

    [Test]
    public async Task CopyFormula2()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        ws.Cell("A1").FormulaA1 = "A2-1";
        ws.Cell("A1").CopyTo("B1");
        await Assert.That(ws.Cell("A1").FormulaR1C1).IsEqualTo("R[1]C-1");
        await Assert.That(ws.Cell("B1").FormulaR1C1).IsEqualTo("R[1]C-1");
        await Assert.That(ws.Cell("B1").FormulaA1).IsEqualTo("B2-1");

        ws.Cell("A1").FormulaA1 = "B1+1";
        ws.Cell("A1").CopyTo("A2");
        await Assert.That(ws.Cell("A1").FormulaR1C1).IsEqualTo("RC[1]+1");
        await Assert.That(ws.Cell("A2").FormulaR1C1).IsEqualTo("RC[1]+1");
        await Assert.That(ws.Cell("A2").FormulaA1).IsEqualTo("B2+1");
    }

    [Test]
    public async Task CopyFormulaWithSheetNameThatResemblesFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("S10 Data");
        ws.Cell("A1").Value = "Some value";
        ws.Cell("A2").Value = 123;

        ws = wb.Worksheets.Add("Summary");
        ws.Cell("A1").FormulaA1 = "='S10 Data'!A1";
        await Assert.That(ws.Cell("A1").Value).IsEqualTo("Some value");

        ws.Cell("A1").CopyTo("A2");
        await Assert.That(ws.Cell("A2").FormulaA1).IsEqualTo("'S10 Data'!A2");

        ws.Cell("A1").CopyTo("B1");
        await Assert.That(ws.Cell("B1").FormulaA1).IsEqualTo("'S10 Data'!B1");

        ws.Cell("A3").FormulaA1 = "=SUM('S10 Data'!A2)";
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(123);
    }

    [Test]
    public async Task FormulaWithReferenceIncludingSheetName()
    {
        using var wb = new XLWorkbook();
        object value;
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
        ws.Cell("B1").FormulaA1 = "=SUM(A1:A50)";
        value = ws.Cell("B1").Value;
        await Assert.That(value).IsEqualTo(ExpectedCellValue.From(1275));

        ws = wb.AddWorksheet("Sheet2");

        ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet1!A50)";
        value = ws.Cell("A1").Value;
        await Assert.That(value).IsEqualTo(ExpectedCellValue.From(1275));

        ws.Cell("B1").FormulaA1 = "=SUM(Sheet1!A1:A50)";
        value = ws.Cell("B1").Value;
        await Assert.That(value).IsEqualTo(ExpectedCellValue.From(1275));
    }

    [Test]
    public async Task InvalidReferences()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
        ws = wb.AddWorksheet("Sheet2");

        ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet2!A50)";
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.IncompatibleValue);

        ws.Cell("B1").FormulaA1 = "=SUM(UnknownSheet!A50)";
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(XLError.CellReference);
    }

    [Test]
    public async Task DateAgainstStringComparison()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = new DateTime(2016, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

        ws.Cell("A2").FormulaA1 = @"=IF(A1 = """", ""A"", ""B"")";
        var actual = ws.Cell("A2").Value;
        await Assert.That("B").IsEqualTo(actual);

        ws.Cell("A3").FormulaA1 = @"=IF("""" = A1, ""A"", ""B"")";
        actual = ws.Cell("A3").Value;
        await Assert.That("B").IsEqualTo(actual);
    }

    [Test]
    public async Task FormulaThatReferencesEntireRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().Value = 1;
        ws.FirstCell().CellRight().Value = 2;
        ws.FirstCell().CellRight(5).Value = 3;

        ws.FirstCell().CellBelow().FormulaA1 = "=SUM(1:1)";

        var actual = ws.FirstCell().CellBelow().Value;
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(6));
    }

    [Test]
    public async Task FormulaThatReferencesEntireColumn()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().Value = 1;
        ws.FirstCell().CellBelow().Value = 2;
        ws.FirstCell().CellBelow(5).Value = 3;

        ws.FirstCell().CellRight().FormulaA1 = "=SUM(A:A)";

        var actual = ws.FirstCell().CellRight().Value;
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(6));
    }

    [Test]
    public async Task FormulaThatStartsWithEqualsAndPlus()
    {
        object actual = XLWorkbook.EvaluateExpr("=MID(\"This is a test\", 6, 2)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("is"));

        actual = XLWorkbook.EvaluateExpr("=+MID(\"This is a test\", 6, 2)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("is"));

        actual = XLWorkbook.EvaluateExpr("=+++++MID(\"This is a test\", 6, 2)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("is"));

        actual = XLWorkbook.EvaluateExpr("+MID(\"This is a test\", 6, 2)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("is"));
    }

    [Test]
    public async Task UnimplementedStandardFunctionsAreEvaluatedToNameNotFoundError()
    {
        // RTD will never be implemented
        var actual =
            XLWorkbook.EvaluateExpr("RTD(\"MyRTDServerProdID\",\"MyServer\",\"RaceNum\",\"RunnerID\",\"StatType\")");
        await Assert.That(actual).IsEqualTo(XLError.NameNotRecognized);
    }

    [Test]
    public async Task FormulasWithErrors()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("YEAR(#REF!)")).IsEqualTo(XLError.CellReference);
        await Assert.That(XLWorkbook.EvaluateExpr("YEAR(#VALUE!)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("YEAR(#DIV/0!)")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(XLWorkbook.EvaluateExpr("YEAR(#NAME?)")).IsEqualTo(XLError.NameNotRecognized);
        await Assert.That(XLWorkbook.EvaluateExpr("YEAR(#N/A)")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(XLWorkbook.EvaluateExpr("YEAR(#NULL!)")).IsEqualTo(XLError.NullValue);
        await Assert.That(XLWorkbook.EvaluateExpr("YEAR(#NUM!)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task LegacyFunctionPropagateErrorWithoutException()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("SIN(YEAR(#NAME?))+1")).IsEqualTo(XLError.NameNotRecognized);
    }

    [Test]
    public async Task UnicodeLetterParsing()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet C CÄ");
        var ws2 = wb.AddWorksheet("ÖC");
        var ws3 = wb.AddWorksheet("Sheet3");

        ws1.FirstCell().SetValue(100);
        ws2.FirstCell().SetValue(50);

        ws3.FirstCell().FormulaA1 = "='Sheet C CÄ'!A1";
        ws3.FirstCell().CellBelow().FormulaA1 = "ÖC!A1";

        await Assert.That(ws3.FirstCell().Value).IsEqualTo(100);
        await Assert.That(ws3.FirstCell().CellBelow().Value).IsEqualTo(50);
    }

    [Test]
    public async Task ShiftFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").FormulaA1 = "ATAN2(C1,C2)";
        ws.Cell("B2").FormulaA1 = "DEC2HEX(C2)";
        ws.Range("B3:B5").FormulaArrayA1 = "DAYS360(C3:C5, D3:D5)";

        ws.Column(1).Delete();

        await Assert.That(ws.Cell("A1").FormulaA1).IsEqualTo("ATAN2(B1,B2)");
        await Assert.That(ws.Cell("A2").FormulaA1).IsEqualTo("DEC2HEX(B2)");
        await Assert.That(ws.Cell("A3").HasArrayFormula).IsTrue();
        await Assert.That(ws.Cell("A3").FormulaA1).IsEqualTo("DAYS360(B3:B5, C3:C5)");
    }
}
