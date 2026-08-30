using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The modern text functions: TEXTSPLIT, TEXTBEFORE, TEXTAFTER, VALUETOTEXT, ARRAYTOTEXT, UNICHAR,
/// UNICODE, DBCS and ENCODEURL. Expected values are the worked examples from Microsoft's
/// per-function documentation, or the documented behaviour applied to the arguments.
/// </summary>
[SetCulture("en-US")]
public class ModernTextTests
{
    private static IXLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return wb.AddWorksheet("Sheet1");
    }

    [Test]
    [Arguments("UNICHAR(65)", "A")]
    [Arguments("UNICHAR(66)", "B")]
    [Arguments("UNICHAR(9731)", "☃")]
    [Arguments("UNICHAR(128512)", "😀")] // Beyond the basic plane: a surrogate pair.
    public async Task UniChar_ReturnsTheCharacterForACodePoint(string formula, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("UNICHAR(0)")] // Zero is not a character.
    [Arguments("UNICHAR(-1)")]
    [Arguments("UNICHAR(1114112)")] // One past the last code point.
    public async Task UniChar_OutOfRangeReturnsIncompatibleValue(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("UNICODE(\"A\")", 65d)]
    [Arguments("UNICODE(\"Ant\")", 65d)] // Only the first character counts.
    [Arguments("UNICODE(\"☃\")", 9731d)]
    [Arguments("UNICODE(\"😀\")", 128512d)] // The pair is read as one code point.
    public async Task Unicode_ReturnsTheCodePointOfTheFirstCharacter(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task Unicode_OfEmptyTextReturnsIncompatibleValue()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("UNICODE(\"\")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task UniCharAndUnicode_RoundTrip()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").FormulaA1 = "UNICHAR(1000)";
            ws.Cell("A2").FormulaA1 = "UNICODE(A1)";

            await Assert.That((double)ws.Cell("A2").Value).IsEqualTo(1000d);
        }
    }

    [Test]
    // Microsoft's ENCODEURL example escapes the whole URL, colons and slashes included.
    [Arguments("ENCODEURL(\"http://contoso.sharepoint.com/teams/Finance/Documents/April Reports/Profit and Loss Statement.xlsx\")",
        "http%3A%2F%2Fcontoso.sharepoint.com%2Fteams%2FFinance%2FDocuments%2FApril%20Reports%2FProfit%20and%20Loss%20Statement.xlsx")]
    [Arguments("ENCODEURL(\"a b\")", "a%20b")]
    [Arguments("ENCODEURL(\"abcXYZ012-_.~\")", "abcXYZ012-_.~")] // The unreserved set passes through.
    [Arguments("ENCODEURL(\"a+b\")", "a%2Bb")]
    [Arguments("ENCODEURL(\"é\")", "%C3%A9")] // Escaped as its UTF-8 bytes.
    [Arguments("ENCODEURL(\"\")", "")]
    public async Task EncodeUrl_PercentEncodesEverythingOutsideTheUnreservedSet(string formula, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("DBCS(\"A\")", "Ａ")]
    [Arguments("DBCS(\"123\")", "１２３")]
    [Arguments("DBCS(\" \")", "　")]
    [Arguments("DBCS(\"ｱ\")", "ア")] // Half-width katakana.
    [Arguments("DBCS(\"ｶﾞ\")", "ガ")] // Base plus the voiced mark becomes one character.
    [Arguments("DBCS(\"ﾊﾟ\")", "パ")] // And the semi-voiced mark likewise.
    public async Task Dbcs_ConvertsHalfWidthToFullWidth(string formula, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("ＡＢＣ")]
    [Arguments("アイウエオ")]
    [Arguments("ガギグゲゴ")]
    [Arguments("パピプペポ")]
    public async Task Dbcs_IsTheInverseOfAsc(string fullWidth)
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = fullWidth;
            ws.Cell("A2").FormulaA1 = "DBCS(ASC(A1))";

            await Assert.That(ws.Cell("A2").Value).IsEqualTo(fullWidth);
        }
    }

    [Test]
    [Arguments("VALUETOTEXT(1234.01)", "1234.01")]
    [Arguments("VALUETOTEXT(\"abc\")", "abc")]
    [Arguments("VALUETOTEXT(TRUE)", "TRUE")]
    [Arguments("VALUETOTEXT(1234.01, 1)", "1234.01")] // Numbers read the same in either format.
    [Arguments("VALUETOTEXT(\"abc\", 1)", "\"abc\"")] // Strict form quotes text.
    [Arguments("VALUETOTEXT(1/0)", "#DIV/0!")] // Errors become their own text.
    [Arguments("VALUETOTEXT(1/0, 1)", "#DIV/0!")]
    public async Task ValueToText_RendersAnyValue(string formula, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task ValueToText_StrictFormDoublesInnerQuotes()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = "say \"hi\"";
            ws.Cell("A2").FormulaA1 = "VALUETOTEXT(A1, 1)";

            await Assert.That(ws.Cell("A2").Value).IsEqualTo("\"say \"\"hi\"\"\"");
        }
    }

    [Test]
    public async Task ArrayToText_JoinsWithCommasInConciseForm()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").Value = "b";
            ws.Cell("A2").Value = true;
            ws.Cell("B2").Value = 4;
            ws.Cell("D1").FormulaA1 = "ARRAYTOTEXT(A1:B2)";

            await Assert.That(ws.Cell("D1").Value).IsEqualTo("1, b, TRUE, 4");
        }
    }

    [Test]
    public async Task ArrayToText_ReproducesTheArrayLiteralInStrictForm()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").Value = "b";
            ws.Cell("A2").Value = true;
            ws.Cell("B2").Value = 4;
            ws.Cell("D1").FormulaA1 = "ARRAYTOTEXT(A1:B2, 1)";

            await Assert.That(ws.Cell("D1").Value).IsEqualTo("{1,\"b\";TRUE,4}");
        }
    }

    [Test]
    [Arguments("VALUETOTEXT(1, 2)")] // Only formats 0 and 1 exist.
    [Arguments("ARRAYTOTEXT(1, -1)")]
    public async Task ValueRendering_UnknownFormatReturnsIncompatibleValue(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("TEXTBEFORE(\"Red riding hood\", \" \")", "Red")]
    [Arguments("TEXTAFTER(\"Red riding hood\", \" \")", "riding hood")]
    [Arguments("TEXTBEFORE(\"Red riding hood\", \" \", 2)", "Red riding")]
    [Arguments("TEXTAFTER(\"Red riding hood\", \" \", 2)", "hood")]
    [Arguments("TEXTBEFORE(\"Red riding hood\", \" \", -1)", "Red riding")] // Counting from the end.
    [Arguments("TEXTAFTER(\"Red riding hood\", \" \", -2)", "riding hood")]
    public async Task TextBeforeAndAfter_CutAtTheNthDelimiter(string formula, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("TEXTBEFORE(\"Jones,Bob\", \",\")", "Jones")]
    [Arguments("TEXTAFTER(\"Jones,Bob\", \",\")", "Bob")]
    [Arguments("TEXTBEFORE(\"abc\", \"b\")", "a")]
    [Arguments("TEXTAFTER(\"abc\", \"b\")", "c")]
    [Arguments("TEXTBEFORE(\"abc\", \"a\")", "")] // Nothing before a leading delimiter.
    [Arguments("TEXTAFTER(\"abc\", \"c\")", "")]
    public async Task TextBeforeAndAfter_SplitOnASingleOccurrence(string formula, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task TextBefore_IsCaseSensitiveUnlessTold()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTBEFORE(\"aXbXc\", \"x\")")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTBEFORE(\"aXbXc\", \"x\", 1, 1)")).IsEqualTo("a");
    }

    [Test]
    public async Task TextBeforeAndAfter_CanTreatTheEndOfTheTextAsADelimiter()
    {
        // With match_end set, the very end counts as one more delimiter, so the last instance
        // returns the whole remaining text instead of not being found.
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTAFTER(\"a-b\", \"-\", 2, 0, 1)")).IsEqualTo("");
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTBEFORE(\"a-b\", \"-\", 2, 0, 1)")).IsEqualTo("a-b");
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTBEFORE(\"a-b\", \"-\", 2)")).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    public async Task TextBefore_AcceptsSeveralDelimitersAndPrefersTheLongestMatch()
    {
        // Splitting on both "<br>" and "<b>" must not leave a stray "r>".
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTBEFORE(\"a<br>b\", {\"<b>\",\"<br>\"})")).IsEqualTo("a");
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTAFTER(\"a<br>b\", {\"<b>\",\"<br>\"})")).IsEqualTo("b");
    }

    [Test]
    public async Task TextBefore_NotFoundReturnsTheFallbackWhenGiven()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTBEFORE(\"abc\", \"z\")")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTBEFORE(\"abc\", \"z\", 1, 0, 0, \"none\")")).IsEqualTo("none");
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTBEFORE(\"abc\", \"b\", 5, 0, 0, \"none\")")).IsEqualTo("none");
    }

    [Test]
    public async Task TextBefore_ZeroInstanceReturnsIncompatibleValue()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTBEFORE(\"abc\", \"b\", 0)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task TextSplit_SplitsIntoColumns()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("A1:C1").FormulaArrayA1 = "TEXTSPLIT(\"a,b,c\", \",\")";

            await Assert.That(ws.Cell("A1").Value).IsEqualTo("a");
            await Assert.That(ws.Cell("B1").Value).IsEqualTo("b");
            await Assert.That(ws.Cell("C1").Value).IsEqualTo("c");
        }
    }

    [Test]
    public async Task TextSplit_SplitsIntoRowsWhenOnlyARowDelimiterIsGiven()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("A1:A3").FormulaArrayA1 = "TEXTSPLIT(\"a;b;c\", , \";\")";

            await Assert.That(ws.Cell("A1").Value).IsEqualTo("a");
            await Assert.That(ws.Cell("A3").Value).IsEqualTo("c");
        }
    }

    [Test]
    public async Task TextSplit_ProducesATwoDimensionalGrid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("A1:B2").FormulaArrayA1 = "TEXTSPLIT(\"a,b;c,d\", \",\", \";\")";

            await Assert.That(ws.Cell("A1").Value).IsEqualTo("a");
            await Assert.That(ws.Cell("B1").Value).IsEqualTo("b");
            await Assert.That(ws.Cell("A2").Value).IsEqualTo("c");
            await Assert.That(ws.Cell("B2").Value).IsEqualTo("d");
        }
    }

    [Test]
    public async Task TextSplit_PadsShortRows()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("A1:C2").FormulaArrayA1 = "TEXTSPLIT(\"a,b,c;d\", \",\", \";\")";
            ws.Range("E1:G2").FormulaArrayA1 = "TEXTSPLIT(\"a,b,c;d\", \",\", \";\", FALSE, 0, \"-\")";

            await Assert.That(ws.Cell("B2").Value).IsEqualTo(XLError.NoValueAvailable);
            await Assert.That(ws.Cell("F2").Value).IsEqualTo("-");
        }
    }

    [Test]
    public async Task TextSplit_CanDropEmptyPieces()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("A1:C1").FormulaArrayA1 = "TEXTSPLIT(\"a,,b\", \",\")";
            ws.Range("E1:F1").FormulaArrayA1 = "TEXTSPLIT(\"a,,b\", \",\", , TRUE)";

            await Assert.That(ws.Cell("B1").Value).IsEqualTo(string.Empty);
            await Assert.That(ws.Cell("E1").Value).IsEqualTo("a");
            await Assert.That(ws.Cell("F1").Value).IsEqualTo("b");
        }
    }

    [Test]
    public async Task TextSplit_ReadsIgnoreEmptyFromACellReference()
    {
        // Spec 37 — was #VALUE! before the fix, because ignore_empty was a single-cell reference
        // rather than a literal.
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("H1").Value = true;
            ws.Range("E1:F1").FormulaArrayA1 = "TEXTSPLIT(\"a,,b\", \",\", , H1)";

            await Assert.That(ws.Cell("E1").Value).IsEqualTo("a");
            await Assert.That(ws.Cell("F1").Value).IsEqualTo("b");
        }
    }

    [Test]
    public async Task TextSplit_SpillsIntoTheGrid()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").SetDynamicFormulaA1("TEXTSPLIT(\"a,b;c,d\", \",\", \";\")");

            await Assert.That(ws.Cell("A1").Value).IsEqualTo("a");
            await Assert.That(ws.Cell("B1").Value).IsEqualTo("b");
            await Assert.That(ws.Cell("A2").Value).IsEqualTo("c");
            await Assert.That(ws.Cell("B2").Value).IsEqualTo("d");
        }
    }

    [Test]
    public async Task TextSplit_WithNoDelimiterAtAllReturnsIncompatibleValue()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("TEXTSPLIT(\"a,b\", )")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task TextSplit_IsCaseSensitiveUnlessTold()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Range("A1:A1").FormulaArrayA1 = "TEXTSPLIT(\"aXb\", \"x\")";
            ws.Range("C1:D1").FormulaArrayA1 = "TEXTSPLIT(\"aXb\", \"x\", , FALSE, 1)";

            await Assert.That(ws.Cell("A1").Value).IsEqualTo("aXb");
            await Assert.That(ws.Cell("C1").Value).IsEqualTo("a");
            await Assert.That(ws.Cell("D1").Value).IsEqualTo("b");
        }
    }

    [Test]
    public async Task ModernTextFunctions_ReadTheirArgumentsFromCells()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = "first.second.third";
            ws.Cell("A2").Value = ".";

            ws.Cell("B1").FormulaA1 = "TEXTBEFORE(A1, A2)";
            ws.Cell("B2").FormulaA1 = "TEXTAFTER(A1, A2, -1)";
            ws.Cell("B3").FormulaA1 = "ENCODEURL(A1)";

            await Assert.That(ws.Cell("B1").Value).IsEqualTo("first");
            await Assert.That(ws.Cell("B2").Value).IsEqualTo("third");
            await Assert.That(ws.Cell("B3").Value).IsEqualTo("first.second.third");
        }
    }
}
