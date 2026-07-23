using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

// IFS and SWITCH — scalar logical selectors added alongside IF.
[TestFixture]
public class IfsSwitchTests
{
    private static XLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return (XLWorksheet)wb.AddWorksheet("Sheet1");
    }

    [Test]
    public void Ifs_ReturnsValueOfFirstTrueCondition()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 2;
            Assert.AreEqual("two", ws.Evaluate(@"IFS(A1=1, ""one"", A1=2, ""two"", A1=3, ""three"")"));
            // First TRUE wins even if a later condition also matches.
            Assert.AreEqual("first", ws.Evaluate(@"IFS(TRUE, ""first"", TRUE, ""second"")"));
            // A numeric (non-zero) condition is truthy.
            Assert.AreEqual("y", ws.Evaluate(@"IFS(0, ""x"", 5, ""y"")"));
        }
    }

    [Test]
    public void Ifs_NoTrueCondition_ReturnsNotAvailable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"IFS(FALSE, 1, FALSE, 2)"));
            // Odd trailing argument with no earlier match -> #N/A.
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"IFS(FALSE, 1, 2)"));
        }
    }

    [Test]
    public void Ifs_OddTrailingArgument_IgnoredWhenEarlierConditionMatches()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(1, ws.Evaluate(@"IFS(TRUE, 1, 2)"));
        }
    }

    [Test]
    public void Ifs_ErrorConditionIsPropagated()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate(@"IFS(1/0, ""x"", TRUE, ""y"")"));
        }
    }

    [Test]
    public void Switch_ReturnsResultOfFirstMatch()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual("b", ws.Evaluate(@"SWITCH(2, 1, ""a"", 2, ""b"", 3, ""c"")"));
            // First match wins.
            Assert.AreEqual("a", ws.Evaluate(@"SWITCH(1, 1, ""a"", 1, ""b"")"));
        }
    }

    [Test]
    public void Switch_TextMatchIsCaseInsensitive()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual("match", ws.Evaluate(@"SWITCH(""red"", ""RED"", ""match"", ""no"")"));
        }
    }

    [Test]
    public void Switch_NoMatch_UsesDefaultWhenPresent()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual("none", ws.Evaluate(@"SWITCH(9, 1, ""a"", 2, ""b"", ""none"")"));
        }
    }

    [Test]
    public void Switch_NoMatch_NoDefault_ReturnsNotAvailable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.NoValueAvailable, ws.Evaluate(@"SWITCH(9, 1, ""a"", 2, ""b"")"));
        }
    }

    [Test]
    public void Switch_ErrorExpressionIsPropagated()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate(@"SWITCH(1/0, 1, ""a"", ""default"")"));
        }
    }
}
