using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;
// IFS and SWITCH — scalar logical selectors added alongside IF.
public class IfsSwitchTests
{
    private static XLWorksheet NewSheet(out XLWorkbook wb)
    {
        wb = new XLWorkbook();
        return (XLWorksheet)wb.AddWorksheet("Sheet1");
    }

    [Test]
    public async Task Ifs_ReturnsValueOfFirstTrueCondition()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            ws.Cell("A1").Value = 2;
            await Assert.That(ws.Evaluate(@"IFS(A1=1, ""one"", A1=2, ""two"", A1=3, ""three"")")).IsEqualTo("two");
            // First TRUE wins even if a later condition also matches.
            await Assert.That(ws.Evaluate(@"IFS(TRUE, ""first"", TRUE, ""second"")")).IsEqualTo("first");
            // A numeric (non-zero) condition is truthy.
            await Assert.That(ws.Evaluate(@"IFS(0, ""x"", 5, ""y"")")).IsEqualTo("y");
        }
    }

    [Test]
    public async Task Ifs_NoTrueCondition_ReturnsNotAvailable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate(@"IFS(FALSE, 1, FALSE, 2)")).IsEqualTo(XLError.NoValueAvailable);
            // Odd trailing argument with no earlier match -> #N/A.
            await Assert.That(ws.Evaluate(@"IFS(FALSE, 1, 2)")).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task Ifs_OddTrailingArgument_IgnoredWhenEarlierConditionMatches()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate(@"IFS(TRUE, 1, 2)")).IsEqualTo(1);
        }
    }

    [Test]
    public async Task Ifs_ErrorConditionIsPropagated()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate(@"IFS(1/0, ""x"", TRUE, ""y"")")).IsEqualTo(XLError.DivisionByZero);
        }
    }

    [Test]
    public async Task Switch_ReturnsResultOfFirstMatch()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate(@"SWITCH(2, 1, ""a"", 2, ""b"", 3, ""c"")")).IsEqualTo("b");
            // First match wins.
            await Assert.That(ws.Evaluate(@"SWITCH(1, 1, ""a"", 1, ""b"")")).IsEqualTo("a");
        }
    }

    [Test]
    public async Task Switch_TextMatchIsCaseInsensitive()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate(@"SWITCH(""red"", ""RED"", ""match"", ""no"")")).IsEqualTo("match");
        }
    }

    [Test]
    public async Task Switch_NoMatch_UsesDefaultWhenPresent()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate(@"SWITCH(9, 1, ""a"", 2, ""b"", ""none"")")).IsEqualTo("none");
        }
    }

    [Test]
    public async Task Switch_NoMatch_NoDefault_ReturnsNotAvailable()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate(@"SWITCH(9, 1, ""a"", 2, ""b"")")).IsEqualTo(XLError.NoValueAvailable);
        }
    }

    [Test]
    public async Task Switch_ErrorExpressionIsPropagated()
    {
        var ws = NewSheet(out var wb);
        using (wb)
        {
            await Assert.That(ws.Evaluate(@"SWITCH(1/0, 1, ""a"", ""default"")")).IsEqualTo(XLError.DivisionByZero);
        }
    }
}
