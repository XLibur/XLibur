using System;
using System.Globalization;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// A parsed function call keeps its name. The parser hands the name over as a span of the formula
/// text, and turning it into a new string cost 32 bytes per function per parse. On .NET 9 and later
/// the name of a registered function is the registry's own string (#686).
/// </summary>
public class FunctionNameAllocationTests
{
    private const int Parses = 1_000;

#if NET9_0_OR_GREATER
    /// <summary>
    /// Parsing <c>SUM(D1:H1)</c> measured 201 bytes with the registry's name and 233 with a copy,
    /// in Debug and Release. The ceiling sits between the two.
    /// </summary>
    private const long CeilingBytesPerParse = 217;
#else
    /// <summary>
    /// .NET 8 has no lookup of a string key by a span, so the name is still copied: 233 bytes.
    /// </summary>
    private const long CeilingBytesPerParse = 250;
#endif

    [Test]
    public async Task ParsingAFunctionCall_DoesNotCopyTheName()
    {
        var engine = new XLCalcEngine(CultureInfo.InvariantCulture);
        _ = engine.Parse("SUM(D1:H1)");

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var before = GC.GetAllocatedBytesForCurrentThread();
        for (var i = 0; i < Parses; i++)
            engine.Parse("SUM(D1:H1)");
        var perParse = (GC.GetAllocatedBytesForCurrentThread() - before) / Parses;
        await Assert.That(perParse).IsLessThan(CeilingBytesPerParse);
    }

    /// <summary>
    /// The node may keep the registry's casing instead of the text's. Every reader of the name
    /// ignores case, so a function written in any case, or with the prefix Excel stores a newer
    /// function under, is still found, and a name that is not registered is still <c>#NAME?</c>.
    /// </summary>
    [Test]
    [Arguments("SUM(1,2)", 3d)]
    [Arguments("sum(1,2)", 3d)]
    [Arguments("Sum(1,2)", 3d)]
    [Arguments("_xlfn.CONCAT(\"a\",\"b\")", "ab")]
    [Arguments("concat(\"a\",\"b\")", "ab")]
    [Arguments("NOSUCHFUNCTION(1)", "#NAME?")]
    public async Task AFunctionIsFoundWhateverTheCaseOfItsName(string formula, object expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var actual = ws.Evaluate(formula);

        var expectedValue = expected is "#NAME?" ? XLError.NameNotRecognized : XLCellValue.FromObject(expected);
        await Assert.That(actual).IsEqualTo(expectedValue);
    }

    [Test]
    public async Task TheNodeKeepsTheNameOfTheFunction()
    {
        var engine = new XLCalcEngine(CultureInfo.InvariantCulture);

        var node = (FunctionNode)engine.Parse("sum(D1:H1)").AstRoot;

        await Assert.That(node.Name).IsEqualTo("SUM").IgnoringCase();
    }
}
