using System;
using System.Globalization;
using System.Threading.Tasks;
using XLibur.Excel.CalcEngine;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// <c>ReferenceNode.Address</c> is the display text of a parsed reference. Evaluation never reads
/// it, so the node builds it on first read instead of for every reference it parses (#686).
/// </summary>
public class ReferenceNodeAddressTests
{
    private const int Parses = 1_000;

    /// <summary>
    /// Parsing <c>SUM(D1:H1)</c> measured 593 bytes with the text left unbuilt and 729 bytes with it
    /// built, in Debug and Release on net8.0 and net10.0 alike. The ceiling sits between the two.
    /// </summary>
    private const long CeilingBytesPerParse = 660;

    [Test]
    [Arguments("D1:H1", "D1:H1")]
    [Arguments("$A$1", "$A$1")]
    [Arguments("A:B", "A:B")]
    [Arguments("3:5", "3:5")]
    public async Task Address_IsTheA1TextOfTheReference(string formula, string expected)
    {
        var engine = new XLCalcEngine(CultureInfo.InvariantCulture);

        var node = (ReferenceNode)engine.Parse(formula).AstRoot;

        await Assert.That(node.Address).IsEqualTo(expected);
    }

    [Test]
    public async Task Address_IsTheR1C1TextOfAnR1C1Reference()
    {
        var engine = new XLCalcEngine(CultureInfo.InvariantCulture);

        await Assert.That(engine.TryParseR1C1("R1C4:R1C8", out var formula)).IsTrue();

        var node = (ReferenceNode)formula!.AstRoot;
        await Assert.That(node.Address).IsEqualTo("R1C4:R1C8");
    }

    [Test]
    public async Task ParsingAFormula_DoesNotBuildTheAddressText()
    {
        var engine = new XLCalcEngine(CultureInfo.InvariantCulture);
        _ = engine.Parse("SUM(D1:H1)");

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var before = GC.GetTotalAllocatedBytes(precise: true);
        for (var i = 0; i < Parses; i++)
            engine.Parse("SUM(D1:H1)");
        var perParse = (GC.GetTotalAllocatedBytes(precise: true) - before) / Parses;
        await Assert.That(perParse).IsLessThan(CeilingBytesPerParse);
    }
}
