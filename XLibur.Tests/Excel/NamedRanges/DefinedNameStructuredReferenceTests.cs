using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.NamedRanges;

/// <summary>
/// What <see cref="IXLDefinedName.Ranges"/> answers for a name whose formula is a structured
/// reference (<c>=Sales[Amount]</c>).
/// </summary>
/// <remarks>
/// <para>
/// The area is resolved when it is asked for, through the same
/// <c>StructuredReferenceResolver</c> that evaluation and the dependency tree use, so a name
/// honours the <c>#Headers</c> / <c>#Totals</c> / <c>#All</c> specifiers and a column span rather
/// than always answering with the first column's data.
/// </para>
/// <para>
/// The two-specifier form Excel writes as <c>Sales[[#Headers],[#Data]]</c> is not covered: it
/// throws inside ClosedXML.Parser while the formula is being parsed, before any of this is
/// reached, so it is not something resolution can answer for either way. The resolver does handle
/// the combination once a formula carrying it parses.
/// </para>
/// </remarks>
public class DefinedNameStructuredReferenceTests
{
    /// <summary>
    /// A table at <c>A1:C4</c> — headers <c>A1:C1</c>, data <c>A2:C3</c>, totals <c>A4:C4</c>.
    /// </summary>
    private static XLWorkbook TableBook()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Product";
        ws.Cell("B1").Value = "Amount";
        ws.Cell("C1").Value = "Tax";
        ws.Cell("A2").Value = "Widget";
        ws.Cell("B2").Value = 10;
        ws.Cell("C2").Value = 1;
        ws.Cell("A3").Value = "Gadget";
        ws.Cell("B3").Value = 20;
        ws.Cell("C3").Value = 2;
        ws.Range("A1:C3").CreateTable("Sales").SetShowTotalsRow(true);
        return wb;
    }

    private static string[] RangesOf(XLWorkbook wb, string formula) =>
        wb.DefinedNames.Add("Probe", formula).Ranges
            .Select(r => r.RangeAddress.ToString()!)
            .ToArray();

    [Test]
    [Arguments("Sales[Amount]", "B2:B3")]
    [Arguments("Sales[[#Data],[Amount]]", "B2:B3")]
    [Arguments("Sales[[#Headers],[Amount]]", "B1:B1")]
    [Arguments("Sales[[#Totals],[Amount]]", "B4:B4")]
    [Arguments("Sales[[#All],[Amount]]", "B1:B4")]
    public async Task AnAreaSpecifierSelectsThatPartOfTheColumn(string formula, string expected)
    {
        using var wb = TableBook();

        await Assert.That(RangesOf(wb, formula)).IsEquivalentTo(new[] { expected });
    }

    /// <summary>A column span covers every column between its ends, not just the first.</summary>
    [Test]
    [Arguments("Sales[[Amount]:[Tax]]", "B2:C3")]
    [Arguments("Sales[[Product]:[Tax]]", "A2:C3")]
    [Arguments("Sales[[#Headers],[Amount]:[Tax]]", "B1:C1")]
    public async Task AColumnSpanCoversItsWholeWidth(string formula, string expected)
    {
        using var wb = TableBook();

        await Assert.That(RangesOf(wb, formula)).IsEquivalentTo(new[] { expected });
    }

    /// <summary>
    /// A reference naming no column covers the table's full width. These used to resolve to
    /// nothing at all.
    /// </summary>
    [Test]
    [Arguments("Sales[#All]", "A1:C4")]
    [Arguments("Sales[#Data]", "A2:C3")]
    [Arguments("Sales[#Headers]", "A1:C1")]
    [Arguments("Sales[#Totals]", "A4:C4")]
    public async Task AWholeTableReferenceCoversTheTable(string formula, string expected)
    {
        using var wb = TableBook();

        await Assert.That(RangesOf(wb, formula)).IsEquivalentTo(new[] { expected });
    }

    /// <summary>
    /// A name Excel would show as <c>#REF!</c> yields no range rather than throwing. Reachable
    /// from an ordinary load — a workbook whose table or column has since been renamed — so
    /// reading the property must not be able to fail.
    /// </summary>
    [Test]
    [Arguments("Sales[NoSuchColumn]")]
    [Arguments("NoSuchTable[Amount]")]
    [Arguments("Sales[[NoSuchColumn]:[Tax]]")]
    public async Task AReferenceThatCannotResolveYieldsNoRange(string formula)
    {
        using var wb = TableBook();

        await Assert.That(RangesOf(wb, formula)).IsEmpty();
    }

    /// <summary>
    /// A defined name has no anchoring cell, so <c>[@Column]</c> — which means "this row" — has no
    /// row to mean. It resolves to nothing rather than guessing one.
    /// </summary>
    [Test]
    public async Task AThisRowReferenceYieldsNoRange()
    {
        using var wb = TableBook();

        await Assert.That(RangesOf(wb, "Sales[@Amount]")).IsEmpty();
    }

    /// <summary>Several references in one formula each contribute their own range.</summary>
    [Test]
    public async Task EveryReferenceInTheFormulaContributes()
    {
        using var wb = TableBook();

        await Assert.That(RangesOf(wb, "SUM(Sales[Amount], Sales[Tax])"))
            .IsEquivalentTo(new[] { "B2:B3", "C2:C3" });
    }

    /// <summary>
    /// Structured references and ordinary sheet references are gathered by separate passes, so a
    /// formula holding both has to come back with both.
    /// </summary>
    /// <remarks>
    /// A sheet reference keeps the fixed markers it was written with, while a structured reference
    /// is resolved to the area the table currently covers and so has none.
    /// </remarks>
    [Test]
    public async Task AStructuredReferenceAndASheetReferenceBothContribute()
    {
        using var wb = TableBook();

        await Assert.That(RangesOf(wb, "SUM(Sales[Amount], Data!$A$10:$A$11)"))
            .IsEquivalentTo(new[] { "B2:B3", "$A$10:$A$11" });
    }

    /// <summary>
    /// A structured reference that cannot resolve does not make the name invalid. Deliberate:
    /// <see cref="IXLDefinedName.IsValid"/> reports a <c>#REF!</c> written into the formula, and it
    /// is what <c>ValidNamedRanges()</c> filters on — which the IO writers and XLibur.Report's range
    /// binder both consume. A name that survives a load and save today has to keep surviving one.
    /// </summary>
    [Test]
    [Arguments("NoSuchTable[Amount]")]
    [Arguments("Sales[NoSuchColumn]")]
    public async Task AnUnresolvableReferenceLeavesTheNameValid(string formula)
    {
        using var wb = TableBook();
        var name = wb.DefinedNames.Add("Probe", formula);

        await Assert.That(name.Ranges).IsEmpty();
        await Assert.That(name.IsValid).IsTrue();
        await Assert.That(wb.DefinedNames.ValidNamedRanges().Any(n => n.Name == "Probe")).IsTrue();
    }

    /// <summary>The resolved area follows the table, rather than being fixed when the name was added.</summary>
    [Test]
    public async Task TheAreaFollowsTheTableWhenItGrows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Amount";
        ws.Cell("A2").Value = 10;
        ws.Cell("A3").Value = 20;
        var table = ws.Range("A1:A3").CreateTable("Sales");

        var name = wb.DefinedNames.Add("Probe", "Sales[Amount]");
        await Assert.That(name.Ranges.Single().RangeAddress.ToString()).IsEqualTo("A2:A3");

        ws.Cell("A4").Value = 30;
        table.Resize(ws.Range("A1:A4"));

        await Assert.That(name.Ranges.Single().RangeAddress.ToString()).IsEqualTo("A2:A4");
    }

    /// <summary>A table with no totals row has no totals to point at.</summary>
    [Test]
    public async Task TotalsOfATableWithoutATotalsRowYieldNoRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Amount";
        ws.Cell("A2").Value = 10;
        ws.Range("A1:A2").CreateTable("Sales");

        await Assert.That(RangesOf(wb, "Sales[#Totals]")).IsEmpty();
    }
}
