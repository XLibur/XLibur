using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tests.Tags;

/// <summary>
/// Row order and group labels follow the culture the <em>engine</em> was given, never the machine's.
/// </summary>
/// <remarks>
/// The engines default to invariant so that a report does not change shape with the locale it happens
/// to be generated on, and take a culture so that a localised report can be ordered and formatted the
/// way its readers expect. Both promises are broken if the comparer or the label reaches for
/// <see cref="CultureInfo.CurrentCulture"/> instead, which is
/// <see href="https://github.com/XLibur/XLibur/issues/275">#275</see>. These tests move the machine's
/// culture and the engine's independently, so a regression to either one shows up as a different
/// answer rather than as nothing at all.
/// <para>
/// The machine culture is set inside the test rather than by attribute because the suite resets it
/// before every test, so nothing after is affected. Collection assertions pass
/// <see cref="CollectionOrdering.Matching"/> because the order <em>is</em> the assertion — two
/// collations of the same three products differ in nothing else.
/// </para>
/// </remarks>
public class CultureBoundOrderingTests
{
    /// <summary>Sorts <c>Ä</c> after <c>Z</c>, where the invariant culture sorts it as an <c>A</c>.</summary>
    private static readonly CultureInfo Swedish = CultureInfo.GetCultureInfo("sv-SE");

    /// <summary>Writes a date as <c>30.07.2026</c> and a decimal with a comma.</summary>
    private static readonly CultureInfo Czech = CultureInfo.GetCultureInfo("cs-CZ");

    /// <summary>
    /// Three products the two collations disagree about: invariant gives Äpple, Banan, Zebra;
    /// Swedish gives Banan, Zebra, Äpple.
    /// </summary>
    private static List<SaleItem> Accented() => new()
    {
        new() { Product = "Zebra", Quantity = 1 },
        new() { Product = "Äpple", Quantity = 2 },
        new() { Product = "Banan", Quantity = 3 },
    };

    private static List<SaleItem> Dated() => new()
    {
        new() { Product = "Widget", Quantity = 2, SoldOn = new DateTime(2026, 7, 30) },
        new() { Product = "Gadget", Quantity = 5, SoldOn = new DateTime(2026, 12, 3) },
    };

    /// <summary>
    /// A two-column range over A3:B4 — the first column holding <paramref name="expression"/>, the
    /// second a quantity — with row 4 the options row.
    /// </summary>
    private static XLWorkbook Template(string expression, string optionsA, string optionsB = "")
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        sheet.Cell("A2").Value = "Key";
        sheet.Cell("B2").Value = "Quantity";
        sheet.Cell("A3").Value = expression;
        sheet.Cell("B3").Value = "{{ item.Quantity }}";
        sheet.Cell("A4").Value = optionsA;

        if (optionsB.Length > 0)
        {
            sheet.Cell("B4").Value = optionsB;
        }

        workbook.DefinedNames.Add("Items", sheet.Range("A3:B4"));
        return workbook;
    }

    private static void Generate(IXLWorkbook workbook, object items, IExpressionEngine? engine = null)
    {
        using var template = new XLTemplate(workbook, engine);
        template.AddVariable("Items", items);
        template.Generate();
    }

    private static List<string> ColumnText(IXLWorksheet sheet, string column, int firstRow, int count) =>
        Enumerable.Range(firstRow, count).Select(r => sheet.Cell(column + r).Value.ToString() ?? string.Empty).ToList();

    /// <summary>The subtotal rows, which with one item per group land on 4, 6 and 8.</summary>
    private static List<string> GroupLabels(IXLWorksheet sheet, params int[] rows) =>
        rows.Select(row => sheet.Cell("A" + row).Value.GetText()).ToList();

    [Test]
    public async Task SortIgnoresTheMachineCulture()
    {
        using var workbook = Template("{{ item.Product }}", optionsA: "<<Sort>>");
        TestDefaults.SetCulture(Swedish);

        // No engine given, so the default engine's invariant culture decides and the machine's
        // Swedish collation gets no say.
        Generate(workbook, Accented());

        await Assert.That(ColumnText(workbook.Worksheet("Report"), "A", 3, 3))
            .IsEquivalentTo(new[] { "Äpple", "Banan", "Zebra" }, CollectionOrdering.Matching);
    }

    [Test]
    public async Task SortFollowsTheEngineCulture()
    {
        using var workbook = Template("{{ item.Product }}", optionsA: "<<Sort>>");

        Generate(workbook, Accented(), new ScribanExpressionEngine(Swedish));

        await Assert.That(ColumnText(workbook.Worksheet("Report"), "A", 3, 3))
            .IsEquivalentTo(new[] { "Banan", "Zebra", "Äpple" }, CollectionOrdering.Matching);
    }

    /// <summary>
    /// Two keys of different types are compared as text, which is the comparer's other arm and was
    /// the other half of the same bug. A string against a bool reaches it; invariant collation puts
    /// <c>Äpple</c> before <c>True</c>, Swedish does not.
    /// </summary>
    [Test]
    public async Task SortOfMixedTypesIgnoresTheMachineCulture()
    {
        using var workbook = Template("{{ item }}", optionsA: "<<Sort>>");
        TestDefaults.SetCulture(Swedish);

        Generate(workbook, new List<object?> { true, "Äpple" });

        await Assert.That(workbook.Worksheet("Report").Cell("A3").Value.IsText).IsTrue();
    }

    [Test]
    public async Task SortOfMixedTypesFollowsTheEngineCulture()
    {
        using var workbook = Template("{{ item }}", optionsA: "<<Sort>>");

        Generate(workbook, new List<object?> { true, "Äpple" }, new ScribanExpressionEngine(Swedish));

        await Assert.That(workbook.Worksheet("Report").Cell("A3").Value.IsBoolean).IsTrue();
    }

    [Test]
    public async Task GroupOrderIgnoresTheMachineCulture()
    {
        using var workbook = Template("{{ item.Product }}", optionsA: "<<Group>>", optionsB: "<<Sum>>");
        TestDefaults.SetCulture(Swedish);

        Generate(workbook, Accented());

        await Assert.That(GroupLabels(workbook.Worksheet("Report"), 4, 6, 8))
            .IsEquivalentTo(new[] { "Äpple Total", "Banan Total", "Zebra Total" }, CollectionOrdering.Matching);
    }

    [Test]
    public async Task GroupOrderFollowsTheEngineCulture()
    {
        using var workbook = Template("{{ item.Product }}", optionsA: "<<Group>>", optionsB: "<<Sum>>");

        Generate(workbook, Accented(), new ScribanExpressionEngine(Swedish));

        await Assert.That(GroupLabels(workbook.Worksheet("Report"), 4, 6, 8))
            .IsEquivalentTo(new[] { "Banan Total", "Zebra Total", "Äpple Total" }, CollectionOrdering.Matching);
    }

    [Test]
    public async Task AGroupLabelOverADateIgnoresTheMachineCulture()
    {
        using var workbook = Template("{{ item.SoldOn }}", optionsA: "<<Group>>", optionsB: "<<Sum>>");
        TestDefaults.SetCulture(Czech);

        Generate(workbook, Dated());

        var expected = new DateTime(2026, 7, 30).ToString(null, CultureInfo.InvariantCulture);
        await Assert.That(workbook.Worksheet("Report").Cell("A4").Value.GetText()).IsEqualTo(expected + " Total");
    }

    [Test]
    public async Task AGroupLabelOverADateFollowsTheEngineCulture()
    {
        using var workbook = Template("{{ item.SoldOn }}", optionsA: "<<Group>>", optionsB: "<<Sum>>");

        Generate(workbook, Dated(), new ScribanExpressionEngine(Czech));

        var expected = new DateTime(2026, 7, 30).ToString(null, Czech);
        await Assert.That(workbook.Worksheet("Report").Cell("A4").Value.GetText()).IsEqualTo(expected + " Total");
    }

    [Test]
    public async Task AGroupLabelOverADecimalFollowsTheEngineCulture()
    {
        using var workbook = Template("{{ item.UnitPrice }}", optionsA: "<<Group>>", optionsB: "<<Sum>>");

        Generate(
            workbook,
            new List<SaleItem> { new() { Product = "Widget", Quantity = 2, UnitPrice = 1.5m } },
            new ScribanExpressionEngine(Czech));

        // A decimal comma, because that is how the report this engine was given a culture for reads.
        await Assert.That(workbook.Worksheet("Report").Cell("A4").Value.GetText()).IsEqualTo("1,5 Total");
    }
}
