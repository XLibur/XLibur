using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Report.Ranges;

namespace XLibur.Report.Tests.Ranges;

/// <summary>
/// Which defined names bind when more than one of them shares a name.
/// </summary>
/// <remarks>
/// Excel scopes a defined name either to the workbook or to one sheet, and the same name may be
/// declared once per sheet on top of a workbook-wide one. Binding follows those rules: every
/// sheet-scoped name binds, and a workbook-scoped name binds except on a sheet that declares its
/// own. Every range named <c>Items</c> reads the <c>Items</c> variable, which is what a template
/// repeating one section per sheet wants.
/// </remarks>
public class DefinedNameScopeTests
{
    private static List<SaleItem> ThreeItems() =>
        new[] { ("Widget", 2, 5m), ("Gadget", 3, 10m), ("Doohickey", 1, 7.5m) }
            .Select(r => new SaleItem
            {
                Product = r.Item1,
                Quantity = r.Item2,
                UnitPrice = r.Item3,
                SoldOn = new DateTime(2026, 1, 1),
            })
            .ToList();

    /// <summary>
    /// Adds a one-row template block at <paramref name="firstRow"/> — the row after it is the
    /// options row — and returns the range the two occupy.
    /// </summary>
    private static IXLRange TemplateBlock(IXLWorksheet sheet, int firstRow)
    {
        sheet.Cell(firstRow, 1).Value = "{{ item.Product }}";
        return sheet.Range(firstRow, 1, firstRow + 1, 3);
    }

    private static XLGenerateResult Generate(IXLWorkbook workbook, object? items)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", items);
        return template.Generate();
    }

    /// <summary>
    /// The case from issue #274: one section per sheet, each marked with a sheet-scoped name of the
    /// same name. Both used to be de-duplicated down to one, and the second sheet came out blank
    /// with nothing in <see cref="XLGenerateResult.ParsingErrors"/> to say why.
    /// </summary>
    [Test]
    public async Task TheSameSheetScopedNameOnTwoSheetsBindsOnBoth()
    {
        using var workbook = new XLWorkbook();
        var first = workbook.AddWorksheet("Q1");
        var second = workbook.AddWorksheet("Q2");
        first.DefinedNames.Add("Items", TemplateBlock(first, 1));
        second.DefinedNames.Add("Items", TemplateBlock(second, 1));

        var result = Generate(workbook, ThreeItems());

        await Assert.That(result.ParsingErrors).IsEmpty();
        foreach (var sheet in new[] { first, second })
        {
            await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
            await Assert.That(sheet.Cell("A2").Value.GetText()).IsEqualTo("Gadget");
            await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Doohickey");
        }
    }

    /// <summary>
    /// Excel resolves <c>Items</c> written on <c>Q2</c> to <c>Q2</c>'s own name. A defined name here
    /// is not read from anywhere — it is a range — so the sheet it covers stands in for the sheet a
    /// reference would be written on, and the workbook-scoped name covering that sheet is dropped.
    /// </summary>
    [Test]
    public async Task ASheetScopedNameShadowsTheWorkbookOneCoveringItsOwnSheet()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        workbook.DefinedNames.Add("Items", TemplateBlock(sheet, 1));
        sheet.DefinedNames.Add("Items", TemplateBlock(sheet, 5));

        var result = Generate(workbook, ThreeItems());

        await Assert.That(result.ParsingErrors).IsEmpty();

        // The shadowed block never expanded, so nothing below it moved.
        await Assert.That(sheet.Cell("A5").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A6").Value.GetText()).IsEqualTo("Gadget");
        await Assert.That(sheet.Cell("A7").Value.GetText()).IsEqualTo("Doohickey");
        await Assert.That(sheet.Cell("A1").IsEmpty()).IsTrue();
    }

    /// <summary>
    /// Shadowing reaches one sheet, not the workbook: a workbook-scoped name goes on binding
    /// wherever no sheet has declared its own.
    /// </summary>
    [Test]
    public async Task TheWorkbookNameStillBindsOnSheetsThatDoNotShadowIt()
    {
        using var workbook = new XLWorkbook();
        var first = workbook.AddWorksheet("Q1");
        var second = workbook.AddWorksheet("Q2");
        workbook.DefinedNames.Add("Items", TemplateBlock(first, 1));
        second.DefinedNames.Add("Items", TemplateBlock(second, 1));

        var result = Generate(workbook, ThreeItems());

        await Assert.That(result.ParsingErrors).IsEmpty();
        foreach (var sheet in new[] { first, second })
        {
            await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
            await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Doohickey");
        }
    }

    /// <summary>
    /// Excel holds defined names in one case-insensitive namespace, so a sheet-scoped <c>items</c>
    /// shadows a workbook-scoped <c>ITEMS</c>.
    /// </summary>
    /// <remarks>
    /// Driven through the binder rather than through a template, because a template matches its
    /// variables case-sensitively: neither spelling would find an <c>Items</c> variable, so both
    /// blocks would come out unbound and the two answers would be indistinguishable.
    /// </remarks>
    [Test]
    public async Task ShadowingComparesNamesTheWayExcelDoes()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        workbook.DefinedNames.Add("ITEMS", TemplateBlock(sheet, 1));
        sheet.DefinedNames.Add("items", TemplateBlock(sheet, 5));

        var variables = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["Items"] = ThreeItems(),
        };

        var bound = RangeBinder.Resolve(workbook, variables, new TemplateErrors());

        await Assert.That(bound.Count).IsEqualTo(1);
        await Assert.That(bound[0].Area.FirstRow).IsEqualTo(5);
    }

    /// <summary>
    /// A workbook-scoped name on its own is the ordinary case and is unaffected by any of this.
    /// </summary>
    [Test]
    public async Task AWorkbookNameWithNothingShadowingItBindsOnce()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        workbook.DefinedNames.Add("Items", TemplateBlock(sheet, 1));

        var result = Generate(workbook, ThreeItems());

        await Assert.That(result.ParsingErrors).IsEmpty();
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Doohickey");
        await Assert.That(sheet.Cell("A4").IsEmpty()).IsTrue();
    }

    /// <summary>
    /// A sheet-scoped name on a sheet the workbook-scoped one does not cover leaves it alone: both
    /// bind, on their own sheets, from the same variable.
    /// </summary>
    [Test]
    public async Task BothScopesBindWhenTheyCoverDifferentSheets()
    {
        using var workbook = new XLWorkbook();
        var summary = workbook.AddWorksheet("Summary");
        var detail = workbook.AddWorksheet("Detail");
        workbook.DefinedNames.Add("Items", TemplateBlock(summary, 1));
        detail.DefinedNames.Add("Items", TemplateBlock(detail, 1));

        Generate(workbook, ThreeItems());

        await Assert.That(summary.Cell("A2").Value.GetText()).IsEqualTo("Gadget");
        await Assert.That(detail.Cell("A2").Value.GetText()).IsEqualTo("Gadget");
    }
}
