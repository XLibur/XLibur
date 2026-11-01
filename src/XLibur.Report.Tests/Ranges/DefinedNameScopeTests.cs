using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Ranges;

/// <summary>
/// Which defined names bind when more than one of them shares a name, and how a name is matched to
/// the variable it selects.
/// </summary>
/// <remarks>
/// Excel scopes a defined name either to the workbook or to one sheet, and the same name may be
/// declared once per sheet on top of a workbook-wide one. Binding follows those rules: every
/// sheet-scoped name binds, and a workbook-scoped name binds except on a sheet that declares its
/// own. Every range named <c>Items</c> reads the <c>Items</c> variable, which is what a template
/// repeating one section per sheet wants — and it reads it whatever case the name is typed in,
/// because that is the one namespace Excel keeps names in.
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
    /// shadows a workbook-scoped <c>ITEMS</c> — and both spellings select the <c>Items</c> variable,
    /// which is what makes the shadowing observable through a template at all.
    /// </summary>
    [Test]
    public async Task ShadowingComparesNamesTheWayExcelDoes()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        workbook.DefinedNames.Add("ITEMS", TemplateBlock(sheet, 1));
        sheet.DefinedNames.Add("items", TemplateBlock(sheet, 5));

        var result = Generate(workbook, ThreeItems());

        await Assert.That(result.ParsingErrors).IsEmpty();

        // The shadowed block never expanded, so nothing below it moved.
        await Assert.That(sheet.Cell("A5").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A7").Value.GetText()).IsEqualTo("Doohickey");
        await Assert.That(sheet.Cell("A1").IsEmpty()).IsTrue();
    }

    /// <summary>
    /// The case from issue #308. A defined name is an Excel identifier, written in Excel's name box
    /// and held in Excel's case-insensitive namespace, so the case it is typed in does not decide
    /// whether it finds its variable. It used to: the block was left unbound and rendered blank,
    /// with nothing in <see cref="XLGenerateResult.ParsingErrors"/> to say why.
    /// </summary>
    [Test]
    [Arguments("ITEMS")]
    [Arguments("items")]
    [Arguments("iTeMs")]
    public async Task ADefinedNameFindsItsVariableWhateverCaseItIsTypedIn(string definedName)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        workbook.DefinedNames.Add(definedName, TemplateBlock(sheet, 1));

        var result = Generate(workbook, ThreeItems());

        await Assert.That(result.ParsingErrors).IsEmpty();
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Doohickey");
    }

    /// <summary>
    /// A template holding two variables that differ only by case keeps the one the name spells
    /// exactly, so no template that binds today binds anything different.
    /// </summary>
    [Test]
    public async Task AnExactMatchWinsOverOneDifferingOnlyByCase()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        workbook.DefinedNames.Add("Items", TemplateBlock(sheet, 1));

        using var template = new XLTemplate(workbook);
        template.AddVariable("items", ThreeItems());
        template.AddVariable("Items", ThreeItems().Take(1).ToList());
        var result = template.Generate();

        await Assert.That(result.ParsingErrors).IsEmpty();
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A2").IsEmpty()).IsTrue();
    }

    /// <summary>
    /// Two variables differing only by case, and a name spelling neither of them: the one case a
    /// case-insensitive match has no answer for. Binding either would be binding whichever the
    /// dictionary yielded first, so the name is reported instead and nothing binds.
    /// </summary>
    [Test]
    public async Task VariablesDifferingOnlyByCaseAreReportedRatherThanGuessedAt()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        workbook.DefinedNames.Add("ITEMS", TemplateBlock(sheet, 1));

        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", ThreeItems());
        template.AddVariable("items", ThreeItems());
        var result = template.Generate();

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors.Count).IsEqualTo(1);
        await Assert.That(result.ParsingErrors[0].Message).Contains("'ITEMS'");
        await Assert.That(result.ParsingErrors[0].Message).Contains("differ only by case");
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
