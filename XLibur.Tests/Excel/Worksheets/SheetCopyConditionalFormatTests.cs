using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using XLibur.Tests.Excel.PivotTables;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// #515. A copy of a sheet keeps the conditional formats of each pivot table on it, and the rules the
/// sheet keeps only in its <c>x14</c> extension. Excel writes a rule there when it refers to another
/// sheet, a pivot table's rule (<c>pivot="1"</c>) included. XLibur does not model such a rule: it
/// writes it back from the loaded part, and a copied sheet has no loaded part, so a copy lost them,
/// and lost the pivot table's lists that name them.
/// </summary>
/// <remarks>
/// <para>
/// A copied rule keeps its id. An id need only be unique within its sheet, and the copied pivot
/// table's <c>x14:conditionalFormats</c> list then names the rule the copy holds. The rule keeps its
/// priority too.
/// </para>
/// <para>
/// Unverified: no file shows what Excel writes for a sheet copy. <c>chartex-pivotcf-before.xlsx</c> with
/// <c>Other</c> copied through Move or Copy, Create a copy, and saved by Excel, would settle the ids,
/// the priorities, and the text of a rule that names its own sheet.
/// </para>
/// </remarks>
public class SheetCopyConditionalFormatTests
{
    private const string Folder = @"Other\SheetLifecycle\";

    /// <summary>
    /// On <c>Other</c>, a pivot table with one conditional format, <c>Data!$A$2&gt;0</c>, kept only in the
    /// <c>x14</c> extensions on both sides.
    /// </summary>
    private const string PivotFixture = "chartex-pivotcf-before.xlsx";

    /// <summary>
    /// On <c>Other</c>, two rules that are not a pivot table's and are kept only in the <c>x14</c>
    /// extension: an expression on <c>C1</c>, and a colour scale on <c>C2:C4</c> whose low point is a
    /// formula. Both refer to <c>Data</c>. Excel's file fails validation on its own chart part, whatever
    /// XLibur does, so a save of it is not validated whole.
    /// </summary>
    private const string KeptRulesFixture = "rename-before.xlsx";

    private const string Copy = "OtherCopy";

    [Test]
    public async Task A_copied_sheets_pivot_table_names_the_rule_the_copy_holds()
    {
        using var saved = CopyOtherAndSave(Resource(PivotFixture));

        await Assert.That(PivotConditionalFormatLinkTests.PivotTableLists(saved, Copy))
            .IsEquivalentTo(PivotConditionalFormatLinkTests.PivotTableLists(Resource(PivotFixture)));
        var (named, broken) = PivotConditionalFormatLinkTests.Links(saved, Copy);
        await Assert.That(named).IsEqualTo(1);
        await Assert.That(broken).IsEmpty();
        await Assert.That(KeptRules(saved, Copy)).IsEquivalentTo(new[]
        {
            "pivot expression {FE99ACD7-A252-4AFF-8D90-563FB220535F} priority 1: Data!$A$2>0 on G2:G4",
        });
    }

    [Test]
    public async Task A_sheet_copy_leaves_the_original_sheets_rules_and_links_as_they_were()
    {
        using var saved = CopyOtherAndSave(Resource(PivotFixture));

        await Assert.That(ExtensionXml(saved, "Other")).IsEqualTo(ExtensionXml(Resource(PivotFixture), "Other"));
        await Assert.That(PivotConditionalFormatLinkTests.PivotTableLists(saved))
            .IsEquivalentTo(PivotConditionalFormatLinkTests.PivotTableLists(Resource(PivotFixture)));
        var (named, broken) = PivotConditionalFormatLinkTests.Links(saved);
        await Assert.That(named).IsEqualTo(1);
        await Assert.That(broken).IsEmpty();
    }

    /// <summary>
    /// The copy's rules and lists survive a reload of the saved file and a second save, so a copied rule
    /// is read back as a kept rule of the copy, as it is of the sheet it came from.
    /// </summary>
    [Test]
    public async Task A_copied_sheet_stays_linked_through_a_reload_and_a_second_save()
    {
        using var saved = CopyOtherAndSave(Resource(PivotFixture));
        using var resaved = new MemoryStream();
        using (var wb = new XLWorkbook(saved))
            wb.SaveAs(resaved, true);

        var expected = PivotConditionalFormatLinkTests.PivotTableLists(Resource(PivotFixture));
        foreach (var sheet in new[] { "Other", Copy })
        {
            await Assert.That(PivotConditionalFormatLinkTests.PivotTableLists(resaved, sheet)).IsEquivalentTo(expected);
            var (named, broken) = PivotConditionalFormatLinkTests.Links(resaved, sheet);
            await Assert.That(named).IsEqualTo(1);
            await Assert.That(broken).IsEmpty();
        }

        await Assert.That(KeptRules(resaved, Copy)).IsEquivalentTo(KeptRules(Resource(PivotFixture), "Other"));
    }

    /// <summary>
    /// A rule that is not a pivot table's, and that Excel keeps only in the <c>x14</c> extension because
    /// it refers to another sheet, is copied with its sheet: its type, id, priority, formulas and range.
    /// The copy, and the rules written into its new part, add no validator error to what a plain save of
    /// the workbook has.
    /// </summary>
    [Test]
    public async Task A_copied_sheet_keeps_the_rules_that_are_kept_only_in_x14()
    {
        using var saved = CopyOtherAndSave(Resource(KeptRulesFixture), validate: false);

        var expected = KeptRules(Resource(KeptRulesFixture), "Other");
        await Assert.That(expected).IsEquivalentTo(new[]
        {
            "expression {2EA358DE-EB2F-4288-A420-DA6A39CC4C3B} priority 2: Data!$A$2>0 on C1",
            "colorScale {9496D6DB-CF52-4009-BC20-6298A5ACDC5C} priority 1: Data!$A$2 on C2:C4",
        });
        await Assert.That(KeptRules(saved, Copy)).IsEquivalentTo(expected);
        await Assert.That(ExtensionXml(saved, "Other")).IsEqualTo(ExtensionXml(Resource(KeptRulesFixture), "Other"));

        using var untouched = LoadAndSave(Resource(KeptRulesFixture));
        await Assert.That(ValidationErrors(saved)).IsEquivalentTo(ValidationErrors(untouched));
    }

    /// <summary>
    /// A pivot table's conditional format that the sheet holds in the 2007 schema, linked by priority,
    /// is copied with its sheet, and the copy's pivot table names the rule by the priority the copy's
    /// sheet wrote it with.
    /// </summary>
    [Test]
    public async Task A_pivot_tables_modelled_conditional_format_is_copied_with_its_sheet()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = wb.AddWorksheet("Data");
            var other = wb.AddWorksheet("Other");
            data.Cell("A1").Value = "Num";
            data.Cell("B1").Value = "Label";
            data.Cell("A2").Value = 10;
            data.Cell("B2").Value = "x";
            var pivotTable = (XLPivotTable)other.PivotTables.Add("pt", other.Cell("F1"), data.Range("A1:B2"));
            pivotTable.RowLabels.Add("Label");

            // A rule of the sheet's own, which the sheet writer puts first.
            other.Range("A1:A3").AddConditionalFormat().WhenIsTrue("A1>0");

            var pivotRule = (XLConditionalFormat)other.Range("G2").AddConditionalFormat();
            pivotRule.WhenIsTrue("G2>0");
            ((XLWorksheet)other).ConditionalFormats.Remove(f => f == pivotRule);
            pivotTable.AddConditionalFormat(new XLPivotConditionalFormat(pivotRule));

            other.CopyTo(Copy);
            wb.SaveAs(saved);
        }

        foreach (var sheet in new[] { "Other", Copy })
        {
            var (named, onSheet) = MainListLinks(saved, sheet);
            await Assert.That(named).IsNotEmpty();
            await Assert.That(onSheet.Select(r => r.Priority).ToList()).IsEquivalentTo(named);
            await Assert.That(onSheet.Select(r => r.Formula).ToList()).IsEquivalentTo(new[] { "G2>0" });
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        foreach (var sheet in new[] { "Other", Copy })
        {
            var table = (XLPivotTable)reloaded.Worksheet(sheet).PivotTables.Single();
            var format = table.ConditionalFormats.Single().Format;
            await Assert.That(format.Values[1].Value).IsEqualTo("G2>0");
            await Assert.That(format.Ranges.Single().RangeAddress.ToStringRelative(false)).IsEqualTo("G2:G2");
        }
    }

    /// <summary>
    /// An edit after the copy reaches the copy's kept rules as it reaches the original's, and leaves both
    /// pivot tables naming a rule of their own sheet: a rename and a delete of the sheet the rule refers
    /// to (#498), an insert there, and an insert on the copy (#509).
    /// </summary>
    [Test]
    [Arguments("rename Data", "Renamed!$A$2>0")]
    [Arguments("delete Data", "#REF!>0")]
    [Arguments("insert row on Data", "Data!$A$3>0")]
    [Arguments("insert row on OtherCopy", "Data!$A$2>0")]
    [Arguments("insert column on OtherCopy", "Data!$A$2>0")]
    public async Task An_edit_after_a_copy_leaves_both_sheets_linked(string edit, string formula)
    {
        using var saved = CopyOtherAndSave(Resource(PivotFixture), wb => Edit(wb, edit));

        foreach (var sheet in new[] { "Other", Copy })
        {
            var (named, broken) = PivotConditionalFormatLinkTests.Links(saved, sheet);
            await Assert.That(named).IsEqualTo(1);
            await Assert.That(broken).IsEmpty();
            await Assert.That(KeptFormulas(saved, sheet)).IsEquivalentTo(new[] { formula });
        }
    }

    /// <summary>
    /// A kept rule whose formula names the copied sheet itself is copied as XLibur copies a modelled rule
    /// with its sheet: the text is not changed, so the copy's rule goes on naming the sheet it was copied
    /// from. A cell formula on the copy is repointed at the copy instead.
    /// </summary>
    /// <remarks>
    /// Unverified: this pins XLibur's behaviour, not Excel's. A copy of a sheet whose rule names that
    /// sheet, saved by Excel, would show whether Excel repoints the rule at the copy.
    /// </remarks>
    [Test]
    public async Task A_kept_rule_that_names_its_own_sheet_is_copied_unchanged_as_a_modelled_rule_is()
    {
        using var source = WithKeptExpression(Resource(KeptRulesFixture), "Other!$A$1>0");
        using var saved = CopyOtherAndSave(source, validate: false, before: wb =>
            wb.Worksheet("Other").Range("E1").AddConditionalFormat().WhenIsTrue("Other!$A$1>0")
                .Fill.SetBackgroundColor(XLColor.Red));

        saved.Position = 0;
        List<string> modelled;
        using (var document = SpreadsheetDocument.Open(saved, false))
        {
            modelled = PivotConditionalFormatLinkTests.Sheet(document, Copy).Worksheet!
                .Elements<S.ConditionalFormatting>()
                .SelectMany(c => c.Descendants<S.Formula>())
                .Select(f => f.Text)
                .Distinct()
                .ToList();
        }

        // Distinct: a copy of this sheet writes the modelled rule twice, before #515 as after it
        // (Known_gap_a_copy_of_this_sheet_doubles_a_rule_XLibur_models).
        var kept = KeptRules(saved, Copy).Single(r => r.StartsWith("expression", StringComparison.Ordinal));
        await Assert.That(modelled).IsEquivalentTo(new[] { "Other!$A$1>0" });
        await Assert.That(kept).IsEqualTo(
            "expression {2EA358DE-EB2F-4288-A420-DA6A39CC4C3B} priority 2: Other!$A$1>0 on C1");
    }

    /// <summary>
    /// KNOWN GAP, not #515's and a follow-up candidate: a copy of <c>Other</c> in Excel's
    /// <c>rename-before.xlsx</c> holds a rule XLibur models twice, though the sheet holds it once. The
    /// same rule on a sheet built in code is copied once. The copy doubled it before #515 changed
    /// anything, and it does so with the sheet's pivot table deleted first.
    /// </summary>
    /// <remarks>
    /// Likely path, not proven: <c>XLCellCopyHelper.CopyFromRange</c> copies every rule over its source
    /// range whatever the copy options say, and something on the loaded sheet (its array formula, its
    /// hyperlink or its drawing) reaches it during the copy, on top of <c>XLConditionalFormat.CopyTo</c>.
    /// </remarks>
    [Test]
    [Arguments("loaded", 2)]
    [Arguments("built", 1)]
    public async Task Known_gap_a_copy_of_this_sheet_doubles_a_rule_XLibur_models(string kind, int onCopy)
    {
        using var wb = kind == "loaded" ? new XLWorkbook(Resource(KeptRulesFixture)) : new XLWorkbook();
        var other = kind == "loaded" ? wb.Worksheet("Other") : wb.AddWorksheet("Other");
        other.Range("E1").AddConditionalFormat().WhenIsTrue("Other!$A$1>0").Fill.SetBackgroundColor(XLColor.Red);

        var copy = (XLWorksheet)other.CopyTo(Copy);

        await Assert.That(((XLWorksheet)other).ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(copy.ConditionalFormats.Count()).IsEqualTo(onCopy);
    }

    /// <summary>
    /// KNOWN GAP (#515): a pivot table copied on its own, to another cell or another sheet, copies
    /// neither list of its conditional formats, and no rule. Their ranges and pivot areas would have to
    /// move with the table, and no Excel file shows what that looks like. Nothing is left dangling: the
    /// copy names no rule, and the original keeps its rule and its link.
    /// </summary>
    [Test]
    [Arguments("Other", "J1")]
    [Arguments("Target", "B2")]
    public async Task Known_gap_a_pivot_table_copied_on_its_own_drops_its_conditional_formats(string sheet,
        string cell)
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(PivotFixture)))
        {
            var target = wb.Worksheets.TryGetWorksheet(sheet, out var found) ? found : wb.AddWorksheet(sheet);
            var copy = (XLPivotTable)wb.Worksheet("Other").PivotTables.Single().CopyTo(target.Cell(cell));
            await Assert.That(copy.ConditionalFormats).IsEmpty();
            await Assert.That(copy.ExtensionConditionalFormats).IsEmpty();
            wb.SaveAs(saved, true);
        }

        var original = PivotConditionalFormatLinkTests.PivotTableLists(Resource(PivotFixture)).Single();
        var expected = sheet == "Other"
            ? new[] { original, "no x14:conditionalFormats" }
            : new[] { "no x14:conditionalFormats" };
        await Assert.That(PivotConditionalFormatLinkTests.PivotTableLists(saved, sheet)).IsEquivalentTo(expected);
        await Assert.That(KeptRules(saved, "Other")).IsEquivalentTo(KeptRules(Resource(PivotFixture), "Other"));
        var (named, broken) = PivotConditionalFormatLinkTests.Links(saved);
        await Assert.That(named).IsEqualTo(1);
        await Assert.That(broken).IsEmpty();
        if (sheet != "Other")
            await Assert.That(KeptRules(saved, sheet)).IsEmpty();
    }

    /// <summary>
    /// Loads <paramref name="source"/>, runs <paramref name="before"/>, copies <c>Other</c> to
    /// <c>OtherCopy</c>, runs <paramref name="after"/>, and saves, validating the package unless told
    /// not to (see <see cref="KeptRulesFixture"/>).
    /// </summary>
    private static MemoryStream CopyOtherAndSave(Stream source, Action<XLWorkbook>? after = null,
        Action<XLWorkbook>? before = null, bool validate = true)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            before?.Invoke(wb);
            wb.Worksheet("Other").CopyTo(Copy);
            after?.Invoke(wb);
            wb.SaveAs(ms, validate);
        }

        ms.Position = 0;
        return ms;
    }

    private static MemoryStream LoadAndSave(Stream source)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook(source))
            wb.SaveAs(ms);

        ms.Position = 0;
        return ms;
    }

    private static void Edit(XLWorkbook wb, string edit)
    {
        switch (edit)
        {
            case "rename Data":
                wb.Worksheet("Data").Name = "Renamed";
                break;
            case "delete Data":
                wb.Worksheet("Data").Delete();
                break;
            case "insert row on Data":
                wb.Worksheet("Data").Row(1).InsertRowsAbove(1);
                break;
            case "insert row on OtherCopy":
                wb.Worksheet(Copy).Row(1).InsertRowsAbove(1);
                break;
            case "insert column on OtherCopy":
                wb.Worksheet(Copy).Column(1).InsertColumnsBefore(1);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(edit), edit, null);
        }
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));

    /// <summary>
    /// A copy of <paramref name="source"/> whose expression rule on <c>Other</c>, kept only in the
    /// <c>x14</c> extension, has the formula <paramref name="formula"/>.
    /// </summary>
    private static MemoryStream WithKeptExpression(Stream source, string formula)
    {
        var ms = new MemoryStream();
        using (source)
            source.CopyTo(ms);

        ms.Position = 0;
        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var rule = PivotConditionalFormatLinkTests.Sheet(document, "Other").Worksheet!
                .Descendants<X14.ConditionalFormattingRule>()
                .Single(r => r.Type?.Value == S.ConditionalFormatValues.Expression);
            rule.GetFirstChild<OfficeExcel.Formula>()!.Text = formula;
        }

        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// Each rule <paramref name="sheetName"/> holds in its <c>x14</c> extension: whether it is a pivot
    /// table's, its type, id and priority, its formulas, and its range.
    /// </summary>
    private static List<string> KeptRules(Stream package, string sheetName)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return PivotConditionalFormatLinkTests.Sheet(document, sheetName).Worksheet!
            .Descendants<X14.ConditionalFormattingRule>()
            .Select(rule =>
            {
                var block = (X14.ConditionalFormatting)rule.Parent!;
                var pivot = block.Pivot?.Value == true ? "pivot " : string.Empty;
                var formulas = string.Join("|", rule.Descendants<OfficeExcel.Formula>().Select(f => f.Text));
                var sqref = block.GetFirstChild<OfficeExcel.ReferenceSequence>()?.Text;
                return $"{pivot}{rule.Type?.InnerText} {rule.Id?.Value} priority {rule.Priority?.InnerText}: " +
                       $"{formulas} on {sqref}";
            })
            .ToList();
    }

    private static List<string> KeptFormulas(Stream package, string sheetName)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return PivotConditionalFormatLinkTests.Sheet(document, sheetName).Worksheet!
            .Descendants<X14.ConditionalFormattingRule>()
            .SelectMany(rule => rule.Descendants<OfficeExcel.Formula>())
            .Select(f => f.Text)
            .ToList();
    }

    private static string ExtensionXml(Stream package, string sheetName)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return PivotConditionalFormatLinkTests.Sheet(document, sheetName).Worksheet!
            .Descendants<X14.ConditionalFormattings>().Single().OuterXml;
    }

    /// <summary>Every <c>OpenXmlValidator</c> error of a saved package, one line each.</summary>
    private static List<string> ValidationErrors(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}")
            .Order(StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>
    /// The priorities the pivot tables on <paramref name="sheetName"/> name in their 2007-schema
    /// <c>conditionalFormats</c> list, and each <c>pivot="1"</c> rule the sheet holds in that schema.
    /// </summary>
    private static (List<uint> Named, List<(uint Priority, string Formula)> OnSheet) MainListLinks(Stream package,
        string sheetName)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var sheet = PivotConditionalFormatLinkTests.Sheet(document, sheetName);
        var named = sheet.PivotTableParts
            .SelectMany(p => p.PivotTableDefinition!.ConditionalFormats?.Elements<S.ConditionalFormat>() ?? [])
            .Select(cf => cf.Priority!.Value)
            .ToList();
        var onSheet = sheet.Worksheet!.Elements<S.ConditionalFormatting>()
            .Where(cf => cf.Pivot?.Value == true)
            .SelectMany(cf => cf.Elements<S.ConditionalFormattingRule>())
            .Select(r => ((uint)r.Priority!.Value, r.Elements<S.Formula>().Single().Text))
            .ToList();
        return (named, onSheet);
    }
}
