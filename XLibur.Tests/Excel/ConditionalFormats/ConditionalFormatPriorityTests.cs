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

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// #552. A save gives every conditional format rule on a sheet its own priority: the rules XLibur models
/// and the rules it keeps only in the sheet's <c>x14</c> extension. A save used to number the modelled
/// rules 1, 2, 3 and so on, and write each kept rule with the priority it was loaded with, so a modelled
/// rule and a kept rule could share a priority.
/// </summary>
/// <remarks>
/// <para>
/// <c>cf-copy-before.xlsx</c>, which Excel saved, has one expression rule on each cell of
/// <c>Other!B1:B5</c>. Excel writes <c>B1</c>, <c>B4</c> and <c>B5</c> in the standard
/// <c>conditionalFormatting</c> element, with the priorities 1, 4 and 5, and <c>B2</c> and <c>B3</c>, which
/// refer to <c>Data</c>, only in the <c>x14</c> extension, with the priorities 2 and 3. A save numbered
/// the modelled rules 1, 2 and 3, so <c>B4</c> and <c>B5</c> shared 2 and 3 with <c>B2</c> and <c>B3</c>.
/// </para>
/// <para>
/// A pivot table's conditional formats name their rules by priority: the 2007 schema's list by priority
/// alone, and the <c>x14</c> list by rule id and priority. Each name follows the priority the sheet writes.
/// </para>
/// </remarks>
public class ConditionalFormatPriorityTests
{
    private const string Folder = @"Other\SheetLifecycle\";

    /// <summary>Excel's file, described in the class remarks.</summary>
    private const string Fixture = "cf-copy-before.xlsx";

    /// <summary>
    /// On <c>Other</c>, a pivot table whose one conditional format, <c>Data!$A$2&gt;0</c> on <c>G2:G4</c>,
    /// Excel keeps only in the <c>x14</c> extensions, with priority 1.
    /// </summary>
    private const string PivotFixture = "chartex-pivotcf-before.xlsx";

    /// <summary>
    /// On <c>Other</c>, a pivot table with no conditional format, and two rules kept only in the <c>x14</c>
    /// extension: a colour scale on <c>C2:C4</c> with priority 1, and an expression on <c>C1</c> with
    /// priority 2. Excel's file fails validation on its own chart part, so a save of it is not validated.
    /// </summary>
    private const string KeptRulesFixture = "rename-before.xlsx";

    private const string B1 = "B1 standard";
    private const string B2 = "B2 x14 {FB1A234C-B2AF-4E2B-8FD1-11C88052B41D}";
    private const string B3 = "B3 x14 {A0A3F7CA-B1BB-4004-95EF-F554EFAF8C40}";
    private const string B4 = "B4 standard";
    private const string B5 = "B5 standard";

    /// <summary>Excel's rules on <c>Other</c>, in Excel's order, with Excel's priorities.</summary>
    private static readonly string[] ExcelRules = [$"1: {B1}", $"2: {B2}", $"3: {B3}", $"4: {B4}", $"5: {B5}"];

    [Test]
    public async Task Excels_file_holds_the_rules_the_tests_expect()
    {
        using var excel = Resource(Fixture);

        await AssertRules(excel, "Other", ExcelRules);
    }

    /// <summary>
    /// A load and a save write every rule, modelled or kept, with the priority Excel gave it. Excel's
    /// priorities are 1 to 5 with no gap, so the only numbering from 1 that keeps its order is its own.
    /// </summary>
    [Test]
    public async Task A_load_and_save_writes_the_priorities_Excel_wrote()
    {
        using var saved = LoadEditAndSave(Resource(Fixture), _ => { });

        await AssertRules(saved, "Other", ExcelRules);
    }

    /// <summary>
    /// The saved file reads back with the same priorities, and a second save writes them again.
    /// </summary>
    [Test]
    public async Task A_reload_and_a_second_save_write_them_again()
    {
        using var saved = LoadEditAndSave(Resource(Fixture), _ => { });
        using var resaved = LoadEditAndSave(saved, _ => { });

        await AssertRules(resaved, "Other", ExcelRules);
    }

    /// <summary>
    /// Two saves of one workbook write the same priorities. The first save renumbers the modelled rules
    /// in the model, so the second must number the kept rules in the same terms, not by the priorities in
    /// the file it was loaded from.
    /// </summary>
    [Test]
    public async Task Two_saves_of_one_workbook_write_the_same_priorities()
    {
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(Fixture)))
        {
            wb.Worksheet("Other").ConditionalFormats.Remove(cf => Cell(cf) == "B1");
            wb.SaveAs(first);
            wb.SaveAs(second);
        }

        string[] expected = [$"1: {B2}", $"2: {B3}", $"3: {B4}", $"4: {B5}"];
        await AssertRules(first, "Other", expected);
        await AssertRules(second, "Other", expected);
    }

    /// <summary>
    /// A rule added in code has priority 0, so a save writes it first, as it did before (see
    /// <c>SheetCopyConditionalFormatTests.A_sheet_copy_keeps_the_order_of_a_rule_added_after_the_load</c>).
    /// Every loaded rule then moves down one, the kept ones with the modelled ones, and keeps its order.
    /// </summary>
    [Test]
    public async Task A_rule_added_in_code_moves_every_loaded_rule_down_in_Excels_order()
    {
        using var saved = LoadEditAndSave(Resource(Fixture), wb =>
            wb.Worksheet("Other").Range("C1").AddConditionalFormat().WhenIsTrue("A1>0")
                .Fill.SetBackgroundColor(XLColor.Blue));

        await AssertRules(saved, "Other",
            "1: C1 standard", $"2: {B1}", $"3: {B2}", $"4: {B3}", $"5: {B4}", $"6: {B5}");
    }

    /// <summary>
    /// A removed rule leaves no gap, and no rule takes a priority another already has.
    /// </summary>
    [Test]
    public async Task A_removed_rule_leaves_the_others_in_Excels_order()
    {
        using var saved = LoadEditAndSave(Resource(Fixture), wb =>
            wb.Worksheet("Other").ConditionalFormats.Remove(cf => Cell(cf) == "B1"));

        await AssertRules(saved, "Other", $"1: {B2}", $"2: {B3}", $"3: {B4}", $"4: {B5}");
    }

    /// <summary>
    /// A file an earlier save wrote with a clash, <c>B4</c> and <c>B5</c> on the priorities of <c>B2</c> and
    /// <c>B3</c>, is saved with a priority for each rule. On a tie the kept rule goes first: the earlier
    /// save numbered only the modelled rules, and only ever down, so a modelled rule that ties with a kept
    /// rule came after it.
    /// </summary>
    /// <remarks>
    /// Excel repairs this same file the other way: it leaves <c>B1</c>, <c>B4</c> and <c>B5</c> at 1, 2 and
    /// 3 and renumbers the kept <c>B2</c> and <c>B3</c> to 4 and 5 (driven over COM for #552). The file has
    /// lost the order Excel first wrote, <c>B1</c> to <c>B5</c>, and neither numbering recovers it. What
    /// matters here is that each rule comes out with a priority of its own. Excel leaves a file this writes
    /// from Excel's own untouched, which <see cref="A_load_and_save_writes_the_priorities_Excel_wrote"/>
    /// and the <c>cf-copy</c> fixture comparison pin.
    /// </remarks>
    [Test]
    public async Task A_file_with_a_clash_is_saved_with_a_priority_for_each_rule()
    {
        using var clashing = WithStandardPriorities(Resource(Fixture), "Other", [1, 2, 3]);
        await AssertRules(clashing, "Other", $"1: {B1}", $"2: {B4}", $"2: {B2}", $"3: {B5}", $"3: {B3}");

        using var saved = LoadEditAndSave(clashing, _ => { });

        await AssertRules(saved, "Other", $"1: {B1}", $"2: {B2}", $"3: {B4}", $"4: {B3}", $"5: {B5}");
    }

    /// <summary>
    /// A kept pivot rule that a rule added in code moves down is named by its new priority in the pivot
    /// table's <c>x14</c> list, on the sheet and on a copy of it, and the link survives a reload and a
    /// second save.
    /// </summary>
    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task A_kept_pivot_rule_is_named_by_the_priority_the_sheet_wrote(bool copy)
    {
        var sheets = copy ? new[] { "Other", "OtherCopy" } : new[] { "Other" };
        using var saved = LoadEditAndSave(Resource(PivotFixture), wb =>
        {
            wb.Worksheet("Other").Range("A1").AddConditionalFormat().WhenIsTrue("A1>0")
                .Fill.SetBackgroundColor(XLColor.Blue);
            if (copy)
                wb.Worksheet("Other").CopyTo("OtherCopy");
        });

        foreach (var sheet in sheets)
        {
            await AssertRules(saved, sheet,
                "1: A1 standard", "2: G2:G4 pivot x14 {FE99ACD7-A252-4AFF-8D90-563FB220535F}");
            await AssertLinked(saved, sheet);
        }

        using var resaved = LoadEditAndSave(saved, _ => { });
        foreach (var sheet in sheets)
        {
            await AssertRules(resaved, sheet, Rules(saved, sheet).ToArray());
            await AssertLinked(resaved, sheet);
        }
    }

    /// <summary>
    /// A pivot table's rule in the 2007 schema is named by its priority alone. Added in code with the
    /// sheet's kept rules at 1 and 2, it goes first, as any new rule does, and the kept rules move down, so
    /// the pivot table does not name a priority a kept rule has. The saved file reloads with the pivot
    /// table linked to its rule, and a second save keeps every priority.
    /// </summary>
    [Test]
    public async Task A_pivot_rule_in_the_2007_schema_does_not_share_a_priority_with_a_kept_rule()
    {
        using var saved = LoadEditAndSave(Resource(KeptRulesFixture), wb =>
        {
            var other = (XLWorksheet)wb.Worksheet("Other");
            var pivotTable = (XLPivotTable)((IXLWorksheet)other).PivotTables.Single();
            var pivotRule = (XLConditionalFormat)other.Range("G2")!.AddConditionalFormat();
            pivotRule.WhenIsTrue("G2>0");
            other.ConditionalFormats.Remove(f => f == pivotRule);
            pivotTable.AddConditionalFormat(new XLPivotConditionalFormat(pivotRule));
        });

        string[] expected =
        [
            "1: G2 pivot standard",
            "2: C2:C4 x14 {9496D6DB-CF52-4009-BC20-6298A5ACDC5C}",
            "3: C1 x14 {2EA358DE-EB2F-4288-A420-DA6A39CC4C3B}",
        ];
        await AssertRules(saved, "Other", expected);
        await Assert.That(MainListPriorities(saved, "Other")).IsEquivalentTo(new[] { 1U });

        using var resaved = new MemoryStream();
        saved.Position = 0;
        using (var reloaded = new XLWorkbook(saved))
        {
            var format = ((XLPivotTable)reloaded.Worksheet("Other").PivotTables.Single()).ConditionalFormats
                .Single().Format;
            await Assert.That(format.Values[1].Value).IsEqualTo("G2>0");
            reloaded.SaveAs(resaved, false);
        }

        await AssertRules(resaved, "Other", expected);
        await Assert.That(MainListPriorities(resaved, "Other")).IsEquivalentTo(new[] { 1U });
    }

    /// <summary>
    /// The two halves of a rule Excel writes in both schemas are one rule, and keep the one priority they
    /// were loaded with. Excel links them by id: the rule in the standard element carries
    /// <c>extLst/x14:id</c>, and the rule in the <c>x14</c> extension has that id. A custom icon set is
    /// written this way, and its <c>x14</c> half repeats the priority, where a data bar's carries none.
    /// </summary>
    /// <remarks>
    /// Numbering the <c>x14</c> half as a rule of its own would split the pair: the half in the standard
    /// element would be numbered after it and the two would name different priorities, where before they
    /// agreed. Found by review of #552. No fixture in the suite has this shape, so the file is built here.
    /// KNOWN GAP, as before #552: the <c>x14</c> half keeps the priority it was loaded with, so on a sheet
    /// whose rules changed it no longer matches the half that was renumbered. Only a data bar's id reaches
    /// the model, so there is nothing to carry a custom icon set's new priority across to its other half.
    /// </remarks>
    [Test]
    public async Task The_two_halves_of_one_rule_keep_one_priority()
    {
        using var built = WithIconSetTwin();
        await AssertRules(built, "Other", $"1: A1:A10 standard", $"1: A1:A10 x14 {TwinId}", "2: B1 standard");

        using var saved = LoadEditAndSave(built, _ => { });

        await AssertRules(saved, "Other", $"1: A1:A10 standard", $"1: A1:A10 x14 {TwinId}", "2: B1 standard");
    }

    private const string TwinId = "{6F1E2A34-5C7D-4B8E-9A01-2D3E4F5A6B7C}";

    /// <summary>
    /// A workbook whose sheet <c>Other</c> holds an icon set on <c>A1:A10</c> written in both schemas, as
    /// Excel writes a custom icon set, and an expression on <c>B1</c> after it.
    /// </summary>
    private static MemoryStream WithIconSetTwin()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var sheet = wb.AddWorksheet("Other");
            for (var row = 1; row <= 10; row++)
                sheet.Cell(row, 1).Value = row;

            sheet.Range("A1:A10").AddConditionalFormat().IconSet(XLIconSetStyle.ThreeTrafficLights1)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "0", XLCFContentType.Number)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "3", XLCFContentType.Number)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "6", XLCFContentType.Number);
            sheet.Range("B1").AddConditionalFormat().WhenIsTrue("B1>0").Fill.SetBackgroundColor(XLColor.Red);
            wb.SaveAs(ms, false);
        }

        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var worksheet = PivotConditionalFormatLinkTests.Sheet(document, "Other").Worksheet!;
            var iconSet = worksheet.Elements<S.ConditionalFormatting>()
                .SelectMany(block => block.Elements<S.ConditionalFormattingRule>())
                .Single(rule => rule.Type?.Value == S.ConditionalFormatValues.IconSet);

            iconSet.Append(new S.ConditionalFormattingRuleExtensionList(
                "<x:extLst xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<x:ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" " +
                "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" +
                $"<x14:id>{TwinId}</x14:id></x:ext></x:extLst>"));

            worksheet.Append(new S.WorksheetExtensionList(
                "<x:extLst xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<x:ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" " +
                "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" +
                "<x14:conditionalFormattings>" +
                "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
                $"<x14:cfRule type=\"iconSet\" priority=\"1\" id=\"{TwinId}\">" +
                "<x14:iconSet custom=\"1\">" +
                "<x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo>" +
                "<x14:cfvo type=\"percent\"><xm:f>33</xm:f></x14:cfvo>" +
                "<x14:cfvo type=\"percent\"><xm:f>67</xm:f></x14:cfvo>" +
                "</x14:iconSet></x14:cfRule>" +
                "<xm:sqref>A1:A10</xm:sqref>" +
                "</x14:conditionalFormatting></x14:conditionalFormattings></x:ext></x:extLst>"));
            worksheet.Save();
        }

        ms.Position = 0;
        return ms;
    }

    private static async Task AssertRules(Stream package, string sheetName, params string[] expected)
    {
        await Assert.That(string.Join(Environment.NewLine, Rules(package, sheetName)))
            .IsEqualTo(string.Join(Environment.NewLine, expected));
    }

    private static async Task AssertLinked(Stream package, string sheet)
    {
        var (named, broken) = PivotConditionalFormatLinkTests.Links(package, sheet);
        await Assert.That(named).IsEqualTo(1);
        await Assert.That(broken).IsEmpty();
    }

    private static string Cell(IXLConditionalFormat cf) => cf.Ranges.Single().RangeAddress.FirstAddress.ToString()!;

    private static MemoryStream LoadEditAndSave(Stream source, Action<XLWorkbook> edit)
    {
        var ms = new MemoryStream();
        source.Position = 0;
        using (var wb = new XLWorkbook(source))
        {
            edit(wb);
            wb.SaveAs(ms, false);
        }

        ms.Position = 0;
        return ms;
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));

    /// <summary>
    /// A copy of <paramref name="source"/> whose rules in the standard element on
    /// <paramref name="sheetName"/> have <paramref name="priorities"/>, in the order they are written.
    /// </summary>
    private static MemoryStream WithStandardPriorities(Stream source, string sheetName, int[] priorities)
    {
        var ms = new MemoryStream();
        using (source)
            source.CopyTo(ms);

        ms.Position = 0;
        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var rules = PivotConditionalFormatLinkTests.Sheet(document, sheetName).Worksheet!
                .Elements<S.ConditionalFormatting>()
                .SelectMany(block => block.Elements<S.ConditionalFormattingRule>())
                .ToList();
            for (var i = 0; i < rules.Count; i++)
                rules[i].Priority = priorities[i];
        }

        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// Every rule of <paramref name="sheetName"/> in a package, in either form, lowest priority first, and
    /// on a tie a rule in the standard element first: its priority, its range, a one-cell range written as
    /// its cell, whether it is a pivot table's, its form, and for a rule in the <c>x14</c> extension that
    /// is not a data bar, its id.
    /// </summary>
    internal static List<string> Rules(Stream package, string sheetName)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var worksheet = PivotConditionalFormatLinkTests.Sheet(document, sheetName).Worksheet!;
        var rules = new List<(int Priority, string Line)>();

        AddStandardRules(worksheet, rules);
        AddX14Rules(worksheet, rules);

        package.Position = 0;
        return rules
            .OrderBy(r => r.Priority)
            .Select(r => $"{r.Priority}: {r.Line}")
            .ToList();
    }

    private static void AddStandardRules(S.Worksheet worksheet, List<(int Priority, string Line)> rules)
    {
        foreach (var block in worksheet.Elements<S.ConditionalFormatting>())
        {
            var pivot = block.Pivot?.Value == true ? "pivot " : string.Empty;
            foreach (var rule in block.Elements<S.ConditionalFormattingRule>())
                rules.Add((rule.Priority!.Value, $"{Sqref(block.SequenceOfReferences?.InnerText)} {pivot}standard"));
        }
    }

    private static void AddX14Rules(S.Worksheet worksheet, List<(int Priority, string Line)> rules)
    {
        foreach (var block in worksheet.Descendants<X14.ConditionalFormatting>())
        {
            var pivot = block.Pivot?.Value == true ? "pivot " : string.Empty;
            var sqref = Sqref(block.GetFirstChild<OfficeExcel.ReferenceSequence>()?.Text);
            foreach (var rule in block.Elements<X14.ConditionalFormattingRule>())
            {
                var id = rule.Type?.Value == S.ConditionalFormatValues.DataBar ? string.Empty : $" {rule.Id?.Value}";
                rules.Add((rule.Priority?.Value ?? 0, $"{sqref} {pivot}x14{id}"));
            }
        }
    }

    /// <summary>
    /// A range list in one spelling: a one-cell area as its cell. XLibur writes <c>B1:B1</c> where Excel
    /// writes <c>B1</c>.
    /// </summary>
    private static string Sqref(string? sqref)
        => string.Join(" ", (sqref ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(area => area.Split(':') is [var first, var last] && first == last ? first : area));

    /// <summary>
    /// The priorities the pivot tables on <paramref name="sheetName"/> name in their 2007-schema list.
    /// </summary>
    private static List<uint> MainListPriorities(Stream package, string sheetName)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return PivotConditionalFormatLinkTests.Sheet(document, sheetName).PivotTableParts
            .SelectMany(p => p.PivotTableDefinition!.ConditionalFormats?.Elements<S.ConditionalFormat>() ?? [])
            .Select(cf => cf.Priority!.Value)
            .ToList();
    }
}
