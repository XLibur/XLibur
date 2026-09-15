using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// #507 (D81). A pivot table names its conditional format rules, by rule id, in the
/// <c>x14:conditionalFormats</c> list of its own <c>x14</c> extension. The rules themselves are in the
/// sheet's <c>x14</c> extension, marked <c>pivot="1"</c>. The pivot table writer rebuilt the extension
/// from the model and left the list out, so every save cut the link, with or without an edit.
/// </summary>
/// <remarks>
/// The fixture is <c>chartex-pivotcf-before.xlsx</c>, which Excel saved. On <c>Other</c> it has a pivot
/// table over <c>Data!A1:B4</c> with one conditional format, <c>Data!$A$2&gt;0</c> on the values. Because
/// the rule refers to another sheet, Excel writes it only in the <c>x14</c> extensions, on both sides.
/// <see cref="Worksheets.SheetLifecycleChartExFixtureTests"/> compares the list with the files Excel saved
/// after a rename and a delete.
/// </remarks>
public class PivotConditionalFormatLinkTests
{
    private const string Folder = @"Other\SheetLifecycle\";
    private const string Before = "chartex-pivotcf-before.xlsx";

    [Test]
    public async Task A_load_and_save_keeps_the_pivot_tables_list_of_its_conditional_formats()
    {
        using var saved = LoadEditAndSave("none");

        await Assert.That(Lines(PivotTableLists(saved))).IsEqualTo(Lines(PivotTableLists(Resource(Before))));
    }

    /// <summary>
    /// The list survives a reload of the saved file and a second save, so the reader reads back what
    /// the writer wrote.
    /// </summary>
    [Test]
    public async Task A_reloaded_file_saves_the_list_again()
    {
        using var saved = LoadEditAndSave("none");
        using var resaved = new MemoryStream();
        using (var wb = new XLWorkbook(saved))
            wb.SaveAs(resaved, true);

        await Assert.That(Lines(PivotTableLists(resaved))).IsEqualTo(Lines(PivotTableLists(Resource(Before))));
    }

    /// <summary>
    /// The pivot table's <c>x14</c> extension need not be the first in its <c>extLst</c>. The reader
    /// looked for <c>x14:pivotTableDefinition</c> only in the first extension, so another extension ahead
    /// of it hid the list, and a save cut the link.
    /// </summary>
    /// <remarks>
    /// Excel wrote the fixture with the <c>x14</c> extension first and its <c>xpdl</c> extension after
    /// it. The schema allows them in any order, so this test moves the <c>x14</c> extension to the end.
    /// </remarks>
    [Test]
    public async Task The_list_is_kept_when_another_extension_comes_before_it()
    {
        using var source = Resource(Before);
        using var reordered = WithX14ExtensionLast(source);
        await Assert.That(FirstExtensionHoldsX14(reordered)).IsFalse();

        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(reordered))
            wb.SaveAs(saved, true);

        await Assert.That(Lines(PivotTableLists(saved))).IsEqualTo(Lines(PivotTableLists(Resource(Before))));
        var (named, broken) = Links(saved);
        await Assert.That(named).IsEqualTo(1);
        await Assert.That(broken).IsEmpty();
    }

    /// <summary>
    /// Each rule id the pivot table names is a rule on its sheet, in a <c>pivot="1"</c> block, with the
    /// same priority. The sheet keeps the rule as it was loaded, id and priority included, so an edit
    /// that rewrites or moves the rule must not break the link.
    /// </summary>
    /// <remarks>
    /// The inserts are XLibur's own behaviour: no Excel file here shows one. The list names fields and
    /// items of the pivot table, not cells, so an insert has nothing in it to move.
    /// </remarks>
    [Test]
    [Arguments("none")]
    [Arguments("rename Data")]
    [Arguments("delete Data")]
    [Arguments("insert row on Data")]
    [Arguments("insert column on Data")]
    [Arguments("insert row on Other")]
    [Arguments("insert column on Other")]
    public async Task Every_rule_the_pivot_table_names_is_a_pivot_rule_of_its_sheet(string edit)
    {
        using var saved = LoadEditAndSave(edit);
        var (named, broken) = Links(saved);

        await Assert.That(named).IsEqualTo(1);
        await Assert.That(broken).IsEmpty();
    }

    /// <summary>
    /// The pivot table's list in the 2007 schema, <c>conditionalFormats</c>, names a rule the sheet holds
    /// in the 2007 schema by its priority. The sheet writer gives every rule a new priority on save, and
    /// it runs before the pivot table writer, so the pivot table names the rule by its new priority, and
    /// the saved file loads.
    /// </summary>
    [Test]
    public async Task The_main_list_names_its_rule_by_the_priority_the_sheet_wrote()
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
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using (var document = SpreadsheetDocument.Open(saved, false))
        {
            var sheet = Other(document);
            var named = sheet.PivotTableParts
                .SelectMany(p => p.PivotTableDefinition!.ConditionalFormats!.Elements<S.ConditionalFormat>())
                .Select(cf => cf.Priority!.Value)
                .ToList();
            var onSheet = sheet.Worksheet!.Elements<S.ConditionalFormatting>()
                .Where(cf => cf.Pivot?.Value == true)
                .SelectMany(cf => cf.Elements<S.ConditionalFormattingRule>())
                .Select(r => (uint)r.Priority!.Value)
                .ToList();

            await Assert.That(named).IsEquivalentTo(new[] { 2U });
            await Assert.That(onSheet).IsEquivalentTo(named);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var reloadedTable = (XLPivotTable)reloaded.Worksheet("Other").PivotTables.Single();
        await Assert.That(reloadedTable.ConditionalFormats.Single().Format.Priority).IsEqualTo(2);
    }

    private static MemoryStream LoadEditAndSave(string edit)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(Before)))
        {
            Edit(wb, edit);
            wb.SaveAs(ms, true);
        }

        ms.Position = 0;
        return ms;
    }

    private static void Edit(XLWorkbook wb, string edit)
    {
        switch (edit)
        {
            case "none":
                break;
            case "rename Data":
                wb.Worksheet("Data").Name = "Renamed";
                break;
            case "delete Data":
                wb.Worksheet("Data").Delete();
                break;
            case "insert row on Data":
                wb.Worksheet("Data").Row(1).InsertRowsAbove(1);
                break;
            case "insert column on Data":
                wb.Worksheet("Data").Column(1).InsertColumnsBefore(1);
                break;
            case "insert row on Other":
                wb.Worksheet("Other").Row(1).InsertRowsAbove(1);
                break;
            case "insert column on Other":
                wb.Worksheet("Other").Column(1).InsertColumnsBefore(1);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(edit), edit, null);
        }
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));

    /// <summary>
    /// A copy of <paramref name="source"/> with the <c>x14</c> extension of each pivot table on
    /// <c>Other</c> moved to the end of its <c>extLst</c>, behind the extensions Excel wrote after it.
    /// </summary>
    private static MemoryStream WithX14ExtensionLast(Stream source)
    {
        var ms = new MemoryStream();
        source.CopyTo(ms);
        ms.Position = 0;
        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            foreach (var part in Other(document).PivotTableParts)
            {
                var extList = part.PivotTableDefinition!.GetFirstChild<S.PivotTableDefinitionExtensionList>()!;
                var x14 = extList.Elements<S.PivotTableDefinitionExtension>()
                    .Single(ext => ext.GetFirstChild<X14.PivotTableDefinition>() is not null);
                x14.Remove();
                extList.AppendChild(x14);
            }
        }

        ms.Position = 0;
        return ms;
    }

    private static bool FirstExtensionHoldsX14(Stream package)
    {
        package.Position = 0;
        bool holds;
        using (var document = SpreadsheetDocument.Open(package, false))
        {
            holds = Other(document).PivotTableParts.Any(p =>
                p.PivotTableDefinition!.GetFirstChild<S.PivotTableDefinitionExtensionList>()!
                    .GetFirstChild<S.PivotTableDefinitionExtension>()!
                    .GetFirstChild<X14.PivotTableDefinition>() is not null);
        }

        package.Position = 0;
        return holds;
    }

    private static string Lines(IEnumerable<string> items) => string.Join(Environment.NewLine, items);

    /// <summary>
    /// For each pivot table on <paramref name="sheetName"/>, its <c>x14:conditionalFormats</c> list in
    /// full, attributes in name order so the comparison does not depend on the order an attribute was
    /// written in.
    /// </summary>
    internal static List<string> PivotTableLists(Stream package, string sheetName = "Other")
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return Sheet(document, sheetName).PivotTableParts
            .Select(p => p.PivotTableDefinition!.Descendants<X14.ConditionalFormats>().SingleOrDefault())
            .Select(list => list is null ? "no x14:conditionalFormats" : Canonical(list))
            .ToList();
    }

    /// <summary>
    /// How many rules the pivot tables on <paramref name="sheetName"/> name, and each one that is not a
    /// rule of the sheet in a <c>pivot="1"</c> block with the same priority.
    /// </summary>
    internal static (int Named, List<string> Broken) Links(Stream package, string sheetName = "Other")
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var other = Sheet(document, sheetName);
        var pivotRules = other.Worksheet!.Descendants<X14.ConditionalFormattingRule>()
            .Where(r => r.Parent is X14.ConditionalFormatting { Pivot.Value: true })
            .ToDictionary(r => r.Id?.Value ?? string.Empty, r => r.Priority?.InnerText);

        var named = other.PivotTableParts
            .SelectMany(p => p.PivotTableDefinition!.Descendants<X14.ConditionalFormat>())
            .ToList();
        var broken = named
            .Where(cf => !pivotRules.TryGetValue(cf.Id?.Value ?? string.Empty, out var priority) ||
                         priority != cf.Priority?.InnerText)
            .Select(cf => $"{cf.Id?.Value} priority {cf.Priority?.InnerText}")
            .ToList();
        return (named.Count, broken);
    }

    private static WorksheetPart Other(SpreadsheetDocument document) => Sheet(document, "Other");

    internal static WorksheetPart Sheet(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().Single(s => s.Name == sheetName);
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }

    internal static string Canonical(OpenXmlElement element)
    {
        var attributes = element.GetAttributes()
            .Select(a => $"{a.LocalName}={a.Value}")
            .Order(StringComparer.Ordinal);
        var children = element.Elements().Select(Canonical);
        return $"{element.LocalName}[{string.Join(" ", attributes)}]({string.Join(" ", children)})";
    }
}
