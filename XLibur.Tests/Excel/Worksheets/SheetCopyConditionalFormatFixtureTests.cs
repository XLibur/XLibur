using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// #535, against the <c>cf-copy-*</c> fixtures, as <see cref="SheetCopyDataValidationFixtureTests"/> does for
/// data validations: load the workbook as it was before the copy, copy <c>Other</c> through XLibur as
/// Excel's Move or Copy, Create a copy did, save, and compare every conditional format of every sheet with
/// the workbook Excel saved after the copy.
/// </summary>
/// <remarks>
/// <para>
/// <c>cf-copy-before</c> has the sheets <c>Data</c>, with 10, 20, 30 in <c>A1:A3</c>, and <c>Other</c>, with
/// 1 to 5 in <c>A1:A5</c>. Each cell of <c>Other!B1:B5</c> has one expression rule with a red fill:
/// <c>Other!$A$1&gt;0</c>, <c>AND(Other!$A$1&gt;0,Data!$A$1&gt;0)</c>, <c>Data!$A$1&gt;0</c>,
/// <c>$A$1&gt;0</c>, and <c>Other!A5&gt;0</c>, whose reference is relative. Excel writes the second and the
/// third only in the sheet's <c>x14</c> extension, because they refer to another sheet, and XLibur keeps
/// them there as it loaded them. The others are in the standard <c>conditionalFormatting</c> element, and
/// XLibur models them.
/// </para>
/// <para>
/// Excel drops the name of a rule's own sheet from the rule's formula, both when the rule is entered and
/// when a file is loaded, as it does for a data validation. So the rules of <c>cf-copy-before</c> were
/// entered in Excel through its COM interface and then changed in the XML to name their own sheet. Excel
/// opened that file, copied <c>Other</c> after itself as <c>Other (2)</c>, and saved <c>cf-copy-after</c>.
/// </para>
/// <para>
/// On Excel's copy, each reference to the original sheet, named or not, refers to the copy's own cells:
/// the rules read <c>$A$1&gt;0</c>, <c>AND($A$1&gt;0,Data!$A$1&gt;0)</c>, <c>Data!$A$1&gt;0</c>,
/// <c>$A$1&gt;0</c> and <c>A5&gt;0</c>. Each reference to <c>Data</c> still refers to <c>Data</c>. Each rule
/// keeps its priority, and its place in the standard element or the <c>x14</c> extension. Each copied
/// <c>x14</c> rule has a new id, where XLibur's keeps its id (#515), so the id is not compared.
/// </para>
/// <para>
/// A rule is compared cell by cell: its priority, the form it is written in, its type, and the text of its
/// formulas without the name of the rule's own sheet. Excel never writes that name. XLibur keeps it, so on
/// the copy it names the copy, quoted where the name needs it. A reference without a sheet name is to a
/// cell of the rule's own sheet, so the two forms refer to the same cells.
/// </para>
/// <para>
/// Excel numbers the rules of each sheet 1 to 5, <c>B1</c> to <c>B5</c>, across both forms. A save used to
/// number only the rules XLibur models from 1, so <c>B4</c> and <c>B5</c> were written with the priorities
/// 2 and 3, which <c>B2</c> and <c>B3</c> have too, and this comparison left the priority out. A save
/// numbers the kept rules with the modelled ones now (#552).
/// </para>
/// </remarks>
public class SheetCopyConditionalFormatFixtureTests
{
    private const string Folder = @"Other\SheetLifecycle\";
    private const string Before = "cf-copy-before.xlsx";
    private const string After = "cf-copy-after.xlsx";
    private const string Copy = "Other (2)";

    [Test]
    public async Task The_saved_rules_match_Excel()
    {
        using var saved = CopyAndSave();

        await Assert.That(Lines(Read(saved))).IsEqualTo(Lines(Read(Resource(After))));
    }

    /// <summary>
    /// The rules XLibur models match the modelled rules of Excel's file as XLibur loads it, straight after
    /// the copy and after a save and a reload. A rule kept only in <c>x14</c> is not in the model, and is
    /// compared in the saved file (<see cref="The_saved_rules_match_Excel"/>).
    /// </summary>
    [Test]
    public async Task The_model_matches_Excel_before_and_after_a_reload()
    {
        List<string> excel;
        using (var excelBook = new XLWorkbook(Resource(After)))
            excel = Model(excelBook);

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(Before)))
        {
            CopyOther(wb);
            await Assert.That(Lines(Model(wb))).IsEqualTo(Lines(excel));
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(Lines(Model(reloaded))).IsEqualTo(Lines(excel));
    }

    /// <summary>
    /// The copy's rules name the copy, and the original's still name the original, in the text XLibur holds
    /// and saves.
    /// </summary>
    [Test]
    public async Task The_copys_rules_name_the_copy_and_the_originals_name_the_original()
    {
        using var saved = CopyAndSave();

        await Assert.That(Formulas(saved, Copy)).IsEquivalentTo(new[]
        {
            "B1: 'Other (2)'!$A$1>0",
            "B2: AND('Other (2)'!$A$1>0,Data!$A$1>0)",
            "B3: Data!$A$1>0",
            "B4: $A$1>0",
            "B5: 'Other (2)'!A5>0",
        });
        await Assert.That(Formulas(saved, "Other")).IsEquivalentTo(new[]
        {
            "B1: Other!$A$1>0",
            "B2: AND(Other!$A$1>0,Data!$A$1>0)",
            "B3: Data!$A$1>0",
            "B4: $A$1>0",
            "B5: Other!A5>0",
        });
    }

    private static void CopyOther(XLWorkbook wb) => wb.Worksheet("Other").CopyTo(Copy, 3);

    private static MemoryStream CopyAndSave()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(Before)))
        {
            CopyOther(wb);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return ms;
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));

    private static string Lines(IEnumerable<string> items) => string.Join(Environment.NewLine, items);

    /// <summary>
    /// A formula without a leading <c>=</c>, and without the name of <paramref name="sheetName"/>, the
    /// rule's own sheet, quoted or not.
    /// </summary>
    private static string WithoutOwnSheet(string? text, string sheetName)
    {
        var formula = text is ['=', .. var rest] ? rest : text ?? string.Empty;
        return formula
            .Replace("'" + sheetName.Replace("'", "''") + "'!", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(sheetName + "!", string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Each area of a <c>sqref</c>, a one-cell range written as its cell. XLibur writes <c>B1:B1</c> where
    /// Excel writes <c>B1</c>.
    /// </summary>
    private static IEnumerable<string> Areas(string? sqref)
        => (sqref ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(area => area.Split(':') is [var first, var last] && first == last ? first : area);

    /// <summary>The conditional formats of every sheet of a saved package, in either form.</summary>
    /// <remarks>
    /// One line for each area of each rule: the sheet and the area, the rule's priority, the form, the
    /// rule's type, and its formulas without the name of the rule's own sheet.
    /// </remarks>
    private static List<string> Read(Stream package)
    {
        var lines = new List<string>();
        ForEachRule(package, (sheetName, sqref, priority, form, type, formulas) =>
        {
            var text = string.Join(" | ", formulas.Select(f => WithoutOwnSheet(f, sheetName)));
            lines.AddRange(Areas(sqref)
                .Select(area => $"{sheetName}!{area}: priority {priority} {form} {type} | {text}"));
        });
        return lines.Order(StringComparer.Ordinal).ToList();
    }

    /// <summary>
    /// Each rule of <paramref name="sheetName"/> in a saved package, as its first cell and its formulas as
    /// they are written.
    /// </summary>
    private static List<string> Formulas(Stream package, string sheetName)
    {
        var lines = new List<string>();
        ForEachRule(package, (sheet, sqref, _, _, _, formulas) =>
        {
            if (sheet == sheetName)
                lines.Add($"{Areas(sqref).First()}: {string.Join(" | ", formulas)}");
        });
        return lines;
    }

    /// <summary>
    /// Calls <paramref name="action"/> for each rule of each sheet of a saved package, with the sheet's name,
    /// the rule's <c>sqref</c>, its priority, its form, its type, and its formulas as they are written.
    /// </summary>
    private static void ForEachRule(Stream package,
        Action<string, string?, int?, string, string?, IEnumerable<string>> action)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        foreach (var sheet in workbookPart.Workbook!.Sheets!.Elements<S.Sheet>())
        {
            var sheetName = sheet.Name!.Value!;
            var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;

            foreach (var block in worksheet.Elements<S.ConditionalFormatting>())
            {
                foreach (var rule in block.Elements<S.ConditionalFormattingRule>())
                {
                    action(sheetName, block.SequenceOfReferences?.InnerText, rule.Priority?.Value, "standard",
                        rule.Type?.InnerText, rule.Elements<S.Formula>().Select(f => f.Text));
                }
            }

            foreach (var block in worksheet.Descendants<X14.ConditionalFormatting>())
            {
                foreach (var rule in block.Elements<X14.ConditionalFormattingRule>())
                {
                    action(sheetName, block.GetFirstChild<OfficeExcel.ReferenceSequence>()?.Text,
                        rule.Priority?.Value, "x14", rule.Type?.InnerText,
                        rule.Descendants<OfficeExcel.Formula>().Select(f => f.Text));
                }
            }
        }
    }

    /// <summary>Every rule XLibur models on every sheet, one line for each area.</summary>
    private static List<string> Model(XLWorkbook wb)
        => wb.Worksheets
            .SelectMany(ws => ws.ConditionalFormats.SelectMany(cf =>
            {
                var formulas = string.Join(" | ", cf.Values.Values.Select(v => WithoutOwnSheet(v.Value, ws.Name)));
                var sqref = string.Join(" ", cf.Ranges.Select(r => r.RangeAddress.ToStringRelative(false)));
                return Areas(sqref).Select(area => $"{ws.Name}!{area}: {cf.ConditionalFormatType} | {formulas}");
            }))
            .Order(StringComparer.Ordinal)
            .ToList();
}
