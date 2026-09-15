using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// #525, against the <c>dv-copy-*</c> fixtures, as <see cref="SheetLifecycleDataValidationFixtureTests"/> does
/// for a rename and a delete: load the workbook as it was before the copy, copy <c>Sheet1</c> through XLibur as
/// Excel's Move or Copy, Create a copy did, save, and compare every data validation of every sheet with the
/// workbook Excel saved after the copy.
/// </summary>
/// <remarks>
/// <para>
/// <c>dv-copy-before</c> has the sheets <c>Sheet1</c> and <c>Data</c>, with 1, 2, 3 and 10, 20, 30 in
/// <c>A1:A3</c>. Each cell of <c>Sheet1!B1:B7</c> has one rule: a list over <c>Sheet1!$A$1:$A$3</c>, a list
/// over <c>Data!$A$1:$A$3</c>, a whole-number between rule from <c>Sheet1!$A$1</c> to <c>Sheet1!$A$3</c>, the
/// list of <c>B1</c> without the sheet's name, a list through <c>OFFSET(Sheet1!$A$1,0,0,3,1)</c>, a custom
/// formula <c>Sheet1!$A$1&gt;0</c>, and a whole-number between rule from <c>Sheet1!$A$1</c> to
/// <c>Data!$A$3</c>.
/// </para>
/// <para>
/// Excel drops the name of a rule's own sheet from the rule's formulas, both when the rule is entered and when
/// a file is loaded, so no file Excel saves has <c>Sheet1!</c> in a rule on <c>Sheet1</c>. XLibur's
/// <c>List(IXLRange)</c> does keep it. So the rules of <c>dv-copy-before</c> were entered in Excel and then
/// changed in the XML to name their own sheet. Excel opened that file, copied <c>Sheet1</c> after itself as
/// <c>Sheet1 (2)</c>, and saved <c>dv-copy-after</c>.
/// </para>
/// <para>
/// In Excel's copy, each reference to the original sheet, named or not, refers to the copy's own cells.
/// Each reference to <c>Data</c> still refers to <c>Data</c>.
/// </para>
/// <para>
/// A rule is compared cell by cell: its type, its operator, and the text of both formulas, without a leading
/// <c>=</c> and without the name of the rule's own sheet. Excel never writes that name. XLibur writes it
/// everywhere except before a plain range (<c>DataValidationWriter</c>). A reference without a sheet name is
/// to a cell of the rule's own sheet, so the two forms refer to the same cells. The form a rule is written
/// in, standard or <c>x14</c>, is not compared.
/// </para>
/// </remarks>
public class SheetCopyDataValidationFixtureTests
{
    private const string Folder = @"Other\SheetLifecycle\";
    private const string Before = "dv-copy-before.xlsx";
    private const string After = "dv-copy-after.xlsx";
    private const string Copy = "Sheet1 (2)";

    [Test]
    public async Task The_saved_validations_match_Excel()
    {
        using var saved = CopyAndSave();

        await Assert.That(Lines(Read(saved))).IsEqualTo(Lines(Read(Resource(After))));
    }

    /// <summary>
    /// The model matches the model of Excel's file as XLibur loads it, straight after the copy and after
    /// a save and a reload.
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
            CopySheet1(wb);
            await Assert.That(Lines(Model(wb))).IsEqualTo(Lines(excel));
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(Lines(Model(reloaded))).IsEqualTo(Lines(excel));
    }

    /// <summary>
    /// #525 as reported: a list built from a range of its own sheet refers, on the copy, to the copy's range.
    /// The copy saves it as Excel does, in the standard form and without a sheet name, where it used to be
    /// saved in the <c>x14</c> extension as a reference to the original sheet.
    /// </summary>
    [Test]
    public async Task A_list_over_a_range_of_its_own_sheet_refers_to_the_copys_range()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "a";
            ws.Cell("A2").Value = "b";
            ws.Cell("A3").Value = "c";
            ws.Cell("B1").CreateDataValidation().List(ws.Range("A1:A3"));

            var copy = ws.CopyTo("Copy");

            await Assert.That(copy.Cell("B1").GetDataValidation().MinValue).IsEqualTo("Copy!$A$1:$A$3");
            await Assert.That(ws.Cell("B1").GetDataValidation().MinValue).IsEqualTo("Sheet1!$A$1:$A$3");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var document = SpreadsheetDocument.Open(saved, false);
        var worksheet = SheetPart(document.WorkbookPart!, "Copy").Worksheet!;
        await Assert.That(worksheet.Descendants<X14.DataValidation>()).IsEmpty();
        await Assert.That(worksheet.Descendants<S.DataValidation>().Select(dv => dv.Formula1!.Text).ToList())
            .IsEquivalentTo(new[] { "$A$1:$A$3" });
    }

    private static void CopySheet1(XLWorkbook wb) => wb.Worksheet("Sheet1").CopyTo(Copy, 2);

    private static MemoryStream CopyAndSave()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(Before)))
        {
            CopySheet1(wb);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return ms;
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));

    private static string Lines(IEnumerable<string> items) => string.Join(Environment.NewLine, items);

    /// <summary>
    /// A formula as the rewrite reads it, without a leading <c>=</c>, and without the name of
    /// <paramref name="sheetName"/>, the rule's own sheet, quoted or not.
    /// </summary>
    private static string Criterion(string? text, string sheetName)
    {
        var formula = text is ['=', .. var rest] ? rest : text ?? string.Empty;
        return formula
            .Replace("'" + sheetName.Replace("'", "''") + "'!", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(sheetName + "!", string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Each area of a <c>sqref</c>, a one-cell range written as its cell. XLibur writes <c>B1:B1</c> where
    /// Excel writes <c>B1</c>, and Excel joins rules that XLibur keeps apart, so a rule is compared by cell.
    /// </summary>
    private static IEnumerable<string> Areas(string? sqref)
        => (sqref ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(area => area.Split(':') is [var first, var last] && first == last ? first : area);

    /// <summary>
    /// One line for each area of each rule: the sheet and the area, the rule's type and operator, and its
    /// two formulas as <see cref="Criterion"/> reads them.
    /// </summary>
    private static IEnumerable<string> Describe(string sheetName, string? sqref, string type, string? op,
        string? formula1, string? formula2)
    {
        var comparison = type is "whole" or "decimal" or "date" or "time" or "textLength"
            ? " " + (op ?? "between")
            : string.Empty;
        return Areas(sqref).Select(area =>
            $"{sheetName}!{area}: {type}{comparison} | {Criterion(formula1, sheetName)} | {Criterion(formula2, sheetName)}");
    }

    /// <summary>The data validations of every sheet of a saved package, in either form.</summary>
    private static List<string> Read(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        var lines = new List<string>();
        foreach (var sheet in workbookPart.Workbook!.Sheets!.Elements<S.Sheet>())
        {
            var sheetName = sheet.Name!.Value!;
            var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;

            // An absent type is "none" and an absent operator "between", the schema's defaults.
            lines.AddRange(worksheet.Elements<S.DataValidations>()
                .SelectMany(d => d.Elements<S.DataValidation>())
                .SelectMany(dv => Describe(sheetName, dv.SequenceOfReferences?.InnerText, dv.Type?.InnerText ?? "none",
                    dv.Operator?.InnerText, dv.Formula1?.Text, dv.Formula2?.Text)));
            lines.AddRange(worksheet.Descendants<X14.DataValidation>()
                .SelectMany(dv => Describe(sheetName, dv.ReferenceSequence?.Text, dv.Type?.InnerText ?? "none",
                    dv.Operator?.InnerText, dv.DataValidationForumla1?.InnerText,
                    dv.DataValidationForumla2?.InnerText)));
        }

        return lines.Order(StringComparer.Ordinal).ToList();
    }

    /// <summary>Every rule of every sheet as XLibur holds it, one line for each area.</summary>
    private static List<string> Model(XLWorkbook wb)
        => wb.Worksheets
            .SelectMany(ws => ws.DataValidations.SelectMany(dv =>
            {
                var comparison = dv.AllowedValues is XLAllowedValues.WholeNumber or XLAllowedValues.Decimal
                    or XLAllowedValues.Date or XLAllowedValues.Time or XLAllowedValues.TextLength
                    ? " " + dv.Operator
                    : string.Empty;
                var sqref = string.Join(" ", dv.Ranges.Select(r => r.RangeAddress.ToStringRelative(false)));
                return Areas(sqref).Select(area =>
                    $"{ws.Name}!{area}: {dv.AllowedValues}{comparison} | {Criterion(dv.MinValue, ws.Name)} | " +
                    $"{Criterion(dv.MaxValue, ws.Name)}");
            }))
            .Order(StringComparer.Ordinal)
            .ToList();

    private static WorksheetPart SheetPart(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().Single(s => s.Name == sheetName);
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }
}
