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
/// D63, against the <c>dv-*</c> fixtures, as <see cref="SheetLifecycleFixtureTests"/> does for spec 55's
/// other holders: load the workbook the owner saved in Excel before an edit, make the same edit
/// through XLibur, save, and compare every data validation and every defined name with the workbook
/// Excel saved after it.
/// </summary>
/// <remarks>
/// <para>
/// <c>dv-before</c> has the sheets <c>Data</c> and <c>Other</c>, and the names <c>DvOnly</c> (scope
/// <c>Data</c>) and <c>ListW</c> (workbook), both over cells of <c>Data</c>. Each cell of
/// <c>Other!B1:B7</c> has one rule: a list over <c>Data</c>, a list through <c>OFFSET</c>, a custom
/// formula, a whole-number between rule with both ends on <c>Data</c>, a list through <c>ListW</c>, a
/// list through <c>INDIRECT("Data!…")</c>, and a literal list. Excel keeps the first four in the sheet's
/// <c>x14</c> extension, because they refer to another sheet. Spec 55's Results record what Excel wrote.
/// </para>
/// <para>
/// A rule is compared by its <c>sqref</c>: its type, its operator, and the text of both formulas, a
/// leading <c>=</c> left out, as the rewrite reads it. The form the rule is written in, standard or
/// <c>x14</c>, is not compared. After a delete it differs from Excel's, and
/// <see cref="Known_difference_a_deleted_sheets_rules_are_written_in_the_standard_form"/> pins it.
/// </para>
/// </remarks>
public class SheetLifecycleDataValidationFixtureTests
{
    private const string Folder = @"Other\SheetLifecycle\";
    private const string Before = "dv-before.xlsx";
    private const string Workbook = "(workbook)";

    /// <summary>What happens to the sheet <c>Data</c>.</summary>
    public enum SheetEvent
    {
        /// <summary>Renamed to <c>New Data</c>, as the owner renamed it in Excel.</summary>
        Rename,

        /// <summary><c>IXLWorksheet.Delete()</c>.</summary>
        WorksheetDelete,

        /// <summary><c>IXLWorksheets.Delete("Data")</c>.</summary>
        CollectionDelete,
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task The_saved_validations_and_names_match_Excel(SheetEvent sheetEvent)
    {
        using var saved = EditAndSave(sheetEvent);

        await AssertSameText(Read(saved), Read(Resource(After(sheetEvent))));
    }

    /// <summary>
    /// The model matches the model of Excel's file as XLibur loads it, straight after the edit and
    /// after a save and a reload.
    /// </summary>
    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    [Arguments(SheetEvent.CollectionDelete)]
    public async Task The_model_matches_Excel_before_and_after_a_reload(SheetEvent sheetEvent)
    {
        List<string> excel;
        using (var excelBook = new XLWorkbook(Resource(After(sheetEvent))))
            excel = Model(excelBook);

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(Before)))
        {
            Apply(wb, sheetEvent);
            await Assert.That(Lines(Model(wb))).IsEqualTo(Lines(excel));
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(Lines(Model(reloaded))).IsEqualTo(Lines(excel));
    }

    /// <summary>
    /// <see cref="XLWorkbook.Save()"/> reopens the package the previous save wrote, whose rules the
    /// first save has already rewritten. A second save leaves them as the first one did.
    /// </summary>
    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    public async Task A_second_save_matches_Excel(SheetEvent sheetEvent)
    {
        using var package = new MemoryStream();
        using (var source = Resource(Before))
            source.CopyTo(package);

        package.Position = 0;
        using (var wb = new XLWorkbook(package))
        {
            Apply(wb, sheetEvent);
            wb.Save();
            wb.Save();
        }

        await AssertSameText(Read(package), Read(Resource(After(sheetEvent))));
    }

    /// <summary>The rules and names an edit left survive a reload and a second save as they are.</summary>
    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.WorksheetDelete)]
    public async Task A_reload_and_a_second_save_match_Excel(SheetEvent sheetEvent)
    {
        using var saved = EditAndSave(sheetEvent);
        using var resaved = new MemoryStream();
        using (var reloaded = new XLWorkbook(saved))
            reloaded.SaveAs(resaved);

        await AssertSameText(Read(resaved), Read(Resource(After(sheetEvent))));
    }

    /// <summary>
    /// After the delete, Excel keeps <c>B1</c> to <c>B4</c>, now <c>#REF!</c>, in the <c>x14</c>
    /// extension. XLibur writes them in the standard form. The text is the same; only the XML form
    /// differs. The writer puts a rule in the extension only when a formula is a range address on
    /// another sheet (<c>DataValidationWriter.UsesExternalSheet</c>), and <c>#REF!</c> is not one.
    /// </summary>
    /// <remarks>
    /// The same test sends <c>B2</c> and <c>B3</c>, an <c>OFFSET</c> and a custom formula, to the
    /// standard form on every save, a rename or a plain round trip included, where Excel keeps them in
    /// the extension. Which form a rule is written in belongs to spec 44 (data validation mapping), so
    /// this pins what XLibur writes today and the writer is left alone.
    /// </remarks>
    [Test]
    public async Task Known_difference_a_deleted_sheets_rules_are_written_in_the_standard_form()
    {
        await Assert.That(Lines(Forms(Resource("dv-delete-after.xlsx")))).IsEqualTo(Lines([
            "B1 x14", "B2 x14", "B3 x14", "B4 x14", "B5 standard", "B6 standard", "B7 standard",
        ]));

        using var saved = EditAndSave(SheetEvent.WorksheetDelete);

        await Assert.That(Lines(Forms(saved))).IsEqualTo(Lines([
            "B1 standard", "B2 standard", "B3 standard", "B4 standard", "B5 standard", "B6 standard", "B7 standard",
        ]));
    }

    private static string After(SheetEvent sheetEvent)
        => sheetEvent == SheetEvent.Rename ? "dv-rename-after.xlsx" : "dv-delete-after.xlsx";

    private static void Apply(XLWorkbook wb, SheetEvent sheetEvent)
    {
        switch (sheetEvent)
        {
            case SheetEvent.Rename:
                wb.Worksheet("Data").Name = "New Data";
                break;
            case SheetEvent.WorksheetDelete:
                wb.Worksheet("Data").Delete();
                break;
            case SheetEvent.CollectionDelete:
                wb.Worksheets.Delete("Data");
                break;
        }
    }

    private static MemoryStream EditAndSave(SheetEvent sheetEvent)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(Before)))
        {
            Apply(wb, sheetEvent);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return ms;
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));

    private static async Task AssertSameText(Holders saved, Holders excel)
    {
        await Assert.That(Lines(saved.Validations)).IsEqualTo(Lines(excel.Validations));
        await Assert.That(Lines(saved.Names)).IsEqualTo(Lines(excel.Names));
    }

    private static string Lines(IEnumerable<string> items) => string.Join(Environment.NewLine, items);

    /// <summary>A criterion as the rewrite reads it, without a leading <c>=</c>.</summary>
    private static string Criterion(string? text) => text is ['=', .. var formula] ? formula : text ?? string.Empty;

    /// <summary>
    /// A <c>sqref</c> with each one-cell range written as its cell. XLibur writes <c>B1:B1</c> where
    /// Excel writes <c>B1</c>, on any save, and the two mean the same cell.
    /// </summary>
    private static string Sqref(string? sqref)
        => string.Join(" ", (sqref ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(area => area.Split(':') is [var first, var last] && first == last ? first : area));

    /// <summary>
    /// The data validations on <c>Other</c>, one line per rule keyed by its <c>sqref</c>, in either
    /// form, and every defined name keyed by its scope, a sheet name or <see cref="Workbook"/>.
    /// </summary>
    private static Holders Read(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        var sheets = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().ToList();

        var names = (workbookPart.Workbook.DefinedNames?.Elements<S.DefinedName>() ?? [])
            .Select(n =>
            {
                var scope = n.LocalSheetId?.Value is { } id ? sheets[(int)id].Name!.Value! : Workbook;
                return $"{scope}|{n.Name!.Value} = {n.Text}";
            })
            .Order(StringComparer.Ordinal)
            .ToList();

        var worksheet = OtherPart(workbookPart).Worksheet!;
        var standard = worksheet.Elements<S.DataValidations>()
            .SelectMany(d => d.Elements<S.DataValidation>())
            .Select(dv => Describe(dv.SequenceOfReferences?.InnerText, dv.Type?.InnerText, dv.Operator?.InnerText,
                dv.Formula1?.Text, dv.Formula2?.Text));
        var extension = worksheet.Descendants<X14.DataValidation>()
            .Select(dv => Describe(dv.ReferenceSequence?.Text, dv.Type?.InnerText, dv.Operator?.InnerText,
                dv.DataValidationForumla1?.InnerText, dv.DataValidationForumla2?.InnerText));
        var validations = standard.Concat(extension).Order(StringComparer.Ordinal).ToList();

        return new Holders(validations, names);

        // An absent type is "none" and an absent operator "between", the schema's defaults. Only the
        // types that compare values have an operator.
        static string Describe(string? sqref, string? type, string? op, string? formula1, string? formula2)
        {
            type ??= "none";
            var comparison = type is "whole" or "decimal" or "date" or "time" or "textLength"
                ? " " + (op ?? "between")
                : string.Empty;
            return $"{Sqref(sqref)}: {type}{comparison} | {Criterion(formula1)} | {Criterion(formula2)}";
        }
    }

    /// <summary>
    /// Each rule on <c>Other</c> by its <c>sqref</c>, with the form it is written in: <c>standard</c>
    /// in <c>&lt;dataValidations&gt;</c>, <c>x14</c> in the extension.
    /// </summary>
    private static List<string> Forms(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var worksheet = OtherPart(document.WorkbookPart!).Worksheet!;
        var standard = worksheet.Elements<S.DataValidations>()
            .SelectMany(d => d.Elements<S.DataValidation>())
            .Select(dv => $"{Sqref(dv.SequenceOfReferences?.InnerText)} standard");
        var extension = worksheet.Descendants<X14.DataValidation>()
            .Select(dv => $"{Sqref(dv.ReferenceSequence?.Text)} x14");
        return standard.Concat(extension).Order(StringComparer.Ordinal).ToList();
    }

    private static WorksheetPart OtherPart(WorkbookPart workbookPart)
    {
        var other = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().Single(s => s.Name == "Other");
        return (WorksheetPart)workbookPart.GetPartById(other.Id!.Value!);
    }

    /// <summary>
    /// Every rule of every sheet as XLibur holds it, keyed by sheet and range, and every defined name
    /// keyed by its scope.
    /// </summary>
    private static List<string> Model(XLWorkbook wb)
    {
        var validations = wb.Worksheets.SelectMany(ws => ws.DataValidations.Select(dv =>
        {
            var sqref = Sqref(string.Join(" ", dv.Ranges.Select(r => r.RangeAddress.ToString())));
            var comparison = dv.AllowedValues is XLAllowedValues.WholeNumber or XLAllowedValues.Decimal
                or XLAllowedValues.Date or XLAllowedValues.Time or XLAllowedValues.TextLength
                ? " " + dv.Operator
                : string.Empty;
            return $"{ws.Name}!{sqref}: {dv.AllowedValues}{comparison} | {Criterion(dv.MinValue)} | {Criterion(dv.MaxValue)}";
        }));
        var names = wb.DefinedNames.Select(n => $"{Workbook}|{n.Name} = {n.RefersTo}")
            .Concat(wb.Worksheets.SelectMany(ws => ws.DefinedNames.Select(n => $"{ws.Name}|{n.Name} = {n.RefersTo}")));

        return validations.Concat(names).Order(StringComparer.Ordinal).ToList();
    }

    private sealed record Holders(List<string> Validations, List<string> Names);
}
