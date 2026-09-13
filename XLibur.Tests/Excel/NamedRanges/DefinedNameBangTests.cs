using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Excel.IO;

namespace XLibur.Tests.Excel.NamedRanges;

/// <summary>
/// A bang reference (<c>!$A$1</c>) and a bang name (<c>!Total</c>) in a defined name mean "on the
/// sheet the formula is evaluated for" — the sheet of the cell that uses the name, not the sheet the
/// name belongs to. That is what makes <c>!$A$1</c> the form Excel expects in place of the sheet-less
/// <c>$A$1</c> it refuses in a defined name. See https://github.com/XLibur/XLibur/issues/446.
/// </summary>
/// <remarks>
/// Every reference here is absolute. A relative reference in a defined name should be offset by the
/// cell that uses it, and XLibur does not do that for any defined name yet, bang or not.
/// </remarks>
public class DefinedNameBangTests
{
    [Test]
    public async Task A_bang_reference_resolves_against_the_sheet_of_the_cell_using_the_name()
    {
        using var wb = new XLWorkbook();
        var first = wb.AddWorksheet("First");
        var second = wb.AddWorksheet("Second");
        first.Cell("A1").Value = 1;
        second.Cell("A1").Value = 2;
        wb.DefinedNames.Add("Here", "!$A$1");
        first.Cell("C1").FormulaA1 = "Here";
        second.Cell("C1").FormulaA1 = "Here";

        await Assert.That(first.Cell("C1").Value).IsEqualTo(1);
        await Assert.That(second.Cell("C1").Value).IsEqualTo(2);
    }

    [Test]
    public async Task A_bang_name_resolves_the_sheet_scoped_name_of_the_sheet_using_it()
    {
        using var wb = new XLWorkbook();
        var first = wb.AddWorksheet("First");
        var second = wb.AddWorksheet("Second");
        first.Cell("A1").Value = 7;
        second.Cell("A1").Value = 9;
        first.DefinedNames.Add("Total", "First!$A$1");
        second.DefinedNames.Add("Total", "Second!$A$1");
        wb.DefinedNames.Add("LocalTotal", "!Total");
        first.Cell("C1").FormulaA1 = "LocalTotal";
        second.Cell("C1").FormulaA1 = "LocalTotal";

        await Assert.That(first.Cell("C1").Value).IsEqualTo(7);
        await Assert.That(second.Cell("C1").Value).IsEqualTo(9);
    }

    /// <summary>
    /// Excel shadows rather than isolates: a sheet-scoped name wins only on a sheet that has one.
    /// Elsewhere the workbook-scoped name of the same identifier answers, not <c>#NAME?</c>.
    /// </summary>
    [Test]
    public async Task A_bang_name_falls_back_to_the_workbook_scoped_name_where_the_sheet_has_none()
    {
        using var wb = new XLWorkbook();
        var first = wb.AddWorksheet("First");
        var second = wb.AddWorksheet("Second");
        first.Cell("A1").Value = 7;
        first.Cell("B1").Value = 5;
        first.DefinedNames.Add("Total", "First!$A$1");
        wb.DefinedNames.Add("Total", "First!$B$1");
        wb.DefinedNames.Add("LocalTotal", "!Total");
        first.Cell("C1").FormulaA1 = "LocalTotal";
        second.Cell("C1").FormulaA1 = "LocalTotal";

        await Assert.That(first.Cell("C1").Value).IsEqualTo(7);
        await Assert.That(second.Cell("C1").Value).IsEqualTo(5);
    }

    [Test]
    public async Task A_bang_name_that_names_nothing_is_a_name_error()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("First");
        wb.DefinedNames.Add("LocalTotal", "!Total");
        ws.Cell("C1").FormulaA1 = "LocalTotal";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NameNotRecognized);
    }

    /// <summary>
    /// The precedent has to be registered on the sheet of the cell using the name. Registered on any
    /// other sheet — or not at all — the edit would leave the cell holding its first answer.
    /// </summary>
    [Test]
    public async Task A_cell_using_a_bang_reference_recalculates_when_the_referenced_cell_changes()
    {
        using var wb = new XLWorkbook();
        var first = wb.AddWorksheet("First");
        var second = wb.AddWorksheet("Second");
        first.Cell("A1").Value = 1;
        second.Cell("A1").Value = 2;
        wb.DefinedNames.Add("Here", "!$A$1");
        second.Cell("C1").FormulaA1 = "Here";
        await Assert.That(second.Cell("C1").Value).IsEqualTo(2);

        second.Cell("A1").Value = 20;

        await Assert.That(second.Cell("C1").Value).IsEqualTo(20);
    }

    [Test]
    public async Task A_cell_using_a_bang_name_recalculates_when_the_named_cell_changes()
    {
        using var wb = new XLWorkbook();
        var first = wb.AddWorksheet("First");
        var second = wb.AddWorksheet("Second");
        first.Cell("A1").Value = 7;
        second.Cell("A1").Value = 9;
        first.DefinedNames.Add("Total", "First!$A$1");
        second.DefinedNames.Add("Total", "Second!$A$1");
        wb.DefinedNames.Add("LocalTotal", "!Total");
        second.Cell("C1").FormulaA1 = "LocalTotal";
        await Assert.That(second.Cell("C1").Value).IsEqualTo(9);

        second.Cell("A1").Value = 90;

        await Assert.That(second.Cell("C1").Value).IsEqualTo(90);
    }

    [Test]
    public async Task A_bang_name_loaded_from_a_file_evaluates()
    {
        using var package = BookWithBangName();
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Sheet1");
        ws.Cell("C1").FormulaA1 = "x";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(7);
    }

    [Test]
    public async Task A_bang_name_loaded_from_a_file_keeps_its_text_on_save()
    {
        using var package = BookWithBangName();
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(package))
            wb.SaveAs(saved);

        using var reloaded = new XLWorkbook(saved);

        await Assert.That(reloaded.DefinedNames.Single(dn => dn.Name == "x").RefersTo).IsEqualTo("!Total");
    }

    /// <summary>
    /// A package whose <c>Sheet1!A1</c> holds 7, with a name <c>Total</c> scoped to <c>Sheet1</c>
    /// pointing at it and a workbook-scoped name <c>x</c> whose text is exactly <c>!Total</c> — written
    /// into the workbook part directly, so the reader sees the text as Excel would store it.
    /// </summary>
    private static MemoryStream BookWithBangName()
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1").Cell("A1").Value = 7;
            wb.SaveAs(package);
        }

        return package.RewriteWorkbook(xml =>
        {
            const string definedNames =
                "<x:definedNames>" +
                "<x:definedName name=\"Total\" localSheetId=\"0\">Sheet1!$A$1</x:definedName>" +
                "<x:definedName name=\"x\">!Total</x:definedName>" +
                "</x:definedNames>";
            var rewritten = xml.Replace("<x:definedNames />", definedNames);
            if (ReferenceEquals(rewritten, xml) || !rewritten.Contains("definedName name=\"x\"", StringComparison.Ordinal))
                throw new InvalidOperationException("The defined names were not spliced into the workbook part.");

            return rewritten;
        });
    }
}
