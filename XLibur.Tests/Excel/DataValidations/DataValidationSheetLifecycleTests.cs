using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.DataValidations;

/// <summary>
/// A sheet rename or delete rewrites the data-validation criteria that refer to the sheet, as it
/// rewrites a cell formula, a defined name or a conditional format (spec 55, D63). The criteria are
/// <see cref="IXLDataValidation.MinValue"/> and <see cref="IXLDataValidation.MaxValue"/>.
/// </summary>
/// <remarks>
/// <para>
/// Each test renames <c>Data</c> to <c>New Data</c>, a name that needs quotes, or deletes it. The
/// rename is the one Excel was seen making in cell formulas, names, conditional formats and print
/// areas (the <c>rename-*</c> fixtures).
/// </para>
/// <para>
/// No Excel fixture holds a data validation yet. Every expectation for a delete, marked
/// <c>Pending dv-delete-after fixture</c>, is the form Excel was seen writing for a conditional format
/// (<c>#REF!&gt;0</c>, and <c>Sheet!#REF!</c> as a plain <c>#REF!</c>); the fixture's
/// <c>dv-delete-after.xlsx</c> confirms it or corrects it.
/// </para>
/// </remarks>
public class DataValidationSheetLifecycleTests
{
    /// <summary>What happens to the sheet <c>Data</c>.</summary>
    public enum SheetEvent
    {
        /// <summary>Renamed to <c>New Data</c>.</summary>
        Rename,

        /// <summary>Deleted.</summary>
        Delete,
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.Delete)]
    public async Task A_list_on_another_sheet_follows(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Range("B1:B3").CreateDataValidation();
        rule.List("=Data!$A$1:$A$3");

        Apply(wb, sheetEvent);

        var expected = Expect(sheetEvent,
            rename: "='New Data'!$A$1:$A$3",
            delete: "=#REF!"); // Pending dv-delete-after fixture
        await Assert.That(rule.MinValue).IsEqualTo(expected);
        await AssertSavedAndReloaded(wb, "Other", ("standard", expected, ""));
    }

    /// <summary>
    /// <c>List(IXLRange)</c> stores the range with its sheet and no <c>=</c>, the form a loaded rule
    /// has. On another sheet it is written in the <c>x14</c> extension, until the delete leaves it no
    /// reference to another sheet.
    /// </summary>
    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.Delete)]
    public async Task A_list_from_a_range_on_another_sheet_follows(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out var data, out var other);
        var rule = other.Range("B1:B3").CreateDataValidation();
        rule.List(data.Range("A1:A3"));

        Apply(wb, sheetEvent);

        var expected = Expect(sheetEvent,
            rename: "'New Data'!$A$1:$A$3",
            delete: "#REF!"); // Pending dv-delete-after fixture
        await Assert.That(rule.MinValue).IsEqualTo(expected);
        await AssertSavedAndReloaded(wb, "Other", (Expect(sheetEvent, "x14", "standard"), expected, ""));
    }

    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.Delete)]
    public async Task An_OFFSET_list_source_follows(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Range("B1:B3").CreateDataValidation();
        rule.List("=OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)");

        Apply(wb, sheetEvent);

        var expected = Expect(sheetEvent,
            rename: "=OFFSET('New Data'!$A$1,0,0,COUNTA('New Data'!$A:$A),1)",
            delete: "=OFFSET(#REF!,0,0,COUNTA(#REF!),1)"); // Pending dv-delete-after fixture
        await Assert.That(rule.MinValue).IsEqualTo(expected);
        await AssertSavedAndReloaded(wb, "Other", ("standard", expected, ""));
    }

    /// <summary>The relative reference is to the rule's own sheet, and keeps its text.</summary>
    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.Delete)]
    public async Task A_custom_formula_follows_and_keeps_its_relative_reference(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Range("B1:B3").CreateDataValidation();
        rule.Custom("=AND(B1>0,B1<=Data!$A$1)");

        Apply(wb, sheetEvent);

        var expected = Expect(sheetEvent,
            rename: "=AND(B1>0,B1<='New Data'!$A$1)",
            delete: "=AND(B1>0,B1<=#REF!)"); // Pending dv-delete-after fixture
        await Assert.That(rule.Value).IsEqualTo(expected);
        await AssertSavedAndReloaded(wb, "Other", ("standard", expected, ""));
    }

    /// <summary>
    /// Both ends of a between rule, in the form a loaded rule has. Each names another sheet, so the
    /// rule is written in the <c>x14</c> extension until the delete.
    /// </summary>
    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.Delete)]
    public async Task Both_ends_of_a_between_rule_follow(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Range("B1:B3").CreateDataValidation();
        rule.WholeNumber.Between("Data!$A$2", "Data!$A$3");

        Apply(wb, sheetEvent);

        var min = Expect(sheetEvent, rename: "'New Data'!$A$2", delete: "#REF!"); // Pending dv-delete-after fixture
        var max = Expect(sheetEvent, rename: "'New Data'!$A$3", delete: "#REF!"); // Pending dv-delete-after fixture
        await Assert.That(rule.MinValue).IsEqualTo(min);
        await Assert.That(rule.MaxValue).IsEqualTo(max);
        await AssertSavedAndReloaded(wb, "Other", (Expect(sheetEvent, "x14", "standard"), min, max));
    }

    /// <summary>
    /// The criterion names the workbook-scoped name, not the sheet, so its text stays. The name itself
    /// follows the rename or the delete.
    /// </summary>
    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.Delete)]
    public async Task A_list_through_a_workbook_scoped_name_keeps_its_text(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        wb.DefinedNames.Add("Items", "Data!$A$1:$A$3");
        var rule = other.Range("B1:B3").CreateDataValidation();
        rule.List("=Items");

        Apply(wb, sheetEvent);

        await Assert.That(rule.MinValue).IsEqualTo("=Items");
        await Assert.That(wb.DefinedNames.Single().RefersTo).IsEqualTo(Expect(sheetEvent,
            rename: "'New Data'!$A$1:$A$3",
            delete: "#REF!"));
        await AssertSavedAndReloaded(wb, "Other", ("standard", "=Items", ""));
    }

    /// <summary>
    /// A literal list is text, and never changes, even when an item is the sheet's name.
    /// </summary>
    [Test]
    [Arguments(SheetEvent.Rename, "x,y,z")]
    [Arguments(SheetEvent.Delete, "x,y,z")]
    [Arguments(SheetEvent.Rename, "Data,y,z")]
    [Arguments(SheetEvent.Delete, "Data,y,z")]
    public async Task A_literal_list_never_changes(SheetEvent sheetEvent, string items)
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Range("B1:B3").CreateDataValidation();
        rule.List(items);
        var literal = "\"" + items + "\"";
        await Assert.That(rule.MinValue).IsEqualTo(literal);

        Apply(wb, sheetEvent);

        await Assert.That(rule.MinValue).IsEqualTo(literal);
        await AssertSavedAndReloaded(wb, "Other", ("standard", literal, ""));
    }

    /// <summary>
    /// A rule that names its own sheet follows that sheet's rename. A delete of its own sheet takes
    /// the rule with it.
    /// </summary>
    [Test]
    public async Task A_list_on_its_own_sheet_follows_a_rename_of_that_sheet()
    {
        using var wb = NewBook(out var data, out _);
        var rule = data.Range("C1:C3").CreateDataValidation();
        rule.List("=Data!$A$1:$A$3");

        Apply(wb, SheetEvent.Rename);

        await Assert.That(rule.MinValue).IsEqualTo("='New Data'!$A$1:$A$3");
        await AssertSavedAndReloaded(wb, "New Data", ("standard", "='New Data'!$A$1:$A$3", ""));
    }

    /// <summary>
    /// <c>List(IXLRange)</c> on the rule's own sheet stores the sheet's name. After the sheet was
    /// renamed, the save no longer saw that name as the rule's own sheet, and wrote the rule in the
    /// <c>x14</c> extension as a reference to <c>Data</c>, a sheet the workbook no longer has. It is
    /// now written in the standard form, as a reference on the rule's own sheet.
    /// </summary>
    [Test]
    public async Task A_list_from_a_range_on_its_own_sheet_is_saved_as_its_own_after_a_rename()
    {
        using var wb = NewBook(out var data, out _);
        var rule = data.Range("C1:C3").CreateDataValidation();
        rule.List(data.Range("A1:A3"));
        await Assert.That(rule.MinValue).IsEqualTo("Data!$A$1:$A$3");

        Apply(wb, SheetEvent.Rename);

        await Assert.That(rule.MinValue).IsEqualTo("'New Data'!$A$1:$A$3");
        await AssertSavedAndReloaded(wb, "New Data", ("standard", "$A$1:$A$3", ""));
    }

    /// <summary>
    /// A rule the <c>x14</c> extension held when the workbook was loaded is in the same model as any
    /// other, and follows a rename or a delete as one does.
    /// </summary>
    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.Delete)]
    public async Task A_rule_loaded_from_the_x14_extension_follows(SheetEvent sheetEvent)
    {
        using var original = new MemoryStream();
        using (var wb = NewBook(out var data, out var other))
        {
            other.Range("B1:B3").CreateDataValidation().List(data.Range("A1:A3"));
            wb.SaveAs(original);
        }

        await Assert.That(SavedCriteria(original, "Other"))
            .IsEquivalentTo(new[] { ("x14", "Data!$A$1:$A$3", "") }, CollectionOrdering.Matching);

        using var loaded = new XLWorkbook(original);
        Apply(loaded, sheetEvent);

        var expected = Expect(sheetEvent,
            rename: "'New Data'!$A$1:$A$3",
            delete: "#REF!"); // Pending dv-delete-after fixture
        await Assert.That(loaded.Worksheet("Other").DataValidations.Single().MinValue).IsEqualTo(expected);
        await AssertSavedAndReloaded(loaded, "Other", (Expect(sheetEvent, "x14", "standard"), expected, ""));
    }

    /// <summary>A formula the parser refuses keeps its text (ADR 0002).</summary>
    [Test]
    [Arguments(SheetEvent.Rename)]
    [Arguments(SheetEvent.Delete)]
    public async Task A_formula_the_parser_refuses_keeps_its_text(SheetEvent sheetEvent)
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Range("B1:B3").CreateDataValidation();
        rule.Custom("=SUM(Data!A1");

        Apply(wb, sheetEvent);

        await Assert.That(rule.Value).IsEqualTo("=SUM(Data!A1");
    }

    /// <summary>
    /// Pending dv-delete-after fixture: the recipe's <c>=Data!DvOnly</c> shows whether a validation
    /// keeps a name scoped to the deleted sheet alive. Spec 55 keeps such a name only when a cell
    /// formula on another sheet or a defined name refers to it (<c>FindNamesOutlivingSheet</c>), and
    /// a validation is neither. So the name goes with the sheet, and the criterion reads <c>#REF!</c>,
    /// as a conditional format's does in the same case.
    /// </summary>
    [Test]
    public async Task Pending_fixture_a_name_scoped_to_the_deleted_sheet_that_only_a_validation_uses_goes_with_it()
    {
        using var wb = NewBook(out var data, out var other);
        data.DefinedNames.Add("DvOnly", "Data!$A$1:$A$3");
        var rule = other.Range("B1:B3").CreateDataValidation();
        rule.List("=Data!DvOnly");

        data.Delete();

        await Assert.That(wb.DefinedNames.Any(n => n.Name == "DvOnly")).IsFalse();
        await Assert.That(rule.MinValue).IsEqualTo("=#REF!");
    }

    private static XLWorkbook NewBook(out IXLWorksheet data, out IXLWorksheet other)
    {
        var wb = new XLWorkbook();
        data = wb.AddWorksheet("Data");
        other = wb.AddWorksheet("Other");
        data.Cell("A1").Value = 5;
        data.Cell("A2").Value = 2;
        data.Cell("A3").Value = 3;
        return wb;
    }

    private static void Apply(XLWorkbook wb, SheetEvent sheetEvent)
    {
        if (sheetEvent == SheetEvent.Rename)
            wb.Worksheet("Data").Name = "New Data";
        else
            wb.Worksheet("Data").Delete();
    }

    private static T Expect<T>(SheetEvent sheetEvent, T rename, T delete)
        => sheetEvent == SheetEvent.Rename ? rename : delete;

    /// <summary>
    /// Saves <paramref name="wb"/> and checks, on <paramref name="sheetName"/>, that each rule is
    /// written in the form and with the criteria given, that nothing in the sheet's part names the
    /// sheet <c>Data</c>, which the workbook no longer has, and that a reload reads each rule's
    /// criteria back as they were written.
    /// </summary>
    /// <remarks>
    /// <c>'New Data'!</c> does not contain <c>Data!</c>, so the check holds for the rename too.
    /// </remarks>
    private static async Task AssertSavedAndReloaded(XLWorkbook wb, string sheetName,
        params (string Form, string Formula1, string Formula2)[] expected)
    {
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        await Assert.That(SavedCriteria(ms, sheetName)).IsEquivalentTo(expected, CollectionOrdering.Matching);
        await Assert.That(SavedSheetXml(ms, sheetName)).DoesNotContain("Data!");

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var criteria = reloaded.Worksheet(sheetName).DataValidations.Select(dv => (dv.MinValue, dv.MaxValue));
        await Assert.That(criteria)
            .IsEquivalentTo(expected.Select(e => (e.Formula1, e.Formula2)), CollectionOrdering.Matching);
    }

    /// <summary>
    /// Each rule in <paramref name="sheetName"/>'s part: <c>standard</c> for one in
    /// <c>&lt;dataValidations&gt;</c>, <c>x14</c> for one in the extension, with its two formulas, an
    /// absent one read as empty.
    /// </summary>
    private static List<(string Form, string Formula1, string Formula2)> SavedCriteria(Stream package,
        string sheetName)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var worksheet = SheetPart(document, sheetName).Worksheet!;
        var standard = worksheet.Elements<S.DataValidations>()
            .SelectMany(d => d.Elements<S.DataValidation>())
            .Select(dv => ("standard", dv.Formula1?.Text ?? "", dv.Formula2?.Text ?? ""));
        var extension = worksheet.Descendants<X14.DataValidation>()
            .Select(dv => ("x14", dv.DataValidationForumla1?.InnerText ?? "", dv.DataValidationForumla2?.InnerText ?? ""));
        return standard.Concat(extension).ToList();
    }

    private static string SavedSheetXml(Stream package, string sheetName)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        using var reader = new StreamReader(SheetPart(document, sheetName).GetStream(FileMode.Open, FileAccess.Read));
        return reader.ReadToEnd();
    }

    private static WorksheetPart SheetPart(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().Single(s => s.Name?.Value == sheetName);
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }
}
