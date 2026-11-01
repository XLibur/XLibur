using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
using XLibur.Excel;

namespace XLibur.Tests.Excel.DataValidations;

/// <summary>
/// The text of a rule's criteria, <see cref="IXLDataValidation.MinValue"/> and
/// <see cref="IXLDataValidation.MaxValue"/>, as the API stores it and as a save writes it.
/// </summary>
/// <remarks>
/// Each "as Excel wrote" expectation is from a workbook Excel 16 made through COM (2026-09-15), with one
/// rule in each cell of <c>Other!B1:B11</c> and the sheets <c>Data</c> and <c>My Data</c> beside it. A
/// file stores a criterion without a leading <c>=</c>. A reference to another sheet is written in the
/// sheet's <c>x14</c> extension, with the sheet's name, quoted where the name needs it: <c>Data!$A$1</c>,
/// <c>'My Data'!$A$1</c>. A reference to the rule's own sheet is written in the standard element without
/// a sheet name, even when it was typed with one: <c>=Other!$D$1:$D$3</c> was written <c>$D$1:$D$3</c>.
/// </remarks>
public class DataValidationCriteriaTests
{
    /// <summary>The criteria that compare a value, each of which can take it from a cell.</summary>
    public enum Criteria
    {
        WholeNumber,
        Decimal,
        Date,
        Time,
        TextLength,
    }

    /// <summary>
    /// A criterion set from a cell on another sheet keeps that sheet's name. It was stored as the address
    /// alone, which Excel reads as a cell on the rule's own sheet, so the rule compared with a different
    /// cell (#524).
    /// </summary>
    [Test]
    [MatrixDataSource]
    public async Task A_cell_on_another_sheet_keeps_its_sheet(
        [Matrix(Criteria.WholeNumber, Criteria.Decimal, Criteria.Date, Criteria.Time, Criteria.TextLength)]
        Criteria criteria,
        [Matrix(XLOperator.Between, XLOperator.NotBetween, XLOperator.EqualTo, XLOperator.NotEqualTo,
            XLOperator.GreaterThan, XLOperator.LessThan, XLOperator.EqualOrGreaterThan, XLOperator.EqualOrLessThan)]
        XLOperator comparison)
    {
        using var wb = NewBook(out var data, out var other);
        var rule = other.Cell("B1").CreateDataValidation();

        Compare(Of(rule, criteria), comparison, data.Cell("A1"), data.Cell("A2"));

        await Assert.That(rule.Operator).IsEqualTo(comparison);
        await Assert.That(rule.MinValue).IsEqualTo("Data!$A$1");
        await Assert.That(rule.MaxValue).IsEqualTo(TakesTwoValues(comparison) ? "Data!$A$2" : "");
    }

    /// <summary>
    /// A date criterion is stored as the date's serial number, in invariant text whatever the culture,
    /// through the same one-value and two-value paths as every other criteria type.
    /// </summary>
    [Test]
    [SetCulture("cs-CZ")]
    [MatrixDataSource]
    public async Task A_date_criterion_is_stored_as_its_serial_number(
        [Matrix(XLOperator.Between, XLOperator.NotBetween, XLOperator.EqualTo, XLOperator.NotEqualTo,
            XLOperator.GreaterThan, XLOperator.LessThan, XLOperator.EqualOrGreaterThan, XLOperator.EqualOrLessThan)]
        XLOperator comparison)
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Cell("B1").CreateDataValidation();

        CompareDates(rule.Date, comparison, new DateTime(2024, 1, 15, 12, 0, 0), new DateTime(2024, 2, 1));

        await Assert.That(rule.AllowedValues).IsEqualTo(XLAllowedValues.Date);
        await Assert.That(rule.Operator).IsEqualTo(comparison);
        await Assert.That(rule.MinValue).IsEqualTo("45306.5");
        await Assert.That(rule.MaxValue).IsEqualTo(TakesTwoValues(comparison) ? "45323" : "");
    }

    /// <summary>The sheet's name is written as a formula writes it, quoted where it needs to be.</summary>
    [Test]
    [Arguments("My Data", "'My Data'")]
    [Arguments("Bob's", "'Bob''s'")]
    [Arguments("2024", "'2024'")]
    public async Task A_sheet_name_that_needs_quotes_is_quoted(string sheetName, string reference)
    {
        using var wb = NewBook(out _, out var other);
        var source = wb.AddWorksheet(sheetName);
        var rule = other.Cell("B1").CreateDataValidation();

        rule.WholeNumber.Between(source.Cell("A1"), source.Cell("A2"));

        await Assert.That(rule.MinValue).IsEqualTo(reference + "!$A$1");
        await Assert.That(rule.MaxValue).IsEqualTo(reference + "!$A$2");
    }

    /// <summary>
    /// A cell on the rule's own sheet is stored without the sheet's name, as before, and as Excel writes
    /// such a reference. As nothing names the sheet, the rule still refers to its own sheet's cells after
    /// that sheet is renamed, and a copy of the sheet refers to the copy's cells, not the original's.
    /// </summary>
    [Test]
    public async Task A_cell_on_the_rules_own_sheet_is_stored_without_a_sheet_name()
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Cell("B1").CreateDataValidation();

        rule.WholeNumber.Between(other.Cell("D1"), other.Cell("D2"));

        await Assert.That(rule.MinValue).IsEqualTo("$D$1");
        await Assert.That(rule.MaxValue).IsEqualTo("$D$2");

        other.Name = "Renamed";
        var copy = other.CopyTo("Copy");

        await Assert.That(rule.MinValue).IsEqualTo("$D$1");
        await Assert.That(copy.Cell("B1").GetDataValidation().MinValue).IsEqualTo("$D$1");
        await Assert.That(copy.Cell("B1").GetDataValidation().MaxValue).IsEqualTo("$D$2");
    }

    /// <summary>
    /// The address is A1 text, as a formula in a file is, also in a workbook whose reference style is
    /// R1C1. It followed the workbook's style, and was stored as <c>R1C4</c>.
    /// </summary>
    [Test]
    public async Task A_cell_is_stored_as_an_A1_address_whatever_the_workbooks_reference_style()
    {
        using var wb = NewBook(out var data, out var other);
        wb.ReferenceStyle = XLReferenceStyle.R1C1;
        var rule = other.Cell("B1").CreateDataValidation();

        rule.WholeNumber.Between(other.Cell("D1"), data.Cell("A2"));

        await Assert.That(rule.MinValue).IsEqualTo("$D$1");
        await Assert.That(rule.MaxValue).IsEqualTo("Data!$A$2");
    }

    /// <summary>
    /// The sheet's name in the criterion is what lets a rename or a delete of that sheet reach it (D63).
    /// </summary>
    [Test]
    [Arguments(true, "'New Data'!$A$1", "'New Data'!$A$2")]
    [Arguments(false, "#REF!", "#REF!")] // As Excel wrote B4 in dv-delete-after
    public async Task A_cell_on_another_sheet_follows_a_rename_or_a_delete_of_that_sheet(bool rename,
        string min, string max)
    {
        using var wb = NewBook(out var data, out var other);
        var rule = other.Cell("B1").CreateDataValidation();
        rule.WholeNumber.Between(data.Cell("A1"), data.Cell("A2"));

        if (rename)
            data.Name = "New Data";
        else
            data.Delete();

        await Assert.That(rule.MinValue).IsEqualTo(min);
        await Assert.That(rule.MaxValue).IsEqualTo(max);
    }

    /// <summary>
    /// A between rule over cells is written in the form Excel wrote the same rule in, and a reload reads
    /// its criteria back as they were written.
    /// </summary>
    [Test]
    [Arguments("Data", "x14", "Data!$A$1", "Data!$A$2")] // As Excel wrote B3
    [Arguments("My Data", "x14", "'My Data'!$A$1", "'My Data'!$A$2")] // As Excel wrote B9
    [Arguments("Other", "standard", "$A$1", "$A$2")] // As Excel wrote B7, over $D$1 and $D$2
    public async Task A_between_rule_over_cells_is_saved_as_Excel_saves_it(string sheetName, string form,
        string formula1, string formula2)
    {
        using var wb = NewBook(out _, out var other);
        var source = wb.Worksheets.Contains(sheetName) ? wb.Worksheet(sheetName) : wb.AddWorksheet(sheetName);

        other.Cell("B1").CreateDataValidation().WholeNumber.Between(source.Cell("A1"), source.Cell("A2"));

        await AssertSavedAndReloaded(wb, form, formula1, formula2);
    }

    /// <summary>
    /// A list set with a leading <c>=</c>, as a formula is typed, is saved as Excel saves the same rule:
    /// without the <c>=</c>, and in the <c>x14</c> extension when it refers to another sheet. The <c>=</c>
    /// was written into <c>&lt;formula1&gt;</c>, and the rule went to the standard form whatever it
    /// referred to, because the check for another sheet allows no <c>=</c> (#523). The rule keeps its
    /// text, <c>=</c> included, until it is saved; a reload reads the text the file holds.
    /// </summary>
    [Test]
    [Arguments("=Data!$A$1:$A$3", "x14", "Data!$A$1:$A$3")] // As Excel wrote B1
    [Arguments("='My Data'!$A$1:$A$3", "x14", "'My Data'!$A$1:$A$3")] // As Excel wrote B8
    [Arguments("=$D$1:$D$3", "standard", "$D$1:$D$3")] // As Excel wrote B5
    [Arguments("=Other!$D$1:$D$3", "standard", "$D$1:$D$3")] // As Excel wrote B10
    public async Task A_list_set_with_a_leading_equals_is_saved_as_Excel_saves_it(string list, string form,
        string formula1)
    {
        using var wb = NewBook(out _, out var other);
        wb.AddWorksheet("My Data");
        var rule = other.Cell("B1").CreateDataValidation();

        rule.List(list);

        await Assert.That(rule.MinValue).IsEqualTo(list);
        await AssertSavedAndReloaded(wb, form, formula1, "");
    }

    /// <summary>
    /// Both ends of a between rule set as text with a leading <c>=</c> are saved as Excel saves the same
    /// rule.
    /// </summary>
    [Test]
    [Arguments("=Data!$A$1", "=Data!$A$2", "x14", "Data!$A$1", "Data!$A$2")] // As Excel wrote B3
    [Arguments("=$D$1", "=$D$2", "standard", "$D$1", "$D$2")] // As Excel wrote B7
    public async Task A_between_rule_set_with_a_leading_equals_is_saved_as_Excel_saves_it(string min, string max,
        string form, string formula1, string formula2)
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Cell("B1").CreateDataValidation();

        rule.WholeNumber.Between(min, max);

        await Assert.That(rule.MinValue).IsEqualTo(min);
        await Assert.That(rule.MaxValue).IsEqualTo(max);
        await AssertSavedAndReloaded(wb, form, formula1, formula2);
    }

    /// <summary>
    /// A custom formula that refers to another sheet is saved as Excel saved <c>B6</c>: without the
    /// leading <c>=</c>, and in the <c>x14</c> extension. The rule was written in the standard form,
    /// because only a criterion that was a range address on another sheet went to the extension (#536).
    /// </summary>
    [Test]
    public async Task A_custom_formula_that_refers_to_another_sheet_is_saved_as_Excel_saves_it()
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Cell("B6").CreateDataValidation();

        rule.Custom("=B6<=MAX(Data!$A$1:$A$3)");

        await AssertSavedAndReloaded(wb, "x14", "B6<=MAX(Data!$A$1:$A$3)", ""); // As Excel wrote B6
    }

    /// <summary>
    /// A rule goes to the <c>x14</c> extension when a reference anywhere in its formula names another
    /// sheet (#536). A reference to the rule's own sheet, written with the sheet's name or without it, a
    /// defined name, and text in a string name no other sheet, so such a rule stays in the standard form.
    /// </summary>
    /// <remarks>
    /// <c>dv-before</c> and <c>dv-copy-after</c> are in <c>Resource/Other/SheetLifecycle</c>. Excel drops
    /// the name of the rule's own sheet from the formula. XLibur keeps it, so only the form is Excel's for
    /// <c>OFFSET(Other!$D$1,…)</c>.
    /// </remarks>
    [Test]
    [Arguments("=OFFSET(Data!$A$1,0,0,3,1)", "x14")] // As Excel wrote B2 in dv-before
    [Arguments("=OFFSET('My Data'!$A$1,0,0,3,1)", "x14")]
    [Arguments("=OFFSET($D$1,0,0,3,1)", "standard")] // As Excel wrote B5 in dv-copy-after
    [Arguments("=OFFSET(Other!$D$1,0,0,3,1)", "standard")] // As Excel wrote B5 in dv-copy-after
    [Arguments("=Items", "standard")] // As Excel wrote B5 in dv-before, a list through ListW
    [Arguments("=INDIRECT(\"Data!$A$1:$A$3\")", "standard")] // As Excel wrote B6 in dv-before
    public async Task A_list_through_a_formula_is_saved_in_x14_when_it_refers_to_another_sheet(string list,
        string form)
    {
        using var wb = NewBook(out _, out var other);
        wb.AddWorksheet("My Data");
        wb.DefinedNames.Add("Items", "Data!$A$1:$A$3");
        var rule = other.Cell("B1").CreateDataValidation();

        rule.List(list);

        await AssertSavedAndReloaded(wb, form, SavedDataValidations.AsSaved(list), "");
    }

    /// <summary>
    /// A custom formula goes to the <c>x14</c> extension when a reference in it names another sheet: a
    /// reference to a cell, a 3D reference, or <c>Data!#REF!</c>, which a row delete on <c>Data</c>
    /// leaves (#536). No Excel file holds the last two; each names another sheet, as the others do.
    /// </summary>
    [Test]
    [Arguments("=AND(B1>0,B1<=Data!$A$1)", "x14")]
    [Arguments("=B1<=SUM(Data:'My Data'!$A$1)", "x14")]
    [Arguments("=B1<>Data!#REF!", "x14")]
    [Arguments("=B1<='My Data'!$A$1", "x14")]
    [Arguments("=B1<=MAX($D$1:$D$3)", "standard")]
    [Arguments("=Other!$A$1>0", "standard")] // As Excel wrote B6 in dv-copy-after, without the name
    public async Task A_custom_formula_is_saved_in_x14_when_it_refers_to_another_sheet(string formula,
        string form)
    {
        using var wb = NewBook(out _, out var other);
        wb.AddWorksheet("My Data");
        var rule = other.Cell("B1").CreateDataValidation();

        rule.Custom(formula);

        await AssertSavedAndReloaded(wb, form, SavedDataValidations.AsSaved(formula), "");
    }

    /// <summary>
    /// A criterion that reads a bare <c>#REF!</c> names no sheet, and is saved in the standard form, as
    /// before (#536). Excel has written the same list text in both forms. It kept a list that a sheet
    /// delete left as <c>#REF!</c> in the <c>x14</c> extension (<c>dv-delete-after</c>), and Excel 2013
    /// saved a list that reads <c>#REF!</c> in the standard form
    /// (<c>TryToLoad/TemplateWithTableSourcePivotTables.xlsx</c>). The text cannot tell the two apart.
    /// </summary>
    [Test]
    [Arguments("=#REF!")]
    [Arguments("=OFFSET(#REF!,0,0,3,1)")]
    public async Task A_list_that_reads_a_bare_REF_error_is_saved_in_the_standard_form(string list)
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Cell("B1").CreateDataValidation();

        rule.List(list);

        await AssertSavedAndReloaded(wb, "standard", SavedDataValidations.AsSaved(list), "");
    }

    /// <summary>
    /// A formula the parser refuses is saved as it was before #536, in the standard form, and the save
    /// does not fail.
    /// </summary>
    [Test]
    public async Task A_formula_the_parser_refuses_is_saved_in_the_standard_form()
    {
        using var wb = NewBook(out _, out var other);
        other.Cell("B1").CreateDataValidation().Custom("=SUM(Data!A1");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        await Assert.That(SavedDataValidations.Criteria(ms, "Other"))
            .IsEquivalentTo(new[] { ("standard", "SUM(Data!A1", "") }, CollectionOrdering.Matching);
    }

    /// <summary>
    /// A criterion whose sheet name has one apostrophe instead of two names no sheet, so it is saved in
    /// the standard form. <c>XLHelper.IsValidRangeAddress</c> accepted such text, so the writer took
    /// everything before the last <c>!</c> as a sheet name and wrote the rule in the <c>x14</c>
    /// extension with the broken quoting, while the parser it asks about every other criterion refused
    /// the same text (#560).
    /// </summary>
    [Test]
    [Arguments("='Data!A1", "'Data!A1")]
    [Arguments("=Data'!A1", "Data'!A1")]
    public async Task A_sheet_name_with_an_unterminated_apostrophe_is_saved_in_the_standard_form(
        string criterion, string formula1)
    {
        using var wb = NewBook(out _, out var other);
        other.Cell("B1").CreateDataValidation().Custom(criterion);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        await Assert.That(SavedDataValidations.Criteria(ms, "Other"))
            .IsEquivalentTo(new[] { ("standard", formula1, "") }, CollectionOrdering.Matching);
    }

    /// <summary>
    /// A cell on a sheet whose name holds an apostrophe names the sheet with the apostrophe doubled, and
    /// the rule is saved in the <c>x14</c> extension, as Excel saved the same rule over <c>Bob's</c>
    /// (2026-09-15). The check for another sheet did not read <c>'Bob''s'!$A$1</c> as a range address,
    /// so the rule went to the standard form.
    /// </summary>
    [Test]
    public async Task A_between_rule_over_cells_on_a_sheet_with_an_apostrophe_is_saved_as_Excel_saves_it()
    {
        using var wb = NewBook(out _, out var other);
        var bobs = wb.AddWorksheet("Bob's");

        other.Cell("B1").CreateDataValidation().WholeNumber.Between(bobs.Cell("A1"), bobs.Cell("A2"));

        await AssertSavedAndReloaded(wb, "x14", "'Bob''s'!$A$1", "'Bob''s'!$A$2");
    }

    /// <summary>
    /// <c>List(IXLRange)</c> over a sheet whose name holds an apostrophe stores the range, not a literal
    /// list, and a save writes it in the <c>x14</c> extension, as Excel saved the same list. The range
    /// text was not read as a range address, so it was quoted as a list of one item.
    /// </summary>
    [Test]
    public async Task A_list_from_a_range_on_a_sheet_with_an_apostrophe_is_a_range()
    {
        using var wb = NewBook(out _, out var other);
        var bobs = wb.AddWorksheet("Bob's");
        var rule = other.Cell("B1").CreateDataValidation();

        rule.List(bobs.Range("A1:A3"));

        await Assert.That(rule.MinValue).IsEqualTo("'Bob''s'!$A$1:$A$3");
        await AssertSavedAndReloaded(wb, "x14", "'Bob''s'!$A$1:$A$3", "");
    }

    /// <summary>
    /// The same list on the sheet it refers to is saved in the standard element without the sheet's
    /// name, as Excel saved it.
    /// </summary>
    [Test]
    public async Task A_list_from_a_range_on_its_own_sheet_with_an_apostrophe_is_saved_without_the_name()
    {
        using var wb = NewBook(out _, out _);
        var bobs = wb.AddWorksheet("Bob's");

        bobs.Cell("C1").CreateDataValidation().List(bobs.Range("A1:A3"));

        await AssertSavedAndReloaded(wb, "standard", "$A$1:$A$3", "", sheetName: "Bob's");
    }

    private static XLWorkbook NewBook(out IXLWorksheet data, out IXLWorksheet other)
    {
        var wb = new XLWorkbook();
        data = wb.AddWorksheet("Data");
        other = wb.AddWorksheet("Other");
        return wb;
    }

    private static XLValidationCriteria Of(IXLDataValidation rule, Criteria criteria) => criteria switch
    {
        Criteria.WholeNumber => rule.WholeNumber,
        Criteria.Decimal => rule.Decimal,
        Criteria.Date => rule.Date,
        Criteria.Time => rule.Time,
        Criteria.TextLength => rule.TextLength,
        _ => throw new ArgumentOutOfRangeException(nameof(criteria), criteria, null),
    };

    /// <summary>Sets the rule through the <see cref="IXLCell"/> overload for <paramref name="comparison"/>.</summary>
    private static void Compare(XLValidationCriteria criteria, XLOperator comparison, IXLCell first, IXLCell second)
    {
        switch (comparison)
        {
            case XLOperator.Between:
                criteria.Between(first, second);
                break;
            case XLOperator.NotBetween:
                criteria.NotBetween(first, second);
                break;
            case XLOperator.EqualTo:
                criteria.EqualTo(first);
                break;
            case XLOperator.NotEqualTo:
                criteria.NotEqualTo(first);
                break;
            case XLOperator.GreaterThan:
                criteria.GreaterThan(first);
                break;
            case XLOperator.LessThan:
                criteria.LessThan(first);
                break;
            case XLOperator.EqualOrGreaterThan:
                criteria.EqualOrGreaterThan(first);
                break;
            case XLOperator.EqualOrLessThan:
                criteria.EqualOrLessThan(first);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(comparison), comparison, null);
        }
    }

    /// <summary>Sets the rule through the <see cref="DateTime"/> overload for <paramref name="comparison"/>.</summary>
    private static void CompareDates(XLDateCriteria criteria, XLOperator comparison, DateTime first, DateTime second)
    {
        switch (comparison)
        {
            case XLOperator.Between:
                criteria.Between(first, second);
                break;
            case XLOperator.NotBetween:
                criteria.NotBetween(first, second);
                break;
            case XLOperator.EqualTo:
                criteria.EqualTo(first);
                break;
            case XLOperator.NotEqualTo:
                criteria.NotEqualTo(first);
                break;
            case XLOperator.GreaterThan:
                criteria.GreaterThan(first);
                break;
            case XLOperator.LessThan:
                criteria.LessThan(first);
                break;
            case XLOperator.EqualOrGreaterThan:
                criteria.EqualOrGreaterThan(first);
                break;
            case XLOperator.EqualOrLessThan:
                criteria.EqualOrLessThan(first);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(comparison), comparison, null);
        }
    }

    private static bool TakesTwoValues(XLOperator comparison)
        => comparison is XLOperator.Between or XLOperator.NotBetween;

    /// <summary>
    /// Saves <paramref name="wb"/>, checks that <paramref name="sheetName"/> holds one rule, written in
    /// the form and with the criteria given, and that a reload reads the criteria back as they were
    /// written.
    /// </summary>
    private static async Task AssertSavedAndReloaded(XLWorkbook wb, string form, string formula1, string formula2,
        string sheetName = "Other")
    {
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        await Assert.That(SavedDataValidations.Criteria(ms, sheetName))
            .IsEquivalentTo(new[] { (form, formula1, formula2) }, CollectionOrdering.Matching);

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var rule = reloaded.Worksheet(sheetName).DataValidations.Single();
        await Assert.That(rule.MinValue).IsEqualTo(formula1);
        await Assert.That(rule.MaxValue).IsEqualTo(formula2);
    }
}
