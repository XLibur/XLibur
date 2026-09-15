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
    /// A custom formula set with a leading <c>=</c> is saved without it, with the text Excel wrote for
    /// <c>B6</c>. Excel wrote that rule in the <c>x14</c> extension, because the formula refers to another
    /// sheet. XLibur writes a rule there only when a criterion is a range address on another sheet, so
    /// this one goes to the standard form, which Excel reads. Which form a rule takes belongs to spec 44.
    /// </summary>
    [Test]
    public async Task A_custom_formula_set_with_a_leading_equals_is_saved_without_it()
    {
        using var wb = NewBook(out _, out var other);
        var rule = other.Cell("B6").CreateDataValidation();

        rule.Custom("=B6<=MAX(Data!$A$1:$A$3)");

        await AssertSavedAndReloaded(wb, "standard", "B6<=MAX(Data!$A$1:$A$3)", "");
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

    private static bool TakesTwoValues(XLOperator comparison)
        => comparison is XLOperator.Between or XLOperator.NotBetween;

    /// <summary>
    /// Saves <paramref name="wb"/>, checks that <c>Other</c> holds one rule, written in the form and with
    /// the criteria given, and that a reload reads the criteria back as they were written.
    /// </summary>
    private static async Task AssertSavedAndReloaded(XLWorkbook wb, string form, string formula1, string formula2)
    {
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        await Assert.That(SavedDataValidations.Criteria(ms, "Other"))
            .IsEquivalentTo(new[] { (form, formula1, formula2) }, CollectionOrdering.Matching);

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var rule = reloaded.Worksheet("Other").DataValidations.Single();
        await Assert.That(rule.MinValue).IsEqualTo(formula1);
        await Assert.That(rule.MaxValue).IsEqualTo(formula2);
    }
}
