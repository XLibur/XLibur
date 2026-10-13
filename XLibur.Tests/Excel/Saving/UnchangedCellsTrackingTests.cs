using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Rows;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A save can keep a sheet's cells as the file had them only if nothing written to
/// <c>&lt;sheetData&gt;</c> changed since the load (#702). Each test edits through one path, and
/// checks the sheet is reported as changed, or that reading it is not.
/// </summary>
public class UnchangedCellsTrackingTests
{
    [Test]
    public async Task A_loaded_sheet_is_unchanged()
    {
        using var wb = LoadSource();

        await Assert.That(Changed(Sheet(wb, "Data"))).IsFalse();
        await Assert.That(Changed(Sheet(wb, "Other"))).IsFalse();
    }

    [Test]
    public async Task Reading_a_sheet_does_not_change_it()
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");

        _ = ws.Cell("A1").Value;
        _ = ws.Cell("B2").Style.Font.Bold;
        _ = ws.Cell("C3").FormulaA1;
        _ = ws.Cell("Z99").Value;
        _ = ws.Row(6).Height;
        _ = ws.Column(2).Width;
        _ = ws.CellsUsed().ToList();
        _ = ws.RangeUsed();

        await Assert.That(Changed(ws)).IsFalse();
    }

    [Test]
    public async Task Reading_a_new_row_under_a_styled_column_changes_the_sheet()
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");

        // Materialising row 50 gives B50 the style of column 2, and the save writes that cell.
        _ = ws.Row(50).Height;

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task A_sheet_that_was_not_loaded_counts_as_changed()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("New");

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task A_sheet_added_to_a_loaded_workbook_counts_as_changed()
    {
        using var wb = LoadSource();
        var ws = wb.AddWorksheet("New");

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    [Arguments("value")]
    [Arguments("clear")]
    [Arguments("style")]
    [Arguments("rich text")]
    [Arguments("share string")]
    [Arguments("comment")]
    [Arguments("formula")]
    [Arguments("insert row")]
    [Arguments("delete column")]
    public async Task Editing_a_cell_changes_the_sheet(string edit)
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");

        switch (edit)
        {
            case "value":
                ws.Cell("A1").Value = "Changed";
                break;
            case "clear":
                ws.Cell("B2").Clear();
                break;
            case "style":
                ws.Cell("A1").Style.Font.Italic = true;
                break;
            case "rich text":
                ws.Cell("A1").CreateRichText().AddText(" more");
                break;
            case "share string":
                ws.Cell("A1").ShareString = false;
                break;
            case "comment":
                ws.Cell("E5").CreateComment().AddText("Note");
                break;
            case "formula":
                ws.Cell("D1").FormulaA1 = "B2*3";
                break;
            case "insert row":
                ws.Row(1).InsertRowsAbove(1);
                break;
            case "delete column":
                ws.Column(1).Delete();
                break;
        }

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    [Arguments("height")]
    [Arguments("clear height")]
    [Arguments("hide")]
    [Arguments("group")]
    [Arguments("collapse")]
    [Arguments("phonetic")]
    [Arguments("dy descent")]
    [Arguments("style")]
    public async Task Editing_a_row_changes_the_sheet(string edit)
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");

        // Row 6 has a height in the file and no cells, so only the row snapshot can see these edits.
        var row = (XLRow)ws.Row(6);
        switch (edit)
        {
            case "height":
                row.Height = 40;
                break;
            case "clear height":
                row.ClearHeight();
                break;
            case "hide":
                row.Hide();
                break;
            case "group":
                row.Group();
                break;
            case "collapse":
                row.Collapsed = true;
                break;
            case "phonetic":
                row.ShowPhonetic = true;
                break;
            case "dy descent":
                row.DyDescent = 0.25;
                break;
            case "style":
                row.Style.Fill.BackgroundColor = XLColor.Yellow;
                break;
        }

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task Giving_a_new_row_an_attribute_changes_the_sheet()
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");

        ws.Row(40).Height = 25;

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task Giving_a_column_a_style_changes_the_sheet()
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");

        ws.Column(8).Style.Font.Bold = true;

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task Restyling_a_column_without_passing_the_style_on_changes_the_sheet()
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");
        var column = (XLColumn)ws.Column(2);

        // B3 was loaded with no style of its own, so the save writes column 2's style for it.
        // InnerStyle sets the column's style without passing it to the column's cells, as copying a
        // column does, so no cell slice sees the change.
        await Assert.That(OwnStyle(ws, 3, 2)).IsNull();
        column.InnerStyle = ws.Cell("A1").Style;

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task Restyling_the_sheet_without_passing_the_style_on_changes_the_sheet()
    {
        using var wb = LoadSource();

        // Other has no rows or columns of its own, and A1 was loaded with no style of its own: it
        // takes the sheet's.
        var ws = Sheet(wb, "Other");
        await Assert.That(OwnStyle(ws, 1, 1)).IsNull();
        ((XLWorksheet)ws).InnerStyle = Sheet(wb, "Data").Cell("B2").Style;

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task Giving_the_sheet_a_style_changes_the_sheet()
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");

        ws.Style.Font.FontSize = 20;

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task Switching_the_date_system_changes_the_sheet()
    {
        using var wb = LoadSource();

        wb.Use1904DateSystem = true;

        await Assert.That(Changed(Sheet(wb, "Data"))).IsTrue();
    }

    [Test]
    public async Task Renaming_a_sheet_that_a_formula_names_changes_the_sheet_holding_the_formula()
    {
        using var wb = LoadSource();
        var formulaVersion = FormulaVersion(Sheet(wb, "Data"));

        wb.Worksheet("Other").Name = "Renamed";

        // The formula text is rewritten in place, which the formula slice counts as an edit.
        await Assert.That(FormulaVersion(Sheet(wb, "Data"))).IsNotEqualTo(formulaVersion);
        await Assert.That(Changed(Sheet(wb, "Data"))).IsTrue();
    }

    [Test]
    public async Task Renaming_a_sheet_that_no_formula_names_leaves_the_sheets_unchanged()
    {
        using var wb = LoadSource();

        // A rename rewrites only the formulas that name the sheet, and marks only those dirty (#720).
        wb.Worksheet("Unused").Name = "Renamed";

        await Assert.That(Changed(Sheet(wb, "Other"))).IsFalse();
        await Assert.That(Changed(Sheet(wb, "Data"))).IsFalse();
    }

    [Test]
    public async Task Deleting_a_sheet_that_a_formula_names_changes_the_sheet_holding_the_formula()
    {
        using var wb = LoadSource();
        var formulaVersion = FormulaVersion(Sheet(wb, "Data"));

        // Every reference to the deleted sheet becomes #REF!, in place.
        wb.Worksheet("Other").Delete();

        await Assert.That(FormulaVersion(Sheet(wb, "Data"))).IsNotEqualTo(formulaVersion);
        await Assert.That(Changed(Sheet(wb, "Data"))).IsTrue();
    }

    [Test]
    public async Task Inserting_columns_that_an_array_formula_names_changes_the_sheet_holding_the_formula()
    {
        using var wb = LoadSource();
        var formulaVersion = FormulaVersion(Sheet(wb, "Data"));

        // Moves Other!A1:A2, which the array formula in Data!F1:F2 reads, to B1:B2. The array's text
        // is rewritten in place.
        wb.Worksheet("Other").Column(1).InsertColumnsBefore(1);

        await Assert.That(FormulaVersion(Sheet(wb, "Data"))).IsNotEqualTo(formulaVersion);
        await Assert.That(Changed(Sheet(wb, "Data"))).IsTrue();
    }

    [Test]
    public async Task Adding_a_sheet_changes_a_sheet_whose_formulas_hold_cached_values()
    {
        using var wb = LoadSource();

        // Adding a sheet marks every formula dirty, and a save that does not calculate them writes
        // them without the cached values the file has.
        wb.AddWorksheet("Added");

        await Assert.That(Changed(Sheet(wb, "Data"))).IsTrue();
        await Assert.That(Changed(Sheet(wb, "Other"))).IsFalse();
    }

    [Test]
    public async Task Recalculating_to_the_values_the_file_had_leaves_the_sheet_unchanged()
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");

        wb.RecalculateAllFormulas();

        await Assert.That(ws.Cell("D1").Value).IsEqualTo(2);
        await Assert.That(Changed(ws)).IsFalse();
    }

    [Test]
    public async Task Recalculating_to_a_new_value_changes_the_sheet()
    {
        using var wb = LoadSource(ws => ws.Cell("H1").FormulaA1 = "RAND()");
        var ws = Sheet(wb, "Data");

        wb.RecalculateAllFormulas();

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task A_formula_loaded_without_a_cached_value_leaves_the_sheet_unchanged_until_it_is_calculated()
    {
        using var wb = LoadSource(evaluate: false);
        var ws = Sheet(wb, "Data");

        // The file has no cached values, so the formulas load dirty. A save writes them without one,
        // as the file has them.
        await Assert.That(Changed(ws)).IsFalse();

        _ = ws.Cell("D1").Value;

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task Inserting_rows_that_an_array_formula_names_changes_the_sheet_holding_the_formula()
    {
        using var wb = LoadSource();
        var formulaVersion = FormulaVersion(Sheet(wb, "Data"));

        // Widens Other!A1:A2, which the array formula in Data!F1:F2 reads, so its text is rewritten in
        // place. The normal formula in D1 reads only Other!A1, above the insert, and stays as it is.
        wb.Worksheet("Other").Row(2).InsertRowsAbove(2);

        await Assert.That(FormulaVersion(Sheet(wb, "Data"))).IsNotEqualTo(formulaVersion);
        await Assert.That(Changed(Sheet(wb, "Data"))).IsTrue();
    }

    [Test]
    public async Task Editing_a_cell_that_a_formula_reads_changes_the_sheet_holding_the_formula()
    {
        using var wb = LoadSource();

        // The formula on Data becomes dirty, so the save would recalculate or drop its cached value.
        wb.Worksheet("Other").Cell("A1").Value = 100;

        await Assert.That(Changed(Sheet(wb, "Data"))).IsTrue();
        await Assert.That(Changed(Sheet(wb, "Other"))).IsTrue();
    }

    [Test]
    public async Task Moving_an_array_formula_reference_changes_the_sheet()
    {
        using var wb = LoadSource();
        var ws = Sheet(wb, "Data");

        ws.Cell("F1").FormulaReference = ws.Range("F1:F3").RangeAddress;

        await Assert.That(Changed(ws)).IsTrue();
    }

    [Test]
    public async Task A_sheet_loaded_from_a_template_counts_as_changed()
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var source = LoadSource())
                source.SaveAs(path);

            using var wb = XLWorkbook.OpenFromTemplate(path);

            await Assert.That(Changed(Sheet(wb, "Data"))).IsTrue();
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// Saves a workbook with values, styles, rows with attributes, a styled column and formulas that
    /// read another sheet, and loads it again. The formulas are evaluated before the save, so they
    /// load clean, unless <paramref name="evaluate"/> is false.
    /// </summary>
    /// <param name="addToData">Adds more to the Data sheet before the save.</param>
    /// <param name="evaluate">Whether the save writes the formulas' cached values.</param>
    private static XLWorkbook LoadSource(Action<IXLWorksheet>? addToData = null, bool evaluate = true)
    {
        // Not disposed: a loaded workbook saves by copying the stream it was loaded from.
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            var other = wb.AddWorksheet("Other");
            wb.AddWorksheet("Unused");

            other.Cell("A1").Value = 1;
            other.Cell("A2").Value = 2;

            ws.Cell("A1").Value = "Text";
            ws.Cell("B2").Value = 2.5;
            ws.Cell("B2").Style.Font.Bold = true;
            ws.Cell("C3").Value = new DateTime(2026, 9, 27);
            ws.Cell("D1").FormulaA1 = "Other!A1+1";
            ws.Range("F1:F2").FormulaArrayA1 = "Other!A1:A2*2";
            ws.Row(6).Height = 30;
            ws.Column(2).Style.Fill.BackgroundColor = XLColor.LightBlue;
            ws.Cell("B3").Value = 7;
            addToData?.Invoke(ws);

            wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = evaluate });
        }

        stream.Position = 0;
        return new XLWorkbook(stream);
    }

    private static IXLWorksheet Sheet(XLWorkbook wb, string name) => wb.Worksheet(name);

    private static bool Changed(IXLWorksheet ws) => ((XLWorksheet)ws).CellsChangedSinceLoad();

    private static XLStyleValue? OwnStyle(IXLWorksheet ws, int row, int column)
        => ((XLWorksheet)ws).Internals.CellsCollection.StyleSlice[row, column];

    private static int FormulaVersion(IXLWorksheet ws)
        => ((XLWorksheet)ws).Internals.CellsCollection.FormulaSlice.Version;
}
