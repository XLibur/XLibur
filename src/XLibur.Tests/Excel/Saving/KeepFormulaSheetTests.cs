using System;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;
using SaveOptions = XLibur.Excel.SaveOptions;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A sheet with formulas keeps its part (#702, tier 1) or its cells (tier 2) on save while its
/// formulas' text and cached values are as the file had them (#720). Each edit of another sheet that
/// rewrites a formula in place, or makes the save recalculate or drop a cached value, has the sheet
/// written from the model.
/// </summary>
/// <remarks>
/// The source spells the Data sheet's XML declaration as no save writes it, so the saved part has
/// that declaration only when the save left the part as it was. It also writes the number 1.5 as
/// <c>1.50</c>, so the saved part holds <c>1.50</c> only when the cells were kept.
/// </remarks>
public class KeepFormulaSheetTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    private const string Declaration = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";

    private const string KeptCells = "<x:v>1.50</x:v>";

    [Test]
    public async Task An_unchanged_formula_sheet_keeps_its_part()
    {
        using var source = Source();

        using var saved = LoadAndSave(source);

        await Assert.That(IsAsTheFileHadIt(saved)).IsTrue();
        await AssertFormulas(saved, "Other");
    }

    [Test]
    public async Task An_unchanged_formula_sheet_keeps_its_part_when_the_save_evaluates_formulas()
    {
        using var source = Source();

        using var saved = LoadAndSave(source, evaluate: true);

        await Assert.That(IsAsTheFileHadIt(saved)).IsTrue();
        await AssertFormulas(saved, "Other");
    }

    [Test]
    public async Task A_change_outside_the_cells_keeps_the_formula_cells()
    {
        using var source = Source();

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Data").Column("H").Width = 30);

        await Assert.That(IsAsTheFileHadIt(saved)).IsFalse();
        await Assert.That(saved.Sheet1Xml()).Contains(KeptCells);
        await AssertFormulas(saved, "Other");
    }

    [Test]
    public async Task A_shared_formula_is_kept_as_the_file_wrote_it()
    {
        using var source = Source(xml => xml
            .Replace("<x:f>A2*2</x:f>", """<x:f t="shared" ref="C2:C3" si="0">A2*2</x:f>""", StringComparison.Ordinal)
            .Replace("<x:f>A3*2</x:f>", """<x:f t="shared" si="0" />""", StringComparison.Ordinal));
        await Assert.That(source.Sheet1Xml()).Contains("t=\"shared\"");

        using var saved = LoadAndSave(source);

        await Assert.That(IsAsTheFileHadIt(saved)).IsTrue();
        await AssertFormulas(saved, "Other");
    }

    [Test]
    public async Task Renaming_the_sheet_a_formula_names_writes_the_formula_sheet()
    {
        using var source = Source();

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Other").Name = "Renamed");

        await AssertWritten(saved);
        await AssertFormulas(saved, "Renamed");
    }

    [Test]
    public async Task Renaming_a_sheet_that_no_formula_names_keeps_the_formula_sheet()
    {
        using var source = Source();

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Unused").Name = "Renamed");

        await Assert.That(IsAsTheFileHadIt(saved)).IsTrue();
        await AssertFormulas(saved, "Other");
    }

    [Test]
    public async Task Deleting_the_sheet_a_formula_names_writes_the_formula_sheet()
    {
        using var source = Source();

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Other").Delete());

        await AssertWritten(saved);
        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("D1").FormulaA1).Contains("#REF!");
    }

    [Test]
    public async Task Inserting_rows_into_the_sheet_a_formula_reads_writes_the_formula_sheet()
    {
        using var source = Source();

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Other").Row(1).InsertRowsAbove(1));

        await AssertWritten(saved);
        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var ws = reloaded.Worksheet("Data");
        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo("Other!A2+1");
        await Assert.That(ws.Cell("F1").FormulaA1).IsEqualTo("Other!A2:A3*2");
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(2);
    }

    [Test]
    public async Task Editing_a_cell_a_formula_reads_writes_the_formula_sheet_with_the_new_value()
    {
        using var source = Source();

        using var saved = LoadAndSave(source, wb => wb.Worksheet("Other").Cell("A1").Value = 10, evaluate: true);

        await AssertWritten(saved);
        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("D1").CachedValue).IsEqualTo(11);
    }

    [Test]
    public async Task Adding_a_sheet_writes_the_formula_sheet_when_the_save_does_not_evaluate_formulas()
    {
        // Adding a sheet marks every formula dirty. A save that does not calculate them writes them
        // without their cached values, so it cannot keep the file's.
        using var source = Source();

        using var saved = LoadAndSave(source, wb => wb.AddWorksheet("Added"));

        await AssertWritten(saved);
        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("D1").Value).IsEqualTo(2);
    }

    [Test]
    public async Task Adding_a_sheet_keeps_the_formula_sheet_when_the_save_recalculates_the_same_values()
    {
        using var source = Source();

        using var saved = LoadAndSave(source, wb => wb.AddWorksheet("Added"), evaluate: true);

        await Assert.That(IsAsTheFileHadIt(saved)).IsTrue();
        await AssertFormulas(saved, "Other");
    }

    [Test]
    public async Task A_formula_the_file_has_without_a_cached_value_keeps_its_part()
    {
        using var source = Source(evaluate: false);

        using var saved = LoadAndSave(source);

        await Assert.That(IsAsTheFileHadIt(saved)).IsTrue();
        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Data").Cell("D1").Value).IsEqualTo(2);
    }

    [Test]
    public async Task A_save_that_calculates_a_formula_the_file_has_without_a_cached_value_writes_it()
    {
        using var source = Source(evaluate: false);

        using var saved = LoadAndSave(source, evaluate: true);

        await AssertWritten(saved);
        await AssertFormulas(saved, "Other");
    }

    /// <summary>
    /// A package with sheets Data, Other and Unused. Data holds 1.5 and 2 in A2:A3, formulas that
    /// double them in C2:C3, <c>Other!A1+1</c> in D1 and the array formula <c>Other!A1:A2*2</c> in
    /// F1:F2. Other holds 1 and 2 in A1:A2. Data's part has <see cref="Declaration"/> and writes 1.5
    /// as <see cref="KeptCells"/>.
    /// </summary>
    /// <param name="rewrite">Rewrites the Data part further.</param>
    /// <param name="evaluate">Whether the file has the formulas' cached values.</param>
    private static MemoryStream Source(Func<string, string>? rewrite = null, bool evaluate = true)
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            var other = wb.AddWorksheet("Other");
            wb.AddWorksheet("Unused").Cell("A1").Value = "Unused";

            other.Cell("A1").Value = 1;
            other.Cell("A2").Value = 2;

            ws.Cell("A1").Value = "Header";
            ws.Cell("A2").Value = 1.5;
            ws.Cell("A3").Value = 2;
            ws.Cell("C2").FormulaA1 = "A2*2";
            ws.Cell("C3").FormulaA1 = "A3*2";
            ws.Cell("D1").FormulaA1 = "Other!A1+1";
            ws.Range("F1:F2").FormulaArrayA1 = "Other!A1:A2*2";

            wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = evaluate });
        }

        return stream.RewritePart(Sheet1, xml =>
        {
            var afterDeclaration = xml.IndexOf("?>", StringComparison.Ordinal) + 2;
            var marked = Declaration + xml[afterDeclaration..].Replace("<x:v>1.5</x:v>", KeptCells, StringComparison.Ordinal);
            return rewrite is null ? marked : rewrite(marked);
        });
    }

    /// <summary>Does the Data part still have the declaration no save would write?</summary>
    private static bool IsAsTheFileHadIt(Stream package) =>
        package.ReadPart(Sheet1).StartsWith(Declaration, StringComparison.Ordinal);

    private static async Task AssertWritten(MemoryStream package)
    {
        await Assert.That(IsAsTheFileHadIt(package)).IsFalse();
        await Assert.That(package.Sheet1Xml()).DoesNotContain(KeptCells);
    }

    /// <summary>
    /// Reloads the package and checks Data's formulas, which name <paramref name="otherName"/>, and
    /// their values.
    /// </summary>
    private static async Task AssertFormulas(MemoryStream package, string otherName)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A2").GetDouble()).IsEqualTo(1.5);
        await Assert.That(ws.Cell("C2").FormulaA1).IsEqualTo("A2*2");
        await Assert.That(ws.Cell("C3").FormulaA1).IsEqualTo("A3*2");
        await Assert.That(ws.Cell("D1").FormulaA1).IsEqualTo($"{otherName}!A1+1");
        await Assert.That(ws.Cell("F1").FormulaA1).IsEqualTo($"{otherName}!A1:A2*2");
        await Assert.That(ws.Cell("F2").FormulaA1).IsEqualTo($"{otherName}!A1:A2*2");
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(3);
        await Assert.That(ws.Cell("C3").Value).IsEqualTo(4);
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("F1").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("F2").Value).IsEqualTo(4);
    }

    private static MemoryStream LoadAndSave(MemoryStream source, Action<XLWorkbook>? edit = null, bool evaluate = false)
    {
        source.Position = 0;
        using var wb = new XLWorkbook(source);
        edit?.Invoke(wb);

        var saved = new MemoryStream();
        wb.SaveAs(saved, new SaveOptions { EvaluateFormulasBeforeSaving = evaluate });
        saved.Position = 0;
        return saved;
    }
}
