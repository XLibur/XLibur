using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using SaveOptions = XLibur.Excel.SaveOptions;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// #1028. A formula saved with no cached value (ADR 0001) relies on Excel recalculating it on open.
/// A workbook loaded from current Excel keeps a <c>calcId</c> that does not make Excel do so, so the
/// save asks for it with <c>fullCalcOnLoad</c>.
/// </summary>
public class FullCalculationOnLoadTests
{
    /// <summary>XLibur refuses it, so a save that evaluates formulas leaves it dirty.</summary>
    private const string Unevaluable = "ABS(1,2)";

    [Test]
    [Arguments(Unevaluable)]
    [Arguments("A1+1")]
    public async Task A_formula_saved_with_no_cached_value_makes_Excel_recalculate_on_load(string formula)
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1").Cell("A1").FormulaA1 = formula;

        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        await Assert.That(FullCalculationOnLoad(stream)).IsTrue();
        await Assert.That(wb.FullCalculationOnLoad).IsFalse();
    }

    [Test]
    public async Task A_workbook_whose_formulas_all_evaluate_is_not_marked_for_recalculation()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1").Cell("A1").FormulaA1 = "1+1";

        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        await Assert.That(FullCalculationOnLoad(stream)).IsFalse();
    }

    [Test]
    public async Task The_flag_is_not_carried_into_a_later_save_whose_formulas_all_evaluate()
    {
        using var wb = new XLWorkbook();
        var a1 = wb.AddWorksheet("Sheet1").Cell("A1");
        a1.FormulaA1 = Unevaluable;
        var options = new SaveOptions { EvaluateFormulasBeforeSaving = true };

        using var first = new MemoryStream();
        wb.SaveAs(first, options);
        a1.FormulaA1 = "1+1";
        using var second = new MemoryStream();
        wb.SaveAs(second, options);

        await Assert.That(FullCalculationOnLoad(first)).IsTrue();
        await Assert.That(FullCalculationOnLoad(second)).IsFalse();
    }

    /// <summary>
    /// The issue's case: sheet 2, C1 holds <c>averageif(A2:B2)</c>, which XLibur cannot evaluate. An
    /// edit marks it dirty, and the file keeps its own <c>calcPr</c>.
    /// </summary>
    [Test]
    public async Task After_an_edit_a_loaded_workbook_with_an_unevaluable_formula_is_marked_for_recalculation()
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LO\xlsx\bug-fixes.xlsx"));
        using var stream = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            await Assert.That(wb.FullCalculationOnLoad).IsFalse();
            wb.Worksheet(2).Cell("E5").Value = 1;
            wb.SaveAs(stream);
        }

        await Assert.That(FullCalculationOnLoad(stream)).IsTrue();
    }

    /// <summary>
    /// A second save keeps the sheet the first one wrote, so no formula is written again, yet the
    /// formula still has no cached value.
    /// </summary>
    [Test]
    public async Task A_second_save_that_keeps_the_sheet_still_marks_the_workbook_for_recalculation()
    {
        using var package = new MemoryStream();
        using (var created = new XLWorkbook())
        {
            created.AddWorksheet("Sheet1").Cell("B1").Value = 1;
            created.SaveAs(package);
        }

        using (var wb = new XLWorkbook(package))
        {
            wb.Worksheet(1).Cell("A1").FormulaA1 = Unevaluable;
            wb.Save();
            await Assert.That(FullCalculationOnLoad(package)).IsTrue();
            wb.Save();
        }

        await Assert.That(FullCalculationOnLoad(package)).IsTrue();
    }

    private static bool FullCalculationOnLoad(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return document.WorkbookPart!.Workbook!.CalculationProperties?.FullCalculationOnLoad?.Value ?? false;
    }
}
