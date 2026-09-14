using XLibur.Excel;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Tests.Excel.IO;

namespace XLibur.Tests.Excel.CalcEngine;

public class FormulaCachingTests
{
    [Test]
    public async Task StaticCellDoesNotNeedRecalculation()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var cell = sheet.Cell(1, 1);
        cell.Value = "1234567";

        await Assert.That(cell.NeedsRecalculation).IsFalse();
    }

    [Test]
    public async Task EditCellInvalidatesDependentCells()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var cell = sheet.Cell(1, 1);
        var dependentCell = sheet.Cell(2, 1);
        dependentCell.FormulaA1 = "=A1";
        var _ = dependentCell.Value;

        cell.Value = "1234567";

        await Assert.That(dependentCell.NeedsRecalculation).IsTrue();
    }

    [Test]
    public async Task EditFormulaA1InvalidatesDependentCells()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");
        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=SUM(A1:A3)";
        a1.Value = 15;

        var res1 = a4.Value;
        a2.FormulaA1 = "=A1*20";
        var res2 = a4.Value;

        await Assert.That(res1).IsEqualTo(15 + 150 + 1500);
        await Assert.That(res2).IsEqualTo(15 + 300 + 3000);
    }

    [Test]
    public async Task EditFormulaR1C1InvalidatesDependentCells()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");
        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=SUM(A1:A3)";
        a1.Value = 15;

        var res1 = a4.Value;
        a2.FormulaR1C1 = "=R[-1]C*2";
        var res2 = a4.Value;

        await Assert.That(res1).IsEqualTo(15 + 150 + 1500);
        await Assert.That(res2).IsEqualTo(15 + 30 + 300);
    }

    [Test]
    public async Task InsertRowInvalidatesValues()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a4 = sheet.Cell("A4");
        a4.FormulaA1 = "=COUNTBLANK(A1:A3)";

        await Assert.That(a4.Value).IsEqualTo(3);

        sheet.Row(2).InsertRowsAbove(2);

        await Assert.That(sheet.Cell("A6").Value).IsEqualTo(5);
    }

    [Test]
    public async Task DeleteRowModifiesFormulaAndInvalidatesValues()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var original = sheet.Cell("A4");
        original.FormulaA1 = "=COUNTBLANK(A1:A3)";

        await Assert.That(original.Value).IsEqualTo(3);

        sheet.Row(2).Delete();

        var shifted = sheet.Cell("A3");
        await Assert.That(shifted.FormulaA1).IsEqualTo("COUNTBLANK(A1:A2)");
        await Assert.That(shifted.Value).IsEqualTo(2);
    }

    [Test]
    public async Task ChainedCalculationPreservesIntermediateValues()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");
        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=SUM(A1:A3)";

        a1.Value = 15;
        var res = a4.Value;

        await Assert.That(res).IsEqualTo(15 + 150 + 1500);
        await Assert.That(a4.NeedsRecalculation).IsFalse();
        await Assert.That(a3.NeedsRecalculation).IsFalse();
        await Assert.That(a2.NeedsRecalculation).IsFalse();
        await Assert.That(a2.CachedValue).IsEqualTo(150);
        await Assert.That(a3.CachedValue).IsEqualTo(1500);
        await Assert.That(a4.CachedValue).IsEqualTo(15 + 150 + 1500);
    }

    [Test]
    public async Task EditingAffectsDependentCells()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");
        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=SUM(A1:A3)";
        a1.Value = 15;

        var res1 = a4.Value;
        a1.Value = 20;
        var res2 = a4.Value;

        await Assert.That(res1).IsEqualTo(15 + 150 + 1500);
        await Assert.That(res2).IsEqualTo(20 + 200 + 2000);
    }

    [Test]
    [Arguments("C4", new[] { "C5" })]
    [Arguments("D4", new string[] { })]
    [Arguments("A1", new[] { "A2", "A3", "A4", "C1", "C2", "C3", "C5" })]
    [Arguments("B2", new[] { "B3", "B4", "C2", "C3", "C5" })]
    [Arguments("C2", new[] { "C5" })]
    public async Task EditingDoesNotAffectNonDependingCells(string changedCell, string[] affectedCells)
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        sheet.Cell("A2").FormulaA1 = "A1+1";
        sheet.Cell("A3").FormulaA1 = "SUM(A1:A2)";
        sheet.Cell("A4").FormulaA1 = "SUM(A1:A3)";
        sheet.Cell("B2").FormulaA1 = "B1+1";
        sheet.Cell("B3").FormulaA1 = "SUM(B1:B2)";
        sheet.Cell("B4").FormulaA1 = "SUM(B1:B3)";
        sheet.Cell("C1").FormulaA1 = "SUM(A1:B1)";
        sheet.Cell("C2").FormulaA1 = "SUM(A2:B2)";
        sheet.Cell("C3").FormulaA1 = "SUM(A3:B3)";
        sheet.Cell("C5").FormulaA1 = "SUM($A$1:$C$4)";
        sheet.RecalculateAllFormulas();
        var allCells = sheet.CellsUsed();

        sheet.Cell(changedCell).Value = 100;
        var modifiedCells = allCells.Where(cell => cell.NeedsRecalculation);

        var xlCells = modifiedCells as IXLCell[] ?? modifiedCells.ToArray();
        await Assert.That(xlCells.Length).IsEqualTo(affectedCells.Length);
        foreach (var cellAddress in affectedCells)
        {
            await Assert.That(xlCells.Any(cell => cell.Address.ToString() == cellAddress)).IsTrue().Because($"Cell {cellAddress} is expected to need recalculation, but it does not");
        }
    }

    [Test]
    public async Task CircularReferenceFailsCalculating()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");

        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=A3*10";
        a1.FormulaA1 = "A2+A3+A4";

        var getValueA1 = new Action(() => { _ = a1.Value; });
        var getValueA2 = new Action(() => { _ = a2.Value; });
        var getValueA3 = new Action(() => { _ = a3.Value; });
        var getValueA4 = new Action(() => { _ = a4.Value; });

        await Assert.That(getValueA1).Throws<InvalidOperationException>();
        await Assert.That(getValueA2).Throws<InvalidOperationException>();
        await Assert.That(getValueA3).Throws<InvalidOperationException>();
        await Assert.That(getValueA4).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task CircularReferenceRecalculationNeededDoesNotFail()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");

        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=A3*10";
        var _ = a4.Value;
        a1.FormulaA1 = "=SUM(A2:A4)";

        var recalcNeededA1 = a1.NeedsRecalculation;
        var recalcNeededA2 = a2.NeedsRecalculation;
        var recalcNeededA3 = a3.NeedsRecalculation;
        var recalcNeededA4 = a4.NeedsRecalculation;

        await Assert.That(recalcNeededA1).IsTrue();
        await Assert.That(recalcNeededA2).IsTrue();
        await Assert.That(recalcNeededA3).IsTrue();
        await Assert.That(recalcNeededA4).IsTrue();
    }

    [Test]
    public async Task DeleteWorksheetInvalidatesValues()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.Worksheets.Add("Sheet1");
        var sheet2 = wb.Worksheets.Add("Sheet2");
        var sheet1A1 = sheet1.Cell("A1");
        var sheet2A1 = sheet2.Cell("A1");
        sheet1A1.FormulaA1 = "Sheet2!A1";
        sheet2A1.Value = "TestValue";

        var valueBeforeDeletion = sheet1A1.Value;
        sheet2.Delete();
        var valueAfterDeletion = sheet1A1.Value;

        await Assert.That(valueBeforeDeletion).IsEqualTo("TestValue");
        await Assert.That(valueAfterDeletion).IsEqualTo(XLError.CellReference);
    }

    [Test]
    public async Task CachedValueToExternalWorkbook()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\ExternalLinks\WorkbookWithExternalLink.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        var cell = ws.Cell("B2");
        await Assert.That(cell.NeedsRecalculation).IsFalse();
        await Assert.That(cell.HasFormula).IsTrue();

        // This will fail when we start supporting external links
        await Assert.That(cell.FormulaA1.StartsWith("[1]")).IsTrue();

        await Assert.That(cell.CachedValue).IsEqualTo("hello world");
        await Assert.That(cell.Value).IsEqualTo("hello world");

        await Assert.That(ws.Evaluate("LEN(B2)")).IsEqualTo(11);

        // External file references to evaluate to #REF! instead of throwing
        await Assert.That(wb.RecalculateAllFormulas).ThrowsNothing();
        await Assert.That(cell.Value).IsEqualTo(XLError.CellReference);
    }

    [Test]
    public async Task ChangingValueChangesCachedValue()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var cell = ws.Cell(1, 1);

        cell.Value = "Hello";
        await Assert.That(cell.CachedValue).IsEqualTo("Hello");

        cell.Value = 74.0;
        await Assert.That(cell.CachedValue).IsEqualTo(74.0);

        cell.Value = new DateTime(2019, 1, 1, 14, 0, 0, DateTimeKind.Unspecified);
        await Assert.That(cell.CachedValue).IsEqualTo(new DateTime(2019, 1, 1, 14, 0, 0, DateTimeKind.Unspecified));
    }

    /// <summary>How <see cref="EditAfterLoadInvalidatesDependentCells"/> changes A2.</summary>
    public enum EditOfA2
    {
        Value,
        RangeValue,
        FormulaA1,
        FormulaR1C1,
        InsertData,
        Clear,
    }

    /// <summary>
    /// #504 (D80). A load leaves a formula with a cached value clean, and nothing had told the calc
    /// engine, so until a formula was calculated an edit marked nothing dirty and every dependent
    /// kept the value from the file. Each way of changing A2 must mark what reads it dirty: directly,
    /// through a formula, through a range, from another sheet and through a defined name.
    /// </summary>
    [Test]
    [Arguments(EditOfA2.Value)]
    [Arguments(EditOfA2.RangeValue)]
    [Arguments(EditOfA2.FormulaA1)]
    [Arguments(EditOfA2.FormulaR1C1)]
    [Arguments(EditOfA2.InsertData)]
    [Arguments(EditOfA2.Clear)]
    public async Task EditAfterLoadInvalidatesDependentCells(EditOfA2 edit)
    {
        using var wb = LoadedWithCachedValues();
        var ws = wb.Worksheet("Sheet1");
        var sheet2 = wb.Worksheet("Sheet2");
        await AssertLoadedClean(ws.Range("B1:F1").Cells().Append(sheet2.Cell("A1")));

        var a2 = ws.Cell("A2");
        switch (edit)
        {
            case EditOfA2.Value: a2.Value = 5; break;
            case EditOfA2.RangeValue: ws.Range("A2:A2").Value = 5; break;
            case EditOfA2.FormulaA1: a2.FormulaA1 = "2+3"; break;
            case EditOfA2.FormulaR1C1: a2.FormulaR1C1 = "2+3"; break;
            case EditOfA2.InsertData: a2.InsertData(new[] { 5 }); break;
            case EditOfA2.Clear: a2.Clear(); break;
            default: throw new ArgumentOutOfRangeException(nameof(edit));
        }

        var a2Value = edit == EditOfA2.Clear ? 0 : 5;
        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("C1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("E1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("F1").NeedsRecalculation).IsTrue();
        await Assert.That(sheet2.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("D1").NeedsRecalculation).IsFalse();

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(a2Value * 10);
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(a2Value * 20);
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("E1").Value).IsEqualTo(a2Value + 2);
        await Assert.That(ws.Cell("F1").Value).IsEqualTo(a2Value * 100);
        await Assert.That(sheet2.Cell("A1").Value).IsEqualTo(a2Value * 3);
    }

    /// <summary>
    /// #504. A shared formula is loaded as a formula in each of its cells, and each must hear of an
    /// edit to the cell it reads. Only B2 reads A2.
    /// </summary>
    [Test]
    public async Task EditAfterLoadInvalidatesDependentSharedFormula()
    {
        using var wb = Reload(builder =>
        {
            var ws = builder.AddWorksheet("Sheet1");
            for (var row = 1; row <= 3; row++)
            {
                ws.Cell(row, 1).Value = row;
                ws.Cell(row, 2).FormulaA1 = $"A{row}*10";
            }
        }, WithSharedFormulaInB1ToB3);
        var sheet = wb.Worksheet("Sheet1");
        await Assert.That(sheet.Cell("B2").FormulaA1).IsEqualTo("A2*10");
        await AssertLoadedClean(sheet.Range("B1:B3").Cells());

        sheet.Cell("A2").Value = 5;

        await Assert.That(sheet.Cell("B1").NeedsRecalculation).IsFalse();
        await Assert.That(sheet.Cell("B2").NeedsRecalculation).IsTrue();
        await Assert.That(sheet.Cell("B3").NeedsRecalculation).IsFalse();
        await Assert.That(sheet.Cell("B2").Value).IsEqualTo(50);
    }

    /// <summary>#504. A loaded array formula must hear of an edit to a cell of its range.</summary>
    [Test]
    public async Task EditAfterLoadInvalidatesDependentArrayFormula()
    {
        using var wb = Reload(builder =>
        {
            var ws = builder.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Range("C1:C2").FormulaArrayA1 = "A1:A2*10";
        });
        var sheet = wb.Worksheet("Sheet1");
        await AssertLoadedClean(sheet.Range("C1:C2").Cells());

        sheet.Cell("A2").Value = 5;

        await Assert.That(sheet.Cell("C2").NeedsRecalculation).IsTrue();
        await Assert.That(sheet.Cell("C1").Value).IsEqualTo(10);
        await Assert.That(sheet.Cell("C2").Value).IsEqualTo(50);
    }

    /// <summary>
    /// Once an edit has built the dependency tree, a formula set afterwards must be in it. The edit
    /// builds the tree without the calculation chain, and a formula was added to the tree only when
    /// both existed. So G1, calculated on its own and clean, missed the next edit to A2 and kept 6.
    /// The fix for #504 makes the first edit after every load build the tree this way.
    /// </summary>
    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task FormulaSetAfterTheTreeIsBuiltIsInvalidatedByALaterEdit(bool loaded)
    {
        using var wb = loaded ? LoadedWithCachedValues() : new XLWorkbook();
        var ws = loaded ? wb.Worksheet("Sheet1") : wb.AddWorksheet("Sheet1");
        if (!loaded)
        {
            ws.Cell("B1").FormulaA1 = "A2*10";
            _ = ws.Cell("B1").Value;
        }

        ws.Cell("A2").Value = 5;
        var g1 = ws.Cell("G1");
        g1.FormulaA1 = "A2+1";
        await Assert.That(g1.Value).IsEqualTo(6);

        // B1 is replaced by a formula that reads A3 instead, so the edit of A3 below must reach it.
        ws.Cell("B1").FormulaA1 = "A3*10";
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(loaded ? 10 : 0);

        ws.Cell("A2").Value = 7;
        ws.Cell("A3").Value = 4;

        await Assert.That(g1.NeedsRecalculation).IsTrue();
        await Assert.That(g1.Value).IsEqualTo(8);
        await Assert.That(ws.Cell("B1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(40);
    }

    /// <summary>
    /// A formula that reads a sheet the workbook does not have is in the dependency tree with an area
    /// on that sheet, which no sheet tree holds. Taking the formula out of the tree threw
    /// <see cref="InvalidOperationException"/> for that area, so once the tree was built, setting a
    /// value on the cell threw. Since #504 an edit after a load builds the tree, so a single edit was
    /// enough; after a recalculation it threw already.
    /// </summary>
    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task FormulaThatReadsAMissingSheetCanBeReplacedOnceTheTreeIsBuilt(bool recalculated)
    {
        using var wb = Reload(builder => builder.AddWorksheet("Sheet1").Cell("A1").FormulaA1 = "Missing!B1*2");
        var ws = wb.Worksheet("Sheet1");
        await AssertLoadedClean([ws.Cell("A1")]);
        if (recalculated)
            wb.RecalculateAllFormulas();

        ws.Cell("C1").Value = 1;
        ws.Cell("A1").Value = 5;

        await Assert.That(ws.Cell("A1").HasFormula).IsFalse();
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(5);
    }

    /// <summary>How a test cell holds its rich text in the saved file.</summary>
    public enum RichTextKind
    {
        /// <summary>A shared string with runs.</summary>
        SharedString,

        /// <summary>An inline string, <c>&lt;is&gt;</c>, with runs.</summary>
        InlineString,
    }

    /// <summary>
    /// #504 review. The loader writes a rich-text cell through the setter an edit uses, which tells
    /// the calc engine. Once a formula had loaded clean, that built the dependency tree in the middle
    /// of the load, and the tree lacked every formula loaded after it. D10 = A2*10 was one of them,
    /// so the edit of A2 missed it: D10 read 10 and a save wrote 10. Rich text in row 1, before any
    /// formula, is the control: the tree was not built there.
    /// </summary>
    [Test]
    [Arguments(RichTextKind.SharedString, "C3")]
    [Arguments(RichTextKind.InlineString, "C3")]
    [Arguments(RichTextKind.SharedString, "C1")]
    [Arguments(RichTextKind.InlineString, "C1")]
    public async Task EditAfterLoadReachesAFormulaLoadedAfterRichText(RichTextKind kind, string richTextCell)
    {
        using var wb = Reload(builder =>
        {
            var ws = builder.AddWorksheet("Sheet1");
            AddRichText(ws, richTextCell, kind);
            ws.Cell("A2").Value = 1;
            ws.Cell("B2").FormulaA1 = "B9*2";
            ws.Cell("B9").Value = 4;
            ws.Cell("D10").FormulaA1 = "A2*10";
        });
        var sheet = wb.Worksheet("Sheet1");
        await AssertLoadedAsRichText(sheet.Cell(richTextCell), kind);
        await AssertLoadedClean([sheet.Cell("B2"), sheet.Cell("D10")]);

        sheet.Cell("A2").Value = 5;

        await Assert.That(sheet.Cell("D10").NeedsRecalculation).IsTrue();
        await Assert.That(sheet.Cell("B2").NeedsRecalculation).IsFalse();
        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        await Assert.That(EvaluationOutcomeTests.CachedValueInFile(saved, "D10")).IsEqualTo("D10 has no <v>");
        await Assert.That(sheet.Cell("D10").Value).IsEqualTo(50);
        await Assert.That(sheet.Cell("B2").Value).IsEqualTo(8);
    }

    /// <summary>
    /// #504 review. Once the dependency tree existed in the middle of a load, each rich-text cell
    /// loaded after it marked the formulas that read it dirty, and they lost Excel's cached value
    /// before anything was edited. A1 = LEN(C3) was calculated again when read. B1, a formula XLibur
    /// cannot evaluate, threw instead of returning its cached value.
    /// </summary>
    [Test]
    [Arguments(RichTextKind.SharedString)]
    [Arguments(RichTextKind.InlineString)]
    public async Task LoadedFormulaThatReadsRichTextKeepsItsCachedValue(RichTextKind kind)
    {
        using var wb = Reload(
            builder =>
            {
                var ws = builder.AddWorksheet("Sheet1");
                ws.Cell("A1").FormulaA1 = "LEN(C3)";
                ws.Cell("B1").FormulaA1 = "C3&Sdemo123|tik!'id1?req?AAPL'";
                AddRichText(ws, "C3", kind);
            },
            afterCalculation: builder =>
            {
                // B1 cannot be calculated, so it is given the cached value Excel would have saved.
                var b1 = (XLCell)builder.Worksheet("Sheet1").Cell("B1");
                b1.Worksheet.Internals.CellsCollection.ValueSlice.SetCellValue(b1.SheetPoint, 7);
                b1.Formula!.MarkClean();
            });
        var sheet = wb.Worksheet("Sheet1");
        await AssertLoadedAsRichText(sheet.Cell("C3"), kind);

        await Assert.That(sheet.Cell("A1").NeedsRecalculation).IsFalse();
        await Assert.That(sheet.Cell("B1").NeedsRecalculation).IsFalse();
        await Assert.That(sheet.Cell("A1").Value).IsEqualTo(2);
        await Assert.That(sheet.Cell("B1").Value).IsEqualTo(7);
    }

    /// <summary>Gives <paramref name="address"/> the rich text "ab", "a" in bold.</summary>
    private static void AddRichText(IXLWorksheet ws, string address, RichTextKind kind)
    {
        var cell = ws.Cell(address);
        var richText = cell.GetRichText();
        richText.AddText("a").SetBold();
        richText.AddText("b");
        if (kind == RichTextKind.InlineString)
            cell.ShareString = false;
    }

    /// <summary>
    /// Checks the cell came back as rich text of <paramref name="kind"/>, so the test runs the
    /// loader path it names.
    /// </summary>
    private static async Task AssertLoadedAsRichText(IXLCell cell, RichTextKind kind)
    {
        await Assert.That(cell.HasRichText).IsTrue();
        await Assert.That(cell.ShareString).IsEqualTo(kind == RichTextKind.SharedString);
    }

    /// <summary>
    /// A workbook as a load leaves it, every formula clean with the value it was saved with. On
    /// Sheet1, A1:A3 hold 1, and the name <c>Rate</c> refers to A2. B1 = A2*10, C1 = B1*2,
    /// D1 = A3*2, E1 = SUM(A1:A3), F1 = Rate*100, and Sheet2!A1 = Sheet1!A2*3. Everything but D1
    /// reads A2.
    /// </summary>
    private static XLWorkbook LoadedWithCachedValues() => Reload(wb =>
    {
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 1;
        ws.Cell("A3").Value = 1;
        wb.DefinedNames.Add("Rate", "Sheet1!$A$2");
        ws.Cell("B1").FormulaA1 = "A2*10";
        ws.Cell("C1").FormulaA1 = "B1*2";
        ws.Cell("D1").FormulaA1 = "A3*2";
        ws.Cell("E1").FormulaA1 = "SUM(A1:A3)";
        ws.Cell("F1").FormulaA1 = "Rate*100";
        wb.AddWorksheet("Sheet2").Cell("A1").FormulaA1 = "Sheet1!A2*3";
    });

    /// <summary>
    /// Builds a workbook with <paramref name="build"/>, calculates it, so each formula is saved with
    /// a cached value, and loads it again. <paramref name="afterCalculation"/> runs between the
    /// calculation and the save, and <paramref name="rewriteSheet1"/> can edit the saved sheet.
    /// </summary>
    private static XLWorkbook Reload(Action<XLWorkbook> build, Func<string, string>? rewriteSheet1 = null,
        Action<XLWorkbook>? afterCalculation = null)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            build(wb);
            wb.RecalculateAllFormulas();
            afterCalculation?.Invoke(wb);
            wb.SaveAs(package);
        }

        if (rewriteSheet1 is not null)
            package.RewriteSheet1(rewriteSheet1);

        package.Position = 0;
        return new XLWorkbook(package);
    }

    /// <summary>
    /// A load that left a formula dirty would pass these tests without the fix, so each formula is
    /// first checked to be clean.
    /// </summary>
    private static async Task AssertLoadedClean(System.Collections.Generic.IEnumerable<IXLCell> formulaCells)
    {
        foreach (var cell in formulaCells)
        {
            await Assert.That(cell.HasFormula).IsTrue();
            await Assert.That(cell.NeedsRecalculation).IsFalse();
        }
    }

    /// <summary>Makes B1:B3, each <c>A{row}*10</c>, one shared formula, as Excel writes it.</summary>
    private static string WithSharedFormulaInB1ToB3(string sheetXml)
    {
        string[][] replacements =
        [
            ["<x:f>A1*10</x:f>", "<x:f t=\"shared\" ref=\"B1:B3\" si=\"0\">A1*10</x:f>"],
            ["<x:f>A2*10</x:f>", "<x:f t=\"shared\" si=\"0\" />"],
            ["<x:f>A3*10</x:f>", "<x:f t=\"shared\" si=\"0\" />"],
        ];

        foreach (var replacement in replacements)
        {
            var rewritten = sheetXml.Replace(replacement[0], replacement[1], StringComparison.Ordinal);
            if (ReferenceEquals(rewritten, sheetXml))
                throw new InvalidOperationException($"'{replacement[0]}' was not found in the sheet part.");

            sheetXml = rewritten;
        }

        return sheetXml;
    }
}
