using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Coordinates;
using XLibur.Excel.IO;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// A save writes back the shared formulas that a load brought in, limited to the cells that still
/// hold them (ADR 0003, #692). Every way a cell leaves its shared formula writes it as a formula of
/// its own, and every file written must keep the rules Excel reads shared formulas by.
/// </summary>
public class SharedFormulaWriteTests
{
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(10);

    private static readonly Regex FormulaCell = new(
        "<x:c r=\"([A-Z]+[0-9]+)\"[^>]*>(<x:f[^>]*/>|<x:f[^>]*>[^<]*</x:f>)", RegexOptions.None, RegexTimeout);

    private static readonly Regex Attribute = new("(\\w+)=\"([^\"]*)\"", RegexOptions.None, RegexTimeout);

    [Test]
    public async Task An_unchanged_shared_formula_is_written_back_as_the_file_had_it()
    {
        using var wb = LoadWithSharedColumn();

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B10\" si=\"0\">A1*2</x:f>");
        for (var row = 2; row <= 10; row++)
            await Assert.That(f[$"B{row}"]).IsEqualTo("<x:f t=\"shared\" si=\"0\" />");

        await AssertReadsBack(saved, ExpectedColumn());
    }

    [Test]
    public async Task The_switch_off_writes_a_formula_in_each_cell()
    {
        using var wb = LoadWithSharedColumn();

        var saved = Save(wb, shared: false);

        var f = FormulaElements(saved);
        for (var row = 1; row <= 10; row++)
            await Assert.That(f[$"B{row}"]).IsEqualTo($"<x:f>A{row}*2</x:f>");
    }

    [Test]
    public async Task A_formula_XLibur_wrote_is_never_written_as_shared()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        for (var row = 1; row <= 3; row++)
            ws.Cell(row, 2).FormulaA1 = $"A{row}*2";

        var f = FormulaElements(Save(wb));

        for (var row = 1; row <= 3; row++)
            await Assert.That(f[$"B{row}"]).IsEqualTo($"<x:f>A{row}*2</x:f>");
    }

    [Test]
    public async Task An_edited_cell_inside_the_formula_is_written_on_its_own_and_the_rest_stay_shared()
    {
        using var wb = LoadWithSharedColumn();
        wb.Worksheet(1).Cell("B5").FormulaA1 = "A5*3";

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B10\" si=\"0\">A1*2</x:f>");
        await Assert.That(f["B5"]).IsEqualTo("<x:f>A5*3</x:f>");
        await Assert.That(f["B6"]).IsEqualTo("<x:f t=\"shared\" si=\"0\" />");
        await AssertReadsBack(saved, ExpectedColumn(("B5", "A5*3")));
    }

    [Test]
    public async Task A_cell_given_a_value_leaves_the_formula()
    {
        using var wb = LoadWithSharedColumn();
        wb.Worksheet(1).Cell("B10").Value = 99;

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f.ContainsKey("B10")).IsFalse();
        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B9\" si=\"0\">A1*2</x:f>");
        await AssertReadsBack(saved, ExpectedColumn(("B10", null)));
    }

    [Test]
    public async Task When_the_master_is_edited_the_next_cell_becomes_the_master()
    {
        using var wb = LoadWithSharedColumn();
        wb.Worksheet(1).Cell("B1").FormulaA1 = "A1*3";

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f>A1*3</x:f>");
        await Assert.That(f["B2"]).IsEqualTo("<x:f t=\"shared\" ref=\"B2:B10\" si=\"0\">A2*2</x:f>");
        await Assert.That(f["B3"]).IsEqualTo("<x:f t=\"shared\" si=\"0\" />");
        await AssertReadsBack(saved, ExpectedColumn(("B1", "A1*3")));
    }

    [Test]
    public async Task A_formula_left_with_one_cell_is_written_on_its_own()
    {
        using var wb = LoadWithSharedColumn();
        var ws = wb.Worksheet(1);
        for (var row = 1; row <= 10; row++)
        {
            if (row != 4)
                ws.Cell(row, 2).Clear();
        }

        var f = FormulaElements(Save(wb));

        await Assert.That(f.Count).IsEqualTo(1);
        await Assert.That(f["B4"]).IsEqualTo("<x:f>A4*2</x:f>");
    }

    [Test]
    public async Task A_row_inserted_above_the_formula_writes_every_cell_on_its_own()
    {
        using var wb = LoadWithSharedColumn();
        wb.Worksheet(1).Row(1).InsertRowsAbove(1);

        var saved = Save(wb);

        var f = FormulaElements(saved);
        for (var row = 2; row <= 11; row++)
            await Assert.That(f[$"B{row}"]).IsEqualTo($"<x:f>A{row}*2</x:f>");

        await AssertReadsBack(saved, Enumerable.Range(2, 10).ToDictionary(r => $"B{r}", r => (string?)$"A{r}*2"));
    }

    [Test]
    public async Task A_row_inserted_inside_the_formula_keeps_the_cells_above_it_shared()
    {
        using var wb = LoadWithSharedColumn();
        wb.Worksheet(1).Row(5).InsertRowsAbove(1);

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B4\" si=\"0\">A1*2</x:f>");
        await Assert.That(f["B4"]).IsEqualTo("<x:f t=\"shared\" si=\"0\" />");
        for (var row = 6; row <= 11; row++)
            await Assert.That(f[$"B{row}"]).IsEqualTo($"<x:f>A{row}*2</x:f>");

        var expected = new Dictionary<string, string?>();
        for (var row = 1; row <= 4; row++)
            expected[$"B{row}"] = $"A{row}*2";
        for (var row = 6; row <= 11; row++)
            expected[$"B{row}"] = $"A{row}*2";
        await AssertReadsBack(saved, expected);
    }

    [Test]
    public async Task A_row_deleted_inside_the_formula_keeps_the_cells_above_it_shared()
    {
        using var wb = LoadWithSharedColumn();
        wb.Worksheet(1).Row(5).Delete();

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B4\" si=\"0\">A1*2</x:f>");
        for (var row = 5; row <= 9; row++)
            await Assert.That(f[$"B{row}"]).IsEqualTo($"<x:f>A{row}*2</x:f>");

        await AssertReadsBack(saved, Enumerable.Range(1, 9).ToDictionary(r => $"B{r}", r => (string?)$"A{r}*2"));
    }

    [Test]
    public async Task A_column_inserted_left_of_the_formula_writes_every_cell_on_its_own()
    {
        using var wb = LoadWithSharedColumn();
        wb.Worksheet(1).Column(1).InsertColumnsBefore(1);

        var saved = Save(wb);

        var f = FormulaElements(saved);
        for (var row = 1; row <= 10; row++)
            await Assert.That(f[$"C{row}"]).IsEqualTo($"<x:f>B{row}*2</x:f>");

        await AssertReadsBack(saved, Enumerable.Range(1, 10).ToDictionary(r => $"C{r}", r => (string?)$"B{r}*2"));
    }

    /// <summary>
    /// A row deleted below the formula leaves its cells where they are, but a reference to a row below
    /// it is rewritten, which changes the A1 text.
    /// </summary>
    [Test]
    public async Task A_reference_rewritten_by_a_shift_writes_the_cell_on_its_own()
    {
        using var wb = LoadWithSharedColumn("A{0}+$A$20");
        wb.Worksheet(1).Row(15).Delete();

        var saved = Save(wb);

        var f = FormulaElements(saved);
        for (var row = 1; row <= 10; row++)
            await Assert.That(f[$"B{row}"]).IsEqualTo($"<x:f>A{row}+$A$19</x:f>");

        await AssertValidSharedFormulas(saved);
    }

    [Test]
    public async Task A_copy_of_the_formula_is_written_on_its_own_and_the_original_stays_shared()
    {
        using var wb = LoadWithSharedColumn();
        var ws = wb.Worksheet(1);
        ws.Range("B1:B10").CopyTo(ws.Cell("D1"));

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B10\" si=\"0\">A1*2</x:f>");
        for (var row = 1; row <= 10; row++)
            await Assert.That(f[$"D{row}"]).IsEqualTo($"<x:f>C{row}*2</x:f>");

        var expected = ExpectedColumn();
        for (var row = 1; row <= 10; row++)
            expected[$"D{row}"] = $"C{row}*2";
        await AssertReadsBack(saved, expected);
    }

    [Test]
    public async Task A_copied_sheet_reads_back_with_the_same_formulas()
    {
        using var wb = LoadWithSharedColumn();
        wb.Worksheet(1).CopyTo("Copy");

        var saved = Save(wb);

        await AssertValidSharedFormulas(saved);
        await AssertValidSharedFormulas(saved, "xl/worksheets/sheet2.xml");
        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        for (var row = 1; row <= 10; row++)
            await Assert.That(reloaded.Worksheet("Copy").Cell(row, 2).FormulaA1).IsEqualTo($"A{row}*2");
    }

    /// <summary>
    /// Excel refuses a file with a shared formula that refers to a sheet, and never writes one, so a
    /// shared formula from another producer that does is written as a formula in each cell.
    /// </summary>
    [Test]
    public async Task A_shared_formula_with_a_sheet_reference_is_written_on_its_own()
    {
        using var wb = LoadWithSharedColumn("Sheet1!A{0}*2");

        var saved = Save(wb);

        var f = FormulaElements(saved);
        for (var row = 1; row <= 10; row++)
            await Assert.That(f[$"B{row}"]).IsEqualTo($"<x:f>Sheet1!A{row}*2</x:f>");
    }

    /// <summary>
    /// A sheet rename that does not touch the formula's text leaves it shared. A rename that does
    /// touch it can only touch a formula with a sheet reference, which is never written as shared.
    /// </summary>
    [Test]
    public async Task A_sheet_rename_leaves_a_formula_without_a_sheet_reference_shared()
    {
        using var wb = LoadWithSharedColumn();
        wb.AddWorksheet("Other").Name = "Renamed";
        wb.Worksheet(1).Name = "First";

        var f = FormulaElements(Save(wb));

        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B10\" si=\"0\">A1*2</x:f>");
    }

    [Test]
    public async Task A_structured_reference_is_written_on_its_own()
    {
        using var wb = LoadWithSharedColumn("SUM(T[Val])+A{0}");

        var f = FormulaElements(Save(wb));

        await Assert.That(f["B1"]).IsEqualTo("<x:f>SUM(T[Val])+A1</x:f>");
        await Assert.That(f["B2"]).IsEqualTo("<x:f>SUM(T[Val])+A2</x:f>");
    }

    /// <summary>
    /// Excel shares a formula with an error literal such as <c>#DIV/0!</c>, or with <c>!</c> or
    /// <c>[</c> inside a string literal; neither is a sheet or a structured reference.
    /// </summary>
    [Test]
    public async Task An_error_literal_or_a_string_literal_keeps_the_formula_shared()
    {
        using var wb = LoadWithSharedColumn("IF(A{0}=0,#DIV/0!,\"[!\"&amp;A{0})");

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B10\" si=\"0\">IF(A1=0,#DIV/0!,\"[!\"&amp;A1)</x:f>");
        await Assert.That(f["B2"]).IsEqualTo("<x:f t=\"shared\" si=\"0\" />");
        await AssertReadsBack(saved, Enumerable.Range(1, 10).ToDictionary(r => $"B{r}", r => (string?)$"IF(A{r}=0,#DIV/0!,\"[!\"&A{r})"));
    }

    [Test]
    [Arguments("A1*2", true)]
    [Arguments("IF(A1=0,#DIV/0!,A1)", true)]
    [Arguments("IFERROR(A1,#value!)+#NULL!", true)]
    [Arguments("\"a!b[c\"&A1", true)]
    [Arguments("\"say \"\"hi!\"\"\"&A1", true)]
    [Arguments("MyName*A1", true)]
    [Arguments("Sheet2!A1", false)]
    [Arguments("'My Sheet'!A1", false)]
    [Arguments("SUM(Alpha:Beta!A1)", false)]
    [Arguments("[1]Sheet1!A1", false)]
    [Arguments("SUM(Table1[Val])", false)]
    [Arguments("SUM(Table1[])", false)]
    [Arguments("Sheet#REF!A1", false)]
    [Arguments("\"unterminated !", true)]
    public async Task Text_that_Excel_refuses_in_a_shared_formula_is_found(string formulaA1, bool shareable)
    {
        await Assert.That(SharedFormulaPlan.IsShareable(formulaA1)).IsEqualTo(shareable);
    }

    /// <summary>
    /// An array formula written inside a loaded shared formula keeps its own formula, and the shared
    /// formula keeps its <c>ref</c> around it, which Excel opens.
    /// </summary>
    [Test]
    public async Task An_array_formula_inside_the_ref_keeps_its_own_formula()
    {
        using var wb = LoadWithSharedColumn();
        wb.Worksheet(1).Range("B5:B6").FormulaArrayA1 = "A5:A6*3";

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B10\" si=\"0\">A1*2</x:f>");
        await Assert.That(f["B5"]).IsEqualTo("<x:f t=\"array\" ref=\"B5:B6\">A5:A6*3</x:f>");
        await Assert.That(f["B7"]).IsEqualTo("<x:f t=\"shared\" si=\"0\" />");
        await AssertValidSharedFormulas(saved);
    }

    /// <summary>
    /// Excel derives a later cell from the master's text and wraps a reference that runs off the
    /// sheet, where XLibur reads <c>#REF!</c>. Such a cell is written with its own text, and the
    /// formula starts at the next cell.
    /// </summary>
    [Test]
    public async Task A_cell_whose_text_is_REF_is_written_on_its_own()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        SetShared(ws, "R[-1]C", (1, 2), (2, 2), (3, 2));
        await Assert.That(ws.Cell(1, 2).FormulaA1).IsEqualTo("#REF!");

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f>#REF!</x:f>");
        await Assert.That(f["B2"]).IsEqualTo("<x:f t=\"shared\" ref=\"B2:B3\" si=\"0\">B1</x:f>");
        await Assert.That(f["B3"]).IsEqualTo("<x:f t=\"shared\" si=\"0\" />");
        await AssertReadsBack(saved, new Dictionary<string, string?> { ["B1"] = "#REF!", ["B2"] = "B1", ["B3"] = "B2" });
    }

    /// <summary>
    /// The master is the first cell in row order, so it is also the top-left cell of <c>ref</c>. A
    /// cell left of the master's column would be outside that <c>ref</c>, and is written on its own.
    /// </summary>
    [Test]
    public async Task A_cell_left_of_the_master_is_written_on_its_own()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        SetShared(ws, "RC[-1]*2", (1, 3), (2, 2), (2, 3), (3, 3));

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["C1"]).IsEqualTo("<x:f t=\"shared\" ref=\"C1:C3\" si=\"0\">B1*2</x:f>");
        await Assert.That(f["B2"]).IsEqualTo("<x:f>A2*2</x:f>");
        await Assert.That(f["C2"]).IsEqualTo("<x:f t=\"shared\" si=\"0\" />");
        await AssertReadsBack(saved, new Dictionary<string, string?>
        {
            ["C1"] = "B1*2", ["B2"] = "A2*2", ["C2"] = "B2*2", ["C3"] = "B3*2"
        });
    }

    [Test]
    public async Task Shared_formulas_are_numbered_from_zero_in_the_order_their_masters_are_written()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
        SetShared(ws, "RC[-1]*3", (1, 3), (2, 3));
        SetShared(ws, "RC[-1]*2", (1, 2), (2, 2));
        SetShared(ws, "RC[-1]*4", (5, 4), (6, 4));
        SetShared(ws, "RC[-1]*5", (9, 5));

        var saved = Save(wb);

        var f = FormulaElements(saved);
        await Assert.That(f["B1"]).IsEqualTo("<x:f t=\"shared\" ref=\"B1:B2\" si=\"0\">A1*2</x:f>");
        await Assert.That(f["C1"]).IsEqualTo("<x:f t=\"shared\" ref=\"C1:C2\" si=\"1\">B1*3</x:f>");
        await Assert.That(f["D5"]).IsEqualTo("<x:f t=\"shared\" ref=\"D5:D6\" si=\"2\">C5*4</x:f>");
        await Assert.That(f["E9"]).IsEqualTo("<x:f>D9*5</x:f>");
        await AssertValidSharedFormulas(saved);
    }

    [Test]
    public async Task Shared_formulas_are_numbered_on_each_sheet_from_zero()
    {
        using var wb = LoadWithSharedColumn();
        var second = (XLWorksheet)wb.AddWorksheet("Second");
        SetShared(second, "RC[-1]*2", (1, 2), (2, 2));

        var saved = Save(wb);

        await Assert.That(FormulaElements(saved)["B1"])
            .IsEqualTo("<x:f t=\"shared\" ref=\"B1:B10\" si=\"0\">A1*2</x:f>");
        await Assert.That(FormulaElements(saved, "xl/worksheets/sheet2.xml")["B1"])
            .IsEqualTo("<x:f t=\"shared\" ref=\"B1:B2\" si=\"0\">A1*2</x:f>");
    }

    /// <summary>A file of Excel's own, from LibreOffice's tests: every formula reads back the same.</summary>
    [Test]
    [Arguments(@"TryToLoad\LO\xlsx\shared_formula\basic.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\shared_formula\3d-reference.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\functions-excel-2010.xlsx")]
    [Arguments(@"TryToLoad\TemplateWithTableSourcePivotTables.xlsx")]
    public async Task A_round_trip_of_a_file_from_Excel_reads_back_the_same_formulas(string resource)
    {
        using var input = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource));
        using var wb = new XLWorkbook(input);
        var expected = AllFormulas(wb);

        var saved = Save(wb);

        foreach (var part in saved.PartsUnder("xl/worksheets/sheet"))
            await AssertValidSharedFormulas(saved, part);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var actual = AllFormulas(reloaded);
        await Assert.That(actual.Count).IsEqualTo(expected.Count);
        foreach (var (cell, formula) in expected)
            await Assert.That(actual[cell]).IsEqualTo(formula).Because(cell);

        await Assert.That(saved.PartsUnder("xl/worksheets/sheet").Any(p => saved.ReadPart(p).Contains("t=\"shared\"", StringComparison.Ordinal)))
            .IsTrue();
    }

    /// <summary>
    /// A1:A10 hold numbers, and B1:B10 hold <paramref name="format"/> filled down, saved as one shared
    /// formula as Excel saves it, then loaded.
    /// </summary>
    private static XLWorkbook LoadWithSharedColumn(string format = "A{0}*2")
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            for (var row = 1; row <= 20; row++)
                ws.Cell(row, 1).Value = row;

            for (var row = 1; row <= 10; row++)
                ws.Cell(row, 2).FormulaA1 = $"ROW()+{row}";

            wb.SaveAs(package);
        }

        package.RewriteSheet1(xml =>
        {
            for (var row = 1; row <= 10; row++)
            {
                var replacement = row == 1
                    ? $"<x:f t=\"shared\" ref=\"B1:B10\" si=\"0\">{string.Format(format, 1)}</x:f>"
                    : "<x:f t=\"shared\" si=\"0\"/>";
                xml = xml.Replace($"<x:f>ROW()+{row}</x:f>", replacement, StringComparison.Ordinal);
            }

            return xml;
        });

        package.Position = 0;
        return new XLWorkbook(package);
    }

    /// <summary>Put one shared formula in the given cells, as the loader does.</summary>
    private static void SetShared(XLWorksheet sheet, string r1c1, params (int Row, int Column)[] cells)
    {
        var shared = new WorksheetSheetDataReader.SharedFormula(r1c1);
        foreach (var (row, column) in cells)
        {
            var point = new Point(row, column);
            var formula = XLCellFormula.NormalA1(shared.ToA1(point));
            formula.SetShared(shared.Group, point);
            ((XLCell)sheet.Cell(row, column)).Formula = formula;
        }
    }

    private static Dictionary<string, string?> ExpectedColumn(params (string Cell, string? Formula)[] changes)
    {
        var expected = new Dictionary<string, string?>();
        for (var row = 1; row <= 10; row++)
            expected[$"B{row}"] = $"A{row}*2";

        foreach (var (cell, formula) in changes)
            expected[cell] = formula;

        return expected;
    }

    private static MemoryStream Save(XLWorkbook wb, bool shared = true)
    {
        var package = new MemoryStream();
        wb.SaveAs(package, new SaveOptions { WriteSharedFormulas = shared });
        return package;
    }

    /// <summary>The <c>f</c> element of each formula cell of a sheet, by cell.</summary>
    private static Dictionary<string, string> FormulaElements(MemoryStream package,
        string part = "xl/worksheets/sheet1.xml")
    {
        return FormulaCell.Matches(package.ReadPart(part))
            .ToDictionary(m => m.Groups[1].Value, m => m.Groups[2].Value);
    }

    /// <summary>
    /// The rules Excel reads shared formulas by, measured over COM: the master comes first and holds
    /// the text, a <c>ref</c> and the <c>si</c>; every later cell is inside its master's <c>ref</c>.
    /// </summary>
    private static async Task AssertValidSharedFormulas(MemoryStream package, string part = "xl/worksheets/sheet1.xml")
    {
        var masters = new Dictionary<string, Area>();
        foreach (Match match in FormulaCell.Matches(package.ReadPart(part)))
        {
            var cell = Point.Parse(match.Groups[1].Value);
            var element = match.Groups[2].Value;
            var startTag = element[..(element.IndexOf('>') + 1)];
            var attributes = Attribute.Matches(startTag).ToDictionary(m => m.Groups[1].Value, m => m.Groups[2].Value);
            if (attributes.GetValueOrDefault("t") != "shared")
                continue;

            var si = attributes["si"];
            var hasText = !startTag.EndsWith("/>", StringComparison.Ordinal);
            if (attributes.TryGetValue("ref", out var range))
            {
                await Assert.That(masters.ContainsKey(si)).IsFalse().Because($"si {si} has two masters");
                await Assert.That(hasText).IsTrue().Because($"master {match.Groups[1].Value} has no text");
                var area = Area.Parse(range);
                await Assert.That(area.FirstPoint).IsEqualTo(cell).Because("the master is the top-left cell of ref");
                masters.Add(si, area);
            }
            else
            {
                await Assert.That(masters.TryGetValue(si, out var area)).IsTrue()
                    .Because($"{match.Groups[1].Value} comes before the master of si {si}");
                await Assert.That(area.Contains(cell)).IsTrue()
                    .Because($"{match.Groups[1].Value} is outside ref {area} of si {si}");
            }
        }
    }

    /// <summary>
    /// The saved file keeps Excel's rules, and reads back with these formulas in these cells of the
    /// first sheet, and no others. A null formula is a cell without one.
    /// </summary>
    private static async Task AssertReadsBack(MemoryStream package, Dictionary<string, string?> expected)
    {
        await AssertValidSharedFormulas(package);
        package.Position = 0;
        using var reloaded = new XLWorkbook(package);
        var ws = reloaded.Worksheet(1);
        foreach (var (cell, formula) in expected)
        {
            if (formula is null)
                await Assert.That(ws.Cell(cell).HasFormula).IsFalse().Because(cell);
            else
                await Assert.That(ws.Cell(cell).FormulaA1).IsEqualTo(formula).Because(cell);
        }

        var formulaCells = ws.CellsUsed(c => c.HasFormula).Select(c => c.Address.ToString()!).ToHashSet();
        await Assert.That(formulaCells.SetEquals(expected.Where(e => e.Value is not null).Select(e => e.Key))).IsTrue();
    }

    private static Dictionary<string, string> AllFormulas(XLWorkbook wb)
    {
        var formulas = new Dictionary<string, string>();
        foreach (var ws in wb.Worksheets)
        {
            foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.All, c => c.HasFormula))
                formulas[$"{ws.Name}!{cell.Address}"] = cell.FormulaA1;
        }

        return formulas;
    }
}
