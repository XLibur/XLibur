using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;
using XLibur.Excel.IO;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// #542. The loader writes the A1 text of each cell of a shared formula from a template of the
/// formula's R1C1 text, which it parses once. The text must be the one
/// <see cref="FormulaText.TryConvert"/> gives in that cell, character for character, including a
/// reference that the move puts off the sheet. A text the parser refuses must be refused with the
/// same message.
/// </summary>
public class A1TemplateTests
{
    /// <summary>
    /// The corners and edges of a sheet, where a relative reference moves off it, the cells where a
    /// column gets a second and a third letter, and a few cells inside.
    /// </summary>
    private static readonly Point[] Anchors =
    [
        new(1, 1), new(1, 2), new(2, 1), new(2, 2), new(3, 3), new(1, 16384), new(2, 16383),
        new(1048576, 1), new(1048575, 2), new(1048576, 16384), new(1048575, 16383),
        new(26, 26), new(27, 27), new(702, 702), new(703, 703), new(1000, 16000), new(524288, 8192),
    ];

    /// <summary>The cells an A1 text is converted to R1C1 at, as the first cell of a shared formula.</summary>
    private static readonly Point[] Origins = [new(1, 1), new(2, 2), new(500, 30), new(1048576, 16384)];

    [Test]
    [MethodDataSource(nameof(R1C1Texts))]
    public async Task Template_writes_the_text_TryConvert_writes_in_each_cell(string r1c1)
    {
        await Assert.That(Mismatches(r1c1, Anchors)).IsEmpty();
    }

    /// <summary>
    /// The loader's own path: the A1 text of the first cell of a shared formula is converted to R1C1
    /// there, and the template of that text writes the other cells.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(A1Texts))]
    public async Task Template_of_A1_text_converted_at_a_first_cell_writes_the_text_TryConvert_writes(string a1)
    {
        var mismatches = new List<string>();
        foreach (var origin in Origins)
        {
            if (!FormulaText.TryConvert(a1, origin, FormulaNotation.R1C1, out var r1c1, out var refusal))
            {
                mismatches.Add($"'{a1}' at {Name(origin)}: TryConvert refused it: {refusal.Message}");
                continue;
            }

            mismatches.AddRange(Mismatches(r1c1, Anchors));
        }

        await Assert.That(mismatches).IsEmpty();
    }

    /// <summary>
    /// Every row of the formula-text corpus that converts to R1C1: each syntax form spec 54 collected.
    /// </summary>
    [Test]
    public async Task Template_writes_the_text_TryConvert_writes_for_every_corpus_formula()
    {
        var mismatches = new List<string>();
        var compared = 0;
        foreach (var row in FormulaTextCorpusTests.Rows().Select(x => x()))
        {
            if (!FormulaText.TryConvert(row.Text, new Point(3, 3), FormulaNotation.R1C1, out var r1c1, out _))
                continue;

            compared++;
            mismatches.AddRange(Mismatches(r1c1, Anchors).Select(m => $"{row.Form}: {m}"));
        }

        await Assert.That(compared).IsGreaterThan(10);
        await Assert.That(mismatches).IsEmpty();
    }

    /// <summary>
    /// Every shared formula that a workbook in the test resources holds, in each of its cells as the
    /// file places them, and at the edges of the sheet.
    /// </summary>
    [Test]
    public async Task Template_writes_the_text_TryConvert_writes_for_every_shared_formula_in_the_test_resources()
    {
        var mismatches = new List<string>();
        var formulas = 0;
        var cells = 0;
        foreach (var resource in TestHelper.ListResourceFiles(IsWorkbook))
        {
            foreach (var (text, formulaCells) in SharedFormulas(resource))
            {
                // The loader converts the text of the first cell to R1C1 there, and throws when that fails.
                if (!FormulaText.TryConvert(text, formulaCells[0], FormulaNotation.R1C1, out var r1c1, out _))
                    continue;

                formulas++;
                cells += formulaCells.Count;
                mismatches.AddRange(Mismatches(r1c1, formulaCells.Concat(Anchors)).Select(m => $"{resource}: {m}"));
            }
        }

        Console.WriteLine($"Compared {formulas} shared formulas with {cells} cells.");
        await Assert.That(formulas).IsGreaterThan(0);
        await Assert.That(mismatches).IsEmpty();
    }

    /// <summary>
    /// The loader throws the refusal from the second cell of a shared formula, where
    /// <see cref="FormulaText.TryConvert"/> refused the text before. It must be the same refusal.
    /// </summary>
    [Test]
    [Arguments("1+")]
    [Arguments("SUM(RC")]
    [Arguments("RC)")]
    [Arguments("")]
    [Arguments("   ")]
    [Arguments("R1C1:")]
    [Arguments("\tR1C1")]
    [Arguments("'[Book2.xlsx]Sheet1'!RC")]
    public async Task Template_refuses_the_text_TryConvert_refuses_with_the_same_message(string r1c1)
    {
        await Assert.That(FormulaText.TryConvert(r1c1, new Point(5, 5), FormulaNotation.A1, out _, out var expected))
            .IsFalse();

        await Assert.That(FormulaText.TryCreateA1Template(r1c1, out var template, out var refusal)).IsFalse();
        await Assert.That(template).IsNull();
        await Assert.That(refusal.Text).IsEqualTo(r1c1);
        await Assert.That(refusal.Message).IsEqualTo(expected.Message);
    }

    /// <summary>
    /// A cell called as a function, as on a macro sheet, is written as <c>#REF!</c> with all its
    /// arguments when the cell is off the sheet, so it is no slot of its own. The template does not
    /// take it, and a shared formula converts each cell instead.
    /// </summary>
    [Test]
    public async Task Shared_formula_that_calls_a_cell_writes_the_text_TryConvert_writes()
    {
        const string r1c1 = "R[-1]C(RC[1])";
        await Assert.That(FormulaText.TryCreateA1Template(r1c1, out var template, out _)).IsTrue();
        await Assert.That(template).IsNull();

        var formula = new WorksheetSheetDataReader.SharedFormula(r1c1);
        var mismatches = new List<string>();
        foreach (var anchor in Anchors)
        {
            FormulaText.TryConvert(r1c1, anchor, FormulaNotation.A1, out var expected, out _);
            var actual = formula.ToA1(anchor);
            if (actual != expected)
                mismatches.Add($"{Name(anchor)}: TryConvert '{expected}', shared formula '{actual}'");
        }

        await Assert.That(mismatches).IsEmpty();
    }

    /// <summary>
    /// A loaded shared formula: each cell gets the A1 text of its own cell, a reference that the move
    /// puts off the sheet is <c>#REF!</c>, and each cell keeps the one R1C1 string of its formula for
    /// the dependency tree (#513).
    /// </summary>
    [Test]
    public async Task Loaded_shared_formula_cells_get_their_own_A1_text_and_keep_the_R1C1_text_of_the_formula()
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            wb.AddWorksheet("Sheet 2");
            ws.Cell("C1").FormulaA1 = "A1*$B$1+'Sheet 2'!A1";
            ws.Cell("C2").FormulaA1 = "1+2";
            ws.Cell("C3").FormulaA1 = "1+3";
            ws.Cell("XFC1").FormulaA1 = "XFD1*2";
            ws.Cell("XFD1").FormulaA1 = "1+4";
            ws.Cell("A1048575").FormulaA1 = "SUM(B1048576:C1048576)";
            ws.Cell("A1048576").FormulaA1 = "1+5";
            wb.SaveAs(package);
        }

        package.RewriteSheet1(xml =>
        {
            xml = Share(xml, "A1*$B$1+'Sheet 2'!A1", "<x:f t=\"shared\" ref=\"C1:C3\" si=\"0\">A1*$B$1+'Sheet 2'!A1</x:f>");
            xml = Share(xml, "1+2", "<x:f t=\"shared\" si=\"0\" />");
            xml = Share(xml, "1+3", "<x:f t=\"shared\" si=\"0\" />");
            xml = Share(xml, "XFD1*2", "<x:f t=\"shared\" ref=\"XFC1:XFD1\" si=\"1\">XFD1*2</x:f>");
            xml = Share(xml, "1+4", "<x:f t=\"shared\" si=\"1\" />");
            xml = Share(xml, "SUM(B1048576:C1048576)",
                "<x:f t=\"shared\" ref=\"A1048575:A1048576\" si=\"2\">SUM(B1048576:C1048576)</x:f>");
            return Share(xml, "1+5", "<x:f t=\"shared\" si=\"2\" />");
        });

        package.Position = 0;
        using var loaded = new XLWorkbook(package);
        var sheet = loaded.Worksheet("Sheet1");
        (string Cell, string A1, string Group)[] expected =
        [
            ("C1", "A1*$B$1+'Sheet 2'!A1", "C1"),
            ("C2", "A2*$B$1+'Sheet 2'!A2", "C1"),
            ("C3", "A3*$B$1+'Sheet 2'!A3", "C1"),
            ("XFC1", "XFD1*2", "XFC1"),
            ("XFD1", "#REF!*2", "XFC1"),
            ("A1048575", "SUM(B1048576:C1048576)", "A1048575"),
            ("A1048576", "SUM(#REF!)", "A1048575"),
        ];

        foreach (var (address, a1, group) in expected)
        {
            var cell = (XLCell)sheet.Cell(address);
            await Assert.That(cell.FormulaA1).IsEqualTo(a1);

            var first = (XLCell)sheet.Cell(group);
            await Assert.That(first.Formula!.TryGetSharedR1C1(first.SheetPoint, out var groupR1C1)).IsTrue();
            await Assert.That(cell.Formula!.TryGetSharedR1C1(cell.SheetPoint, out var r1c1)).IsTrue();
            await Assert.That(r1c1).IsSameReferenceAs(groupR1C1);
        }
    }

    /// <summary>
    /// #557. The template put the colon placeholder back over the whole text it writes, so a
    /// fullwidth colon (U+FF1A) the formula really had became an ordinary colon: in a quoted sheet
    /// name, which the template writes as a slot's prefix, and in a string, which it keeps as a
    /// literal. <see cref="FormulaText.TryConvert"/> did the same, so comparing the two hid it.
    /// </summary>
    [Test]
    [Arguments("Table1[a:b]&'x：y'!RC", "Table1[a:b]&x：y!C3")]
    [Arguments("Table1[a:b]&\"x：y\"&RC", "Table1[a:b]&\"x：y\"&C3")]
    [Arguments("Table1[a:b]+Table1[x：y]", "Table1[a:b]+Table1[x：y]")]
    [Arguments("Table1[a:b：c]&RC", "Table1[a:b：c]&C3")]
    public async Task Issue557_template_keeps_a_fullwidth_colon_the_formula_really_had(string r1c1, string a1)
    {
        await Assert.That(FormulaText.TryCreateA1Template(r1c1, out var template, out _)).IsTrue();
        await Assert.That(template).IsNotNull();

        await Assert.That(template!.ToA1(new Point(3, 3))).IsEqualTo(a1);
    }

    public static IEnumerable<object[]> R1C1Texts()
    {
        string[] texts =
        [
            // Single cells: relative, absolute and mixed, and the furthest a relative one reaches.
            "RC", "R[1]C[1]", "R[-1]C[-1]", "R1C1", "R[1]C1", "R1C[1]", "R[-1]C1", "R1C[-1]", "R[5]C[-3]",
            "R1048576C16384", "R[1048575]C[16383]", "R[-1048575]C[-16383]",

            // Areas, including one whose ends are the same cell.
            "R1C1:R2C2", "RC:R[1]C[1]", "R[-2]C:RC", "R1C[-1]:R[1]C1", "R[-1]C[-1]:R[1]C[1]", "R1C1:R1C1",
            "RC:RC", "R[1]C:R[1]C", "R[1]C[1]:RC",

            // Whole rows and whole columns.
            "R", "R1", "R[1]", "R[-1]", "R1:R3", "R[-1]:R[1]", "R1:R[2]", "R2:R2", "R[1]:R[1]",
            "C", "C1", "C[1]", "C[-1]", "C1:C3", "C[-1]:C[2]", "C2:C2", "SUM(C[-1])", "SUM(R[1]:R[2])",

            // Another sheet, a 3D reference, another workbook, and a bang reference.
            "Sheet2!RC", "Sheet2!R[-1]C:R[1]C", "'My Sheet'!R[1]C", "'It''s'!R1C1:R[2]C", "'Sheet1'!RC[-1]",
            "'1Sheet'!RC", "'R1C1'!RC", "'A1'!R[1]C", "Sheet2!R[1]", "'My Sheet'!C[-1]:C[1]",
            "Sheet1:Sheet3!RC[1]", "'Sheet 1:Sheet 3'!R[-1]C", "Sheet1:Sheet3!R1C1:R[1]C[1]",
            "[1]Sheet1!RC", "[1]Sheet1!R[-1]C[-1]:R1C1", "[1]Sheet1:Sheet2!R1C[1]", "'[1]My Sheet'!RC",
            "!RC[1]", "!R[-1]C", "!R1C1:R2C2", "Sheet2!RC+Sheet2!R[1]C+'My Sheet'!RC+Sheet2!R[2]C",
            "'Sheet：1'!RC",

            // Errors in place of a reference.
            "Sheet2!#REF!", "#REF!", "#REF!+RC", "Sheet2!#REF!+RC", "[1]Sheet1!#REF!+RC",

            // Names.
            "MyName*2", "Sheet1!MyName+RC", "[1]!Name", "[0]!Total+R[-1]C", "!Local*RC", "'My Sheet'!Local+RC",

            // Structured references, one with a colon in a column name.
            "Table1[Col]", "Table1[[#This Row],[Col]]*RC", "Table1[Start: Date]+RC[1]",
            "SUM(Table1[[Col1]:[Col2]])+R[1]C", "Table1[[#Headers],[Start: Date]]&RC",
            "Table1[Start: Date]+'Sheet：1'!RC",

            // Text that looks like a reference.
            "\"R1C1\"&RC", "\"A1\"&R[1]C", "\"RC\"", "\"a：b\"&RC", "Table1[Start: Date]&\"a：b\"&RC[-1]",
            "\"'Sheet1'!R1C1\"&R[-1]C", "\"\"\"R1C1\"\"\"&RC",

            // Functions.
            "SUM(RC[-3]:RC[-1])", "IF(RC[-1]>100,SUM(RC[-7]:RC[-1]),RC[-1]/2)", "sum( R1C1 , R[1]C )",
            "_xlfn.CONCAT(RC,\"x\")", "SUM(Sheet2!RC,Sheet3!R[-1]C)", "INDEX(C[-1],ROW())",
            "ROUND(R[-1]C/RC[1],2)", "IF(ISERROR(RC[-1]),0,RC[-1]*R1C7)", "NOW()", "PI()*2", "SUM()",
            "Sheet2!MyFunc(RC)", "[1]!MyFunc(RC)", "[1]Sheet1!MyFunc(RC)", "SUM((RC,R[1]C))",

            // Arrays and errors.
            "{1,2;3,4}", "SUM({1,2}*RC)", "{\"R1C1\",#N/A}", "SUM(RC*{1,2,3})",
            "#N/A", "#DIV/0!+RC", "IFERROR(RC,#VALUE!)",

            // Whitespace, operators and parentheses.
            " RC", "RC ", "  R[1]C  +  R[2]C  ", "RC\n", "RC\t", "RC R1C1:R2C2", "( RC )", "RC:R[1]C:R[2]C",
            "-RC", "+R[1]C%", "--RC", "RC%", "R[1]C^2", "((RC))", "(RC+R[1]C)*2", "RC:R[1]C[1] R1C1:R10C10",

            // No reference at all.
            "TRUE+1", "1+2", "\"text\"", "1", "Sdemo123|tik!'id1?req?AAPL'",

            // Longer than the stack buffer, so the text is written in a rented one.
            string.Join("+", Enumerable.Range(1, 60).Select(i => $"R[{i}]C[{i % 40 - 20}]")),
        ];

        return texts.Select(text => new object[] { text });
    }

    public static IEnumerable<object[]> A1Texts()
    {
        string[] texts =
        [
            "A1", "$A$1", "A$1", "$A1", "A1:B2", "$A$1:B2", "B2:A1", "A:A", "$A:$B", "A:$C", "1:1", "$1:$3",
            "2:$5", "XFD1048576", "A1048576", "XFD1", "$XFD$1048576", "Sheet2!A1", "'My Sheet'!$B$2:C3",
            "Sheet1:Sheet3!A1", "[1]Sheet1!A1", "!A1", "SUM(A1:C1)*$G$1", "A1*B1+1",
            "IF(H1>100,SUM(A1:G1),H1/2)", "AVERAGE(B1:F1)*C1", "ROUND(D1/E1,2)", "MAX(A1,B1,C1)-MIN(D1,E1)",
            "IF(ISERROR(N1),0,N1*$G$1)", "A1:C3 B2:D4", "(A1,B2)", "@A1:A3", "A1#", "Table1[Start: Date]+A1",
            "INDIRECT(\"A1\")&B1", "SUM(Sheet2!A:A)", "'Sheet 1'!1:1", "OFFSET(A1,1,1)",
        ];

        return texts.Select(text => new object[] { text });
    }

    /// <summary>
    /// Each cell at which the template's text differs from the text <see cref="FormulaText.TryConvert"/>
    /// gives, as a message. Empty when they agree in every cell.
    /// </summary>
    private static List<string> Mismatches(string r1c1, IEnumerable<Point> cells)
    {
        var mismatches = new List<string>();
        if (!FormulaText.TryCreateA1Template(r1c1, out var template, out var refusal))
        {
            mismatches.Add($"'{r1c1}': the template refused it: {refusal.Message}");
            return mismatches;
        }

        if (template is null)
        {
            mismatches.Add($"'{r1c1}': no template");
            return mismatches;
        }

        foreach (var cell in cells)
        {
            if (!FormulaText.TryConvert(r1c1, cell, FormulaNotation.A1, out var expected, out var convertRefusal))
            {
                mismatches.Add($"'{r1c1}' at {Name(cell)}: TryConvert refused it: {convertRefusal.Message}");
                continue;
            }

            var actual = template.ToA1(cell);
            if (actual != expected)
                mismatches.Add($"'{r1c1}' at {Name(cell)}: TryConvert '{expected}', template '{actual}'");
        }

        return mismatches;
    }

    private static string Name(Point cell) => $"R{cell.Row}C{cell.Column}";

    private static bool IsWorkbook(string resource)
        => resource.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
           resource.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The shared formulas of every sheet of a workbook resource: the text of each one's first cell,
    /// and its cells in the order of the file, the first cell first, as the loader reads them.
    /// </summary>
    private static List<(string Text, List<Point> Cells)> SharedFormulas(string resource)
    {
        var formulas = new List<(string Text, List<Point> Cells)>();
        using var stream = TestHelper.GetStreamFromResource(resource);
        ZipArchive archive;
        try
        {
            archive = new ZipArchive(stream, ZipArchiveMode.Read);
        }
        catch (InvalidDataException)
        {
            // An encrypted workbook is not a zip package.
            return formulas;
        }

        using (archive)
        {
            foreach (var entry in archive.Entries)
            {
                if (!entry.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) ||
                    !entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    continue;

                using var part = entry.Open();
                formulas.AddRange(ReadSharedFormulas(part));
            }
        }

        return formulas;
    }

    private static IEnumerable<(string Text, List<Point> Cells)> ReadSharedFormulas(Stream part)
    {
        var byIndex = new Dictionary<string, (string Text, List<Point> Cells)>(StringComparer.Ordinal);
        using var reader = XmlReader.Create(part);
        Point? cell = null;
        while (!reader.EOF)
        {
            if (IsElement(reader, "c"))
            {
                cell = ReadCellReference(reader);
                reader.Read();
            }
            else if (IsElement(reader, "f") && reader.GetAttribute("t") == "shared" && cell is { } at)
            {
                AddSharedFormula(reader, byIndex, at);
            }
            else
            {
                reader.Read();
            }
        }

        return byIndex.Values;
    }

    private static bool IsElement(XmlReader reader, string localName) =>
        reader.NodeType == XmlNodeType.Element && reader.LocalName == localName;

    private static Point? ReadCellReference(XmlReader reader)
    {
        var reference = reader.GetAttribute("r");
        return reference is not null && Point.TryParse(reference, out var point) ? point : null;
    }

    /// <summary>Consumes a shared formula element and records <paramref name="at"/> under its index.</summary>
    private static void AddSharedFormula(
        XmlReader reader,
        Dictionary<string, (string Text, List<Point> Cells)> byIndex,
        Point at)
    {
        var index = reader.GetAttribute("si");
        var text = reader.ReadElementContentAsString();
        if (index is null)
            return;

        if (byIndex.TryGetValue(index, out var formula))
            formula.Cells.Add(at);
        else
            byIndex.Add(index, (text, [at]));
    }

    /// <summary>Replace the formula element that holds <paramref name="text"/> with <paramref name="shared"/>.</summary>
    private static string Share(string sheetXml, string text, string shared)
    {
        var original = $"<x:f>{text}</x:f>";
        var rewritten = sheetXml.Replace(original, shared, StringComparison.Ordinal);
        if (ReferenceEquals(rewritten, sheetXml))
            throw new InvalidOperationException($"'{original}' was not found in the sheet part.");

        return rewritten;
    }
}
