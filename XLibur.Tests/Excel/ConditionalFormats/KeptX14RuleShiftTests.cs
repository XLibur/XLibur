using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// A conditional format over a row or column insert or delete: its range and the references in its
/// formulas, for a rule XLibur models and for one it keeps in the worksheet's <c>x14</c> extension
/// without modelling it (issue #499, D77). Each differential test builds the same rule both ways on
/// one sheet, makes one edit, saves, and compares what the two rules are written with.
/// </summary>
public class KeptX14RuleShiftTests
{
    private const string Folder = @"Other\SheetLifecycle\";
    private const string RuleId = "{0A5E0C1D-0000-4000-8000-000000000499}";
    private const string Formula = "Data!$A$2>0";

    /// <summary>
    /// The expected range and formula pin the answer, so that the two kinds of rule cannot pass by
    /// being wrong the same way; each kind must also agree with the other. An empty string means the
    /// edit removed the rule. A formula's references shift as a cell formula's do (spec 25's shifter):
    /// a relative reference is written relative to the rule's own range, so it moves with the cell it
    /// names, and a formula the parser refuses keeps its text (ADR 0002).
    /// </summary>
    [Test]
    [Arguments("C3:C5", Formula, "insert rows above", "C5:C7", Formula)]
    [Arguments("C3:C5", Formula, "insert columns left", "E3:E5", Formula)]
    [Arguments("C3:C5", Formula, "insert a row inside", "C3:C6", Formula)]
    [Arguments("C3:E5", Formula, "insert a column inside", "C3:F5", Formula)]
    [Arguments("C3:C5", Formula, "delete a row inside", "C3:C4", Formula)]
    [Arguments("C3:C5", Formula, "delete the rows holding the range", "", "")]
    [Arguments("C3:C5", Formula, "delete the column holding the range", "", "")]
    [Arguments("C3:C5", Formula, "delete rows overlapping the bottom", "C3", Formula)]
    [Arguments("C3:C5", Formula, "delete rows overlapping the top", "C1:C2", Formula)]
    [Arguments("C3:C5", Formula, "delete a cell inside, shifting up", "C3:C4", Formula)]
    [Arguments("C3:C5", Formula, "insert a cell above, shifting down", "C4:C6", Formula)]
    [Arguments("C3:C5", Formula, "insert a cell above another column", "C3:C5", Formula)]
    [Arguments("C3:C5 E3:E5", Formula, "insert rows above", "C5:C7 E5:E7", Formula)]
    [Arguments("C1", Formula, "insert rows above", "C3", Formula)]
    [Arguments("C3:C5", Formula, "insert rows on the referenced sheet", "C3:C5", "Data!$A$3>0")]
    [Arguments("C3:C5", Formula, "insert a row above the referenced cell", "C3:C5", "Data!$A$3>0")]
    [Arguments("C3:C5", Formula, "insert a row below the referenced cell", "C3:C5", Formula)]
    [Arguments("C3:C5", Formula, "insert a column on the referenced sheet", "C3:C5", "Data!$B$2>0")]
    [Arguments("C3:C5", Formula, "delete a row above the referenced cell", "C3:C5", "Data!$A$1>0")]
    [Arguments("C3:C5", Formula, "delete the referenced row", "C3:C5", "Data!#REF!>0")]
    [Arguments("E1", "A1>0", "insert one row above", "E2", "A2>0")]
    [Arguments("E1", "A1>0", "insert columns left", "G1", "C1>0")]
    [Arguments("E3", "A2>0", "delete row 2", "E2", "#REF!>0")]
    [Arguments("E3:E5", "$A$1>E3", "insert a row inside", "E3:E6", "$A$1>E3")]
    [Arguments("C3:C5", "SUM(Data!A2", "insert a row above the referenced cell", "C3:C5", "SUM(Data!A2")]
    [Arguments("E3", "SUM(A2", "insert one row above", "E4", "SUM(A2")]
    public async Task Both_kinds_of_rule_shift_their_range_and_formula_alike(string sqref, string formula,
        string edit, string expectedRange, string expectedFormula)
    {
        var saved = EditAndSave(sqref, formula, edit);

        await Assert.That(saved.ModelledRange).IsEqualTo(expectedRange);
        await Assert.That(saved.KeptRange).IsEqualTo(saved.ModelledRange);
        await Assert.That(saved.ModelledFormula).IsEqualTo(expectedFormula);
        await Assert.That(saved.KeptFormula).IsEqualTo(saved.ModelledFormula);
    }

    /// <summary>
    /// A rule's formula shifts by the rule a cell formula does: the same text, with the same
    /// references, on the same sheet, gives the same answer. That includes a deleted reference, which
    /// the shifter writes as <c>Data!#REF!</c>, keeping the sheet it named.
    /// </summary>
    [Test]
    [Arguments("insert rows on the referenced sheet")]
    [Arguments("insert a row above the referenced cell")]
    [Arguments("insert a column on the referenced sheet")]
    [Arguments("delete a row above the referenced cell")]
    [Arguments("delete the referenced row")]
    public async Task A_rules_formula_shifts_as_a_cell_formula_does(string edit)
    {
        var saved = EditAndSave("C3:C5", Formula, edit);

        await Assert.That(saved.ModelledFormula).IsEqualTo(saved.CellFormula);
        await Assert.That(saved.KeptFormula).IsEqualTo(saved.CellFormula);
    }

    /// <summary>
    /// A value point whose type is <see cref="XLCFContentType.Formula"/> holds a formula without an
    /// <c>=</c>, in each kind of scale that has value points. It shifts like any other formula.
    /// </summary>
    [Test]
    [Arguments("colour scale")]
    [Arguments("data bar")]
    [Arguments("icon set")]
    public async Task A_formula_value_point_shifts(string kind)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Data").Cell("A2").Value = 1;
            var format = wb.AddWorksheet("Other").Range("C2:C4").AddConditionalFormat();
            switch (kind)
            {
                case "colour scale":
                    format.ColorScale()
                        .Minimum(XLCFContentType.Formula, "Data!$A$2", XLColor.Red)
                        .Maximum(XLCFContentType.Maximum, "0", XLColor.Blue);
                    break;
                case "data bar":
                    format.DataBar(XLColor.Red)
                        .Minimum(XLCFContentType.Formula, "Data!$A$2")
                        .Maximum(XLCFContentType.Maximum, "0");
                    break;
                default:
                    format.IconSet(XLIconSetStyle.ThreeArrows)
                        .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "0", XLCFContentType.Number)
                        .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "Data!$A$2", XLCFContentType.Formula)
                        .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, "Data!$A$2", XLCFContentType.Formula);
                    break;
            }

            wb.Worksheet("Data").Row(2).InsertRowsAbove(1);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var document = SpreadsheetDocument.Open(ms, false);
        var formulaPoints = OtherPart(document).Worksheet!.Elements<S.ConditionalFormatting>()
            .SelectMany(c => c.Descendants<S.ConditionalFormatValueObject>())
            .Where(v => v.Type?.Value == S.ConditionalFormatValueObjectValues.Formula)
            .Select(v => v.Val?.Value ?? string.Empty)
            .ToList();
        await Assert.That(formulaPoints).IsNotEmpty();
        await Assert.That(formulaPoints.Distinct().ToList()).IsEquivalentTo(new[] { "Data!$A$3" });
    }

    /// <summary>
    /// The rules Excel wrote in <c>rename-before.xlsx</c>: an expression on <c>Other!C1</c>, and a
    /// colour scale on <c>C2:C4</c> whose low point is a formula. A row inserted above row 1 of the
    /// sheet they are on moves both ranges down one row. Their formulas name another sheet, so they
    /// stay as they were, and so does every other byte of the extension.
    /// </summary>
    [Test]
    public async Task Excels_kept_rules_move_with_a_row_inserted_on_their_sheet()
    {
        var before = ExtensionXml(Resource("rename-before.xlsx"));

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("rename-before.xlsx")))
        {
            wb.Worksheet("Other").Row(1).InsertRowsAbove(1);
            wb.SaveAs(ms);
        }

        var expected = before
            .Replace("<xm:sqref>C1</xm:sqref>", "<xm:sqref>C2</xm:sqref>", StringComparison.Ordinal)
            .Replace("<xm:sqref>C2:C4</xm:sqref>", "<xm:sqref>C3:C5</xm:sqref>", StringComparison.Ordinal);
        await Assert.That(expected).IsNotEqualTo(before);
        await Assert.That(ExtensionXml(ms)).IsEqualTo(expected);
    }

    /// <summary>
    /// A row inserted or deleted on <c>Data</c>, the sheet the rules' formulas refer to, moves no
    /// range on <c>Other</c>, and shifts both formulas: the expression and the colour scale's value
    /// point. Every other byte of the extension stays as Excel wrote it.
    /// </summary>
    [Test]
    [Arguments("insert", "Data!$A$3")]
    [Arguments("delete", "Data!#REF!")]
    public async Task Excels_kept_rules_formulas_follow_an_edit_on_the_referenced_sheet(string edit, string reference)
    {
        var before = ExtensionXml(Resource("rename-before.xlsx"));

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("rename-before.xlsx")))
        {
            var row = wb.Worksheet("Data").Row(2);
            if (edit == "insert")
                row.InsertRowsAbove(1);
            else
                row.Delete();
            wb.SaveAs(ms);
        }

        var expected = before
            .Replace("<xm:f>Data!$A$2&gt;0</xm:f>", $"<xm:f>{reference}&gt;0</xm:f>", StringComparison.Ordinal)
            .Replace("<xm:f>Data!$A$2</xm:f>", $"<xm:f>{reference}</xm:f>", StringComparison.Ordinal);
        await Assert.That(expected).IsNotEqualTo(before);
        await Assert.That(ExtensionXml(ms)).IsEqualTo(expected);
        await Assert.That(KeptRanges(ms)).IsEquivalentTo(new[] { "C1", "C2:C4" });
    }

    /// <summary>
    /// A load and a save with no edit writes the extension back exactly as Excel wrote it.
    /// </summary>
    [Test]
    public async Task Excels_kept_rules_round_trip_unchanged()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("rename-before.xlsx")))
            wb.SaveAs(ms);

        await Assert.That(ExtensionXml(ms)).IsEqualTo(ExtensionXml(Resource("rename-before.xlsx")));
    }

    /// <summary>
    /// The conditional-format listener's worst input for kept ranges: a sheet with no rules at all,
    /// and one whose only rules are kept ones, holding a range already emptied, the whole sheet, a
    /// cell on the last row that an insert pushes off the sheet, and two areas at opposite edges.
    /// </summary>
    [Test]
    public async Task The_conditional_format_listener_does_not_throw_on_kept_ranges()
    {
        using var wb = new XLWorkbook();
        var empty = (XLWorksheet)wb.AddWorksheet("Empty");
        var host = (XLWorksheet)wb.AddWorksheet("Host");
        var formats = host.ConditionalFormats;
        formats.SeedExtensionRuleAreas("emptied", XLAreaList.Empty);
        formats.SeedExtensionRuleAreas("sheet", new XLAreaList(Area.Parse("A1:XFD1048576")));
        formats.SeedExtensionRuleAreas("last row", new XLAreaList(Area.Parse("C1048576")));
        formats.SeedExtensionRuleAreas("edges", new XLAreaList([Area.Parse("A1"), Area.Parse("XFD1")]));

        foreach (var sheet in new[] { empty, host })
            await AssertEveryEditThrowsNothing(sheet.ConditionalFormats, sheet);

        await Assert.That(Kept(formats, "emptied")).IsEqualTo("");
        await Assert.That(Kept(formats, "sheet")).IsEqualTo("A1:XFD1048576");
        await Assert.That(Kept(formats, "last row")).IsEqualTo("");
        await Assert.That(Kept(formats, "edges")).IsEqualTo("A1");
    }

    /// <summary>
    /// The listener each sheet's conditional formats now register (one per sheet, as data validations
    /// do), handed edits on its own sheet and on another. Worst input: a sheet with no rules; an
    /// expression the parser refuses; one already <c>#REF!</c>; a scale's formula value point; a
    /// text value that only looks like a reference; a rule with no values; kept formula text the
    /// parser refuses; kept text that is empty; and a kept rule whose range an edit removed.
    /// </summary>
    [Test]
    public async Task The_per_sheet_conditional_format_listener_does_not_throw_on_formulas()
    {
        using var wb = new XLWorkbook();
        var data = (XLWorksheet)wb.AddWorksheet("Data");
        var empty = (XLWorksheet)wb.AddWorksheet("Empty");
        IXLWorksheet sheet = wb.AddWorksheet("Host");
        var host = (XLWorksheet)sheet;
        sheet.Range("A1:A2").AddConditionalFormat().WhenIsTrue("=SUM(Data!A1").Fill.SetBackgroundColor(XLColor.Red);
        sheet.Range("B1:B2").AddConditionalFormat().WhenIsTrue("=#REF!>0").Fill.SetBackgroundColor(XLColor.Red);
        sheet.Range("C1:C3").AddConditionalFormat().ColorScale()
            .Minimum(XLCFContentType.Formula, "Data!$A$5", XLColor.Red)
            .Maximum(XLCFContentType.Maximum, "0", XLColor.Blue);
        sheet.Range("D1").AddConditionalFormat().WhenContains("Data!$A$5").Fill.SetBackgroundColor(XLColor.Red);
        sheet.Range("E1").AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(XLColor.Red);
        var formats = host.ConditionalFormats;
        formats.SeedExtensionRuleFormulas("refused", ["SUM(Data!A1", "Data!#REF!"]);
        formats.SeedExtensionRuleFormulas("empty", [""]);
        formats.SeedExtensionRuleFormulas("removed", ["Data!$A$5"]);
        formats.SeedExtensionRuleAreas("removed", XLAreaList.Empty);

        // An insert on Data above row 5 first, which the listener must act on, and then every edit on
        // both sheets, handed to both sheets' listeners.
        ISheetListener hostListener = formats;
        hostListener.OnInsertAreaAndShiftDown(Edit(data, data.Range(1, 1, 1, XLHelper.MaxColumnNumber), 1));
        await Assert.That(ScaleMinimum(host)).IsEqualTo("Data!$A$6");

        foreach (var edited in new[] { data, host })
        {
            foreach (var listening in new[] { empty, host })
                await AssertEveryEditThrowsNothing(listening.ConditionalFormats, edited);
        }

        // A refused formula is never rewritten (ADR 0002); text that is not a formula is not either;
        // and a removed kept rule is left alone.
        var cfs = host.ConditionalFormats.Cast<XLConditionalFormat>().ToList();
        await Assert.That(cfs[0].Values[1].Value).IsEqualTo("SUM(Data!A1");
        await Assert.That(cfs[1].Values[1].Value).IsEqualTo("#REF!>0");
        await Assert.That(cfs[3].Values[1].Value).IsEqualTo("Data!$A$5");
        await Assert.That(formats.TryGetExtensionRuleFormulas("refused", out var refused)).IsTrue();
        await Assert.That(refused).IsEquivalentTo(new[] { "SUM(Data!A1", "Data!#REF!" }, CollectionOrdering.Matching);
        await Assert.That(formats.TryGetExtensionRuleFormulas("removed", out var removed)).IsTrue();
        await Assert.That(removed).IsEquivalentTo(new[] { "Data!$A$5" }, CollectionOrdering.Matching);
    }

    private static string? ScaleMinimum(XLWorksheet host)
        => host.ConditionalFormats.Cast<XLConditionalFormat>()
            .Single(c => c.ConditionalFormatType == XLConditionalFormatType.ColorScale).Values[1].Value;

    /// <summary>
    /// Hands <paramref name="listener"/> a whole-row and a whole-column insert and delete at the
    /// first line of <paramref name="edited"/>, asserting that none of them throws.
    /// </summary>
    private static async Task AssertEveryEditThrowsNothing(ISheetListener listener, XLWorksheet edited)
    {
        var row = edited.Range(1, 1, 1, XLHelper.MaxColumnNumber);
        var column = edited.Range(1, 1, XLHelper.MaxRowNumber, 1);

        await Assert.That(() => listener.OnInsertAreaAndShiftDown(Edit(edited, row, 1))).ThrowsNothing();
        await Assert.That(() => listener.OnInsertAreaAndShiftRight(Edit(edited, column, 1))).ThrowsNothing();
        await Assert.That(() => listener.OnDeleteAreaAndShiftUp(Edit(edited, row, -1))).ThrowsNothing();
        await Assert.That(() => listener.OnDeleteAreaAndShiftLeft(Edit(edited, column, -1))).ThrowsNothing();
    }

    private static SheetEdit Edit(XLWorksheet sheet, XLRange range, int shift) => new()
    {
        Sheet = sheet,
        Area = Area.FromRangeAddress(range.RangeAddress),
        Range = range,
        Shift = shift,
    };

    private static string Kept(XLConditionalFormats formats, string ruleId)
        => formats.TryGetExtensionRuleAreas(ruleId, out var areas) ? string.Join(" ", areas) : "(none)";

    private sealed record Saved(
        string ModelledRange,
        string KeptRange,
        string ModelledFormula,
        string KeptFormula,
        string CellFormula);

    private static Saved EditAndSave(string sqref, string formula, string edit)
    {
        using var built = Build(sqref, formula);
        using var ms = new MemoryStream();
        string cellFormula;
        using (var wb = new XLWorkbook(built))
        {
            Apply(wb, edit);
            cellFormula = wb.Worksheet("Other").Cell("H1").FormulaA1;
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var document = SpreadsheetDocument.Open(ms, false);
        var worksheet = OtherPart(document).Worksheet!;
        var modelled = worksheet.Elements<S.ConditionalFormatting>().SingleOrDefault();
        var kept = worksheet.Descendants<X14.ConditionalFormatting>().SingleOrDefault();
        return new Saved(
            Normalize(modelled?.SequenceOfReferences?.InnerText),
            Normalize(kept?.GetFirstChild<OfficeExcel.ReferenceSequence>()?.Text),
            modelled?.Descendants<S.Formula>().Single().Text ?? string.Empty,
            kept?.Descendants<OfficeExcel.Formula>().Single().Text ?? string.Empty,
            cellFormula);
    }

    private static void Apply(XLWorkbook wb, string edit)
    {
        var other = wb.Worksheet("Other");
        var data = wb.Worksheet("Data");
        switch (edit)
        {
            case "insert rows above": other.Row(1).InsertRowsAbove(2); break;
            case "insert one row above": other.Row(1).InsertRowsAbove(1); break;
            case "insert columns left": other.Column(1).InsertColumnsBefore(2); break;
            case "insert a row inside": other.Row(4).InsertRowsAbove(1); break;
            case "insert a column inside": other.Column(4).InsertColumnsBefore(1); break;
            case "delete a row inside": other.Row(4).Delete(); break;
            case "delete row 2": other.Row(2).Delete(); break;
            case "delete the rows holding the range": other.Rows(3, 5).Delete(); break;
            case "delete the column holding the range": other.Column(3).Delete(); break;
            case "delete rows overlapping the bottom": other.Rows(4, 6).Delete(); break;
            case "delete rows overlapping the top": other.Rows(1, 3).Delete(); break;
            case "delete a cell inside, shifting up": other.Range("C4").Delete(XLShiftDeletedCells.ShiftCellsUp); break;
            case "insert a cell above, shifting down": other.Range("C1").InsertRowsAbove(1); break;
            case "insert a cell above another column": other.Range("A1").InsertRowsAbove(1); break;
            case "insert rows on the referenced sheet": data.Row(1).InsertRowsAbove(1); break;
            case "insert a row above the referenced cell": data.Row(2).InsertRowsAbove(1); break;
            case "insert a row below the referenced cell": data.Row(3).InsertRowsAbove(1); break;
            case "insert a column on the referenced sheet": data.Column(1).InsertColumnsBefore(1); break;
            case "delete a row above the referenced cell": data.Row(1).Delete(); break;
            case "delete the referenced row": data.Row(2).Delete(); break;
            default: throw new ArgumentOutOfRangeException(nameof(edit), edit, null);
        }
    }

    /// <summary>
    /// A workbook whose sheet <c>Other</c> holds the rule twice over <paramref name="sqref"/>: once
    /// modelled, and once only in the <c>x14</c> extension, which is how Excel writes a rule that
    /// refers to another sheet. <c>Other!H1</c> holds the same text as a cell formula, which only
    /// edits on <c>Data</c> leave where it is.
    /// </summary>
    private static MemoryStream Build(string sqref, string formula)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Data").Cell("A2").Value = 1;
            var other = wb.AddWorksheet("Other");
            var areas = sqref.Split(' ');
            other.Range(areas[0]).AddConditionalFormat()
                .SetRanges(areas.Select(a => other.Range(a)))
                .WhenIsTrue("=" + formula).Fill.SetBackgroundColor(XLColor.Red);
            if (formula == Formula)
                other.Cell("H1").FormulaA1 = formula;
            wb.SaveAs(ms);
        }

        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var worksheet = OtherPart(document).Worksheet!;
            worksheet.Append(new S.WorksheetExtensionList(
                "<x:extLst xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<x:ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" +
                "<x14:conditionalFormattings>" +
                "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
                $"<x14:cfRule type=\"expression\" priority=\"2\" id=\"{RuleId}\">" +
                $"<xm:f>{SecurityElement.Escape(formula)}</xm:f><x14:dxf/></x14:cfRule>" +
                $"<xm:sqref>{sqref}</xm:sqref>" +
                "</x14:conditionalFormatting></x14:conditionalFormattings></x:ext></x:extLst>"));
            worksheet.Save();
        }

        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// A range list in one spelling: XLibur writes a one-cell area as <c>C3:C3</c> where Excel writes
    /// <c>C3</c>. Absent is the empty string.
    /// </summary>
    private static string Normalize(string? sqref)
        => sqref is null
            ? string.Empty
            : string.Join(" ", sqref.Split(' ').Select(a => Area.Parse(a).ToString()));

    private static List<string> KeptRanges(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return OtherPart(document).Worksheet!.Descendants<X14.ConditionalFormatting>()
            .Select(c => c.GetFirstChild<OfficeExcel.ReferenceSequence>()!.Text)
            .ToList();
    }

    private static string ExtensionXml(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return OtherPart(document).Worksheet!.Descendants<X14.ConditionalFormattings>().Single().OuterXml;
    }

    private static WorksheetPart OtherPart(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().Single(s => s.Name == "Other");
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));
}
