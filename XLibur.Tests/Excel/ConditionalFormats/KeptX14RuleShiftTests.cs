using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.Coordinates;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// A conditional format rule XLibur keeps in the worksheet's <c>x14</c> extension without modelling
/// it (issue #499). Its range, <c>xm:sqref</c>, must move over a row or column insert or delete
/// exactly as a modelled rule's range does. Each differential test builds the same rule both ways on
/// one sheet, makes one edit, saves, and compares the two ranges written.
/// </summary>
public class KeptX14RuleShiftTests
{
    private const string Folder = @"Other\SheetLifecycle\";
    private const string RuleId = "{0A5E0C1D-0000-4000-8000-000000000499}";
    private const string Formula = "Data!$A$2>0";

    /// <summary>
    /// The expected range is what the modelled rule is written with. The kept rule must agree with
    /// it, and <paramref name="expected"/> pins the modelled answer so that the test cannot pass by
    /// both being wrong the same way. An empty string means the edit removed the rule.
    /// </summary>
    [Test]
    [Arguments("C3:C5", "insert rows above", "C5:C7")]
    [Arguments("C3:C5", "insert columns left", "E3:E5")]
    [Arguments("C3:C5", "insert a row inside", "C3:C6")]
    [Arguments("C3:E5", "insert a column inside", "C3:F5")]
    [Arguments("C3:C5", "delete a row inside", "C3:C4")]
    [Arguments("C3:C5", "delete the rows holding the range", "")]
    [Arguments("C3:C5", "delete the column holding the range", "")]
    [Arguments("C3:C5", "delete rows overlapping the bottom", "C3")]
    [Arguments("C3:C5", "delete rows overlapping the top", "C1:C2")]
    [Arguments("C3:C5", "delete a cell inside, shifting up", "C3:C4")]
    [Arguments("C3:C5", "insert a cell above, shifting down", "C4:C6")]
    [Arguments("C3:C5", "insert a cell above another column", "C3:C5")]
    [Arguments("C3:C5", "insert rows on the referenced sheet", "C3:C5")]
    [Arguments("C3:C5 E3:E5", "insert rows above", "C5:C7 E5:E7")]
    [Arguments("C1", "insert rows above", "C3")]
    public async Task A_kept_rules_range_shifts_as_a_modelled_rules_does(string sqref, string edit, string expected)
    {
        var saved = EditAndSave(sqref, edit);

        await Assert.That(saved.Modelled).IsEqualTo(expected);
        await Assert.That(saved.Kept).IsEqualTo(saved.Modelled);
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
    /// A row inserted on <c>Data</c>, the sheet the rules' formulas refer to, moves no range on
    /// <c>Other</c>.
    /// </summary>
    [Test]
    public async Task Excels_kept_rules_ranges_stay_when_the_referenced_sheet_gets_a_row()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("rename-before.xlsx")))
        {
            wb.Worksheet("Data").Row(2).InsertRowsAbove(1);
            wb.SaveAs(ms);
        }

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
        {
            ISheetListener listener = sheet.ConditionalFormats;
            var row = sheet.Range(1, 1, 1, XLHelper.MaxColumnNumber);
            var column = sheet.Range(1, 1, XLHelper.MaxRowNumber, 1);

            await Assert.That(() => listener.OnInsertAreaAndShiftDown(Edit(sheet, row, 1))).ThrowsNothing();
            await Assert.That(() => listener.OnInsertAreaAndShiftRight(Edit(sheet, column, 1))).ThrowsNothing();
            await Assert.That(() => listener.OnDeleteAreaAndShiftUp(Edit(sheet, row, -1))).ThrowsNothing();
            await Assert.That(() => listener.OnDeleteAreaAndShiftLeft(Edit(sheet, column, -1))).ThrowsNothing();
        }

        await Assert.That(Kept(formats, "emptied")).IsEqualTo("");
        await Assert.That(Kept(formats, "sheet")).IsEqualTo("A1:XFD1048576");
        await Assert.That(Kept(formats, "last row")).IsEqualTo("");
        await Assert.That(Kept(formats, "edges")).IsEqualTo("A1");
    }

    private static SheetEdit Edit(XLWorksheet sheet, XLRange range, int shift) => new()
    {
        Sheet = sheet,
        Area = Area.FromRangeAddress(range.RangeAddress),
        Range = range,
        Shift = shift,
    };

    private static string Kept(XLibur.Excel.ConditionalFormats.XLConditionalFormats formats, string ruleId)
        => formats.TryGetExtensionRuleAreas(ruleId, out var areas) ? string.Join(" ", areas) : "(none)";

    private static (string Modelled, string Kept) EditAndSave(string sqref, string edit)
    {
        using var built = Build(sqref);
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(built))
        {
            Apply(wb, edit);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var document = SpreadsheetDocument.Open(ms, false);
        var worksheet = OtherPart(document).Worksheet!;
        var modelled = worksheet.Elements<S.ConditionalFormatting>().SingleOrDefault()?.SequenceOfReferences?.InnerText;
        var kept = worksheet.Descendants<X14.ConditionalFormatting>().SingleOrDefault()
            ?.GetFirstChild<OfficeExcel.ReferenceSequence>()?.Text;
        return (Normalize(modelled), Normalize(kept));
    }

    private static void Apply(XLWorkbook wb, string edit)
    {
        var other = wb.Worksheet("Other");
        switch (edit)
        {
            case "insert rows above": other.Row(1).InsertRowsAbove(2); break;
            case "insert columns left": other.Column(1).InsertColumnsBefore(2); break;
            case "insert a row inside": other.Row(4).InsertRowsAbove(1); break;
            case "insert a column inside": other.Column(4).InsertColumnsBefore(1); break;
            case "delete a row inside": other.Row(4).Delete(); break;
            case "delete the rows holding the range": other.Rows(3, 5).Delete(); break;
            case "delete the column holding the range": other.Column(3).Delete(); break;
            case "delete rows overlapping the bottom": other.Rows(4, 6).Delete(); break;
            case "delete rows overlapping the top": other.Rows(1, 3).Delete(); break;
            case "delete a cell inside, shifting up": other.Range("C4").Delete(XLShiftDeletedCells.ShiftCellsUp); break;
            case "insert a cell above, shifting down": other.Range("C1").InsertRowsAbove(1); break;
            case "insert a cell above another column": other.Range("A1").InsertRowsAbove(1); break;
            case "insert rows on the referenced sheet": wb.Worksheet("Data").Row(1).InsertRowsAbove(1); break;
            default: throw new ArgumentOutOfRangeException(nameof(edit), edit, null);
        }
    }

    /// <summary>
    /// A workbook whose sheet <c>Other</c> holds the rule twice over <paramref name="sqref"/>: once
    /// modelled, and once only in the <c>x14</c> extension, which is how Excel writes a rule that
    /// refers to another sheet.
    /// </summary>
    private static MemoryStream Build(string sqref)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Data").Cell("A2").Value = 1;
            var other = wb.AddWorksheet("Other");
            var areas = sqref.Split(' ');
            other.Range(areas[0]).AddConditionalFormat()
                .SetRanges(areas.Select(a => other.Range(a)))
                .WhenIsTrue("=" + Formula).Fill.SetBackgroundColor(XLColor.Red);
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
                $"<xm:f>{SecurityElement.Escape(Formula)}</xm:f><x14:dxf/></x14:cfRule>" +
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
