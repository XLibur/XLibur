using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.Coordinates;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// What Excel writes for a conditional format's range and formula after a row or column edit, read
/// from <c>cf-anchor-*.xlsx</c>. <c>Sheet1</c> holds three rules: <c>A2:C10</c> with <c>$A2&gt;5</c>,
/// <c>G2:H10</c> with <c>G2&gt;5</c>, and the whole column <c>E1:E1048576</c> with <c>$A1&gt;5</c>.
/// Each test makes Excel's edit through XLibur and compares every rule with the file Excel saved.
/// </summary>
/// <remarks>
/// <para>
/// A formula is relative to the range's first cell, its anchor. Deleting the anchor's row or column
/// while the rule survives does not give <c>#REF!</c>: Excel rebases the formula onto the first cell
/// that survives, then shifts it. So <c>$A2&gt;5</c> on <c>A2:C10</c> is still <c>$A2&gt;5</c> after row
/// 2 is deleted. A whole-column range cannot move, but its formula still shifts: a row inserted at 1
/// turns <c>$A1&gt;5</c> into <c>$A2&gt;5</c>.
/// </para>
/// <para>
/// On the insert Excel also renumbered the rules' priorities and <c>dxfId</c>s. Those are not
/// compared; each rule is compared by its range and its formula.
/// </para>
/// </remarks>
public class ConditionalFormatAnchorFixtureTests
{
    private const string Folder = @"Other\ConditionalFormatShift\";

    /// <summary>
    /// Saved with the default <see cref="SaveOptions"/>, which consolidate conditional formats, so that
    /// a save cannot merge back the pieces an edit cut a rule into.
    /// </summary>
    [Test]
    [Arguments("cf-anchor-before.xlsx", "none", "cf-anchor-before.xlsx")]
    [Arguments("cf-anchor-before.xlsx", "delete row 2", "cf-anchor-deleterow-after.xlsx")]
    [Arguments("cf-anchor-before.xlsx", "delete column G", "cf-anchor-deletecol-after.xlsx")]
    [Arguments("cf-anchor-before.xlsx", "insert a row at 1", "cf-anchor-insertrow-after.xlsx")]
    [Arguments("cf-partial-before.xlsx", "none", "cf-partial-before.xlsx")]
    [Arguments("cf-partial-before.xlsx", "delete A2 shifting left", "cf-partial-deleteleft-after.xlsx")]
    [Arguments("cf-partial-before.xlsx", "insert a cell at A2 shifting down", "cf-partial-insertdown-after.xlsx")]
    public async Task Each_block_matches_what_Excel_wrote(string before, string edit, string after)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(before)))
        {
            var ws = wb.Worksheet("Sheet1");
            switch (edit)
            {
                case "none": break;
                case "delete row 2": ws.Row(2).Delete(); break;
                case "delete column G": ws.Column(7).Delete(); break;
                case "insert a row at 1": ws.Row(1).InsertRowsAbove(1); break;
                case "delete A2 shifting left": ws.Range("A2").Delete(XLShiftDeletedCells.ShiftCellsLeft); break;
                case "insert a cell at A2 shifting down": ws.Cell("A2").InsertCellsAbove(1); break;
                default: throw new ArgumentOutOfRangeException(nameof(edit), edit, null);
            }

            wb.SaveAs(ms);
        }

        await Assert.That(Lines(Blocks(ms))).IsEqualTo(Lines(Blocks(Resource(after))));
    }

    private static string Lines(IEnumerable<string> items) => string.Join(Environment.NewLine, items);

    /// <summary>
    /// Every conditional-formatting block on <c>Sheet1</c> as <c>range = formulas</c>, the formulas of
    /// its rules sorted, and the blocks in a fixed order. Priorities, <c>dxfId</c>s and the order of
    /// blocks and rules are Excel's bookkeeping, and Excel renumbers them on an edit.
    /// </summary>
    private static List<string> Blocks(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().Single(s => s.Name == "Sheet1");
        var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;
        return worksheet.Elements<S.ConditionalFormatting>()
            .Select(c => $"{Normalize(c.SequenceOfReferences?.InnerText ?? string.Empty)} = " +
                         string.Join(" | ", c.Elements<S.ConditionalFormattingRule>()
                             .Select(r => string.Join(" & ", r.Elements<S.Formula>().Select(f => f.Text)))
                             .Order(StringComparer.Ordinal)))
            .Order(StringComparer.Ordinal)
            .ToList();
    }

    private static string Normalize(string sqref) => string.Join(" ", sqref.Split(' ').Select(NormalizeArea));

    /// <summary>
    /// One spelling per area: <c>C3</c> for <c>C3:C3</c>, and a whole column or row written out with
    /// both corners, as Excel writes <c>E1:E1048576</c> where a writer may write <c>E:E</c>.
    /// </summary>
    private static string NormalizeArea(string area)
    {
        if (Area.TryParse(area, out var parsed))
            return parsed.ToString();

        var ends = area.Split(':');
        if (ends.Length == 2 && ends.All(e => e.Length > 0 && e.All(char.IsLetter)))
            return $"{ends[0]}1:{ends[1]}{XLHelper.MaxRowNumber}";
        if (ends.Length == 2 && ends.All(e => e.Length > 0 && e.All(char.IsDigit)))
            return $"A{ends[0]}:XFD{ends[1]}";

        return area;
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));
}
