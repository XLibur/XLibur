using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Extensions;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads defined names (named ranges, print areas, print titles) from the workbook.
/// </summary>
internal static class DefinedNameReader
{
    private static readonly Regex DefinedNameRegex = new(@"\A('?).*\1!.*\z", RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    internal static void LoadDefinedNames(Workbook workbook, XLWorkbook xlWorkbook)
    {
        if (workbook.DefinedNames == null) return;

        var sheetsByPosition = SheetsByPosition(workbook, xlWorkbook);
        foreach (var definedName in workbook.DefinedNames.OfType<DefinedName>())
        {
            var name = definedName.Name;
            var visible = true;
            if (definedName.Hidden != null) visible = !BooleanValue.ToBoolean(definedName.Hidden);

            var localSheetId = -1;
            if (definedName.LocalSheetId?.HasValue ?? false)
                localSheetId = Convert.ToInt32(definedName.LocalSheetId.Value);

            if (name == "_xlnm.Print_Area")
            {
                LoadPrintAreaSafe(definedName, xlWorkbook, SheetAt(sheetsByPosition, localSheetId));
            }
            else if (name == "_xlnm.Print_Titles")
            {
                LoadPrintTitles(definedName, xlWorkbook);
            }
            else
            {
                LoadNamedRange(definedName, xlWorkbook, name!, visible, localSheetId, sheetsByPosition);
            }
        }
    }

    /// <summary>
    /// The worksheet at each position of the file's <c>&lt;sheets&gt;</c> list, or null where the
    /// position holds a sheet XLibur does not model, such as a chartsheet.
    /// </summary>
    /// <remarks>
    /// A <c>localSheetId</c> is a position in that list, which counts every sheet (ECMA-376), and not
    /// a <c>sheetId</c>. The two differ once a sheet has been inserted before others, since it keeps a
    /// higher sheetId than theirs. Reading one as the other put a print area on the wrong sheet (D75).
    /// </remarks>
    private static List<XLWorksheet?> SheetsByPosition(Workbook workbook, XLWorkbook xlWorkbook)
        => (workbook.Sheets?.Elements<Sheet>() ?? [])
            .Select(s => xlWorkbook.WorksheetsInternal.FirstOrDefault<XLWorksheet>(w => w.SheetId == s.SheetId?.Value))
            .ToList();

    private static XLWorksheet? SheetAt(List<XLWorksheet?> sheetsByPosition, int localSheetId)
        => localSheetId >= 0 && localSheetId < sheetsByPosition.Count ? sheetsByPosition[localSheetId] : null;

    internal static IEnumerable<string> ValidateDefinedNames(IEnumerable<string> definedNames)
    {
        var sb = new StringBuilder();
        foreach (var testName in definedNames)
        {
            if (sb.Length > 0)
                sb.Append(',');

            sb.Append(testName);

            var matchedValidPattern = DefinedNameRegex.Match(sb.ToString());
            if (matchedValidPattern.Success)
            {
                yield return sb.ToString();
                sb = new StringBuilder();
            }
        }

        if (sb.Length > 0)
            yield return sb.ToString();
    }

    private static void LoadPrintAreaSafe(DefinedName definedName, XLWorkbook xlWorkbook, XLWorksheet? sheet)
    {
        try
        {
            LoadPrintAreas(definedName, xlWorkbook, sheet);
        }
        catch
        {
            // The print area text is a formula (e.g. OFFSET) that can't be
            // resolved to simple range references. Store the raw text so it
            // can be round-tripped on save.
            if (sheet != null)
                ((XLPrintAreas)sheet.PageSetup.PrintAreas).FormulaReference = definedName.Text;
        }
    }

    private static void LoadNamedRange(DefinedName definedName, XLWorkbook xlWorkbook, string name, bool visible,
        int localSheetId, List<XLWorksheet?> sheetsByPosition)
    {
        var text = definedName.Text;
        var comment = definedName.Comment;
        if (localSheetId == -1)
        {
            if (xlWorkbook.DefinedNamesInternal.All<XLDefinedName>(nr => nr.Name != name))
                xlWorkbook.DefinedNamesInternal.Add(name, text, comment, validateName: false, validateRangeAddress: false)
                    .Visible = visible;
        }
        else
        {
            // A name scoped to a sheet XLibur does not model, such as a chartsheet, has nowhere to go,
            // and fails the load as it did before the scope was read from the file's list. Whether to
            // keep such a name instead is a decision of its own.
            var sheet = SheetAt(sheetsByPosition, localSheetId)
                        ?? throw new ArgumentException("There isn't a worksheet associated with that position.");
            if (sheet.DefinedNames.All<XLDefinedName>(nr => nr.Name != name))
                sheet.DefinedNames.Add(name, text, comment, validateName: false, validateRangeAddress: false)
                    .Visible = visible;
        }
    }

    private static void LoadPrintAreas(DefinedName definedName, XLWorkbook xlWorkbook, XLWorksheet? sheet)
    {
        var fixedNames = ValidateDefinedNames(definedName.Text.Split(','));
        foreach (var area in fixedNames)
        {
            if (area.Contains('['))
            {
                sheet?.PageSetup.PrintAreas.Add(area);
            }
            else
            {
                ParseReference(area, out var sheetName, out var sheetArea);
                if (!(sheetArea.Equals("#REF") || sheetArea.EndsWith("#REF!") || sheetArea.Length == 0 ||
                      sheetName.Length == 0))
                    xlWorkbook.WorksheetsInternal.Worksheet(sheetName).PageSetup.PrintAreas.Add(sheetArea);
            }
        }
    }

    private static void LoadPrintTitles(DefinedName definedName, XLWorkbook xlWorkbook)
    {
        var areas = ValidateDefinedNames(definedName.Text.Split(','));
        foreach (var item in areas)
        {
            if (xlWorkbook.Range(item) != null)
                SetColumnsOrRowsToRepeat(item, xlWorkbook);
        }
    }

    private static void SetColumnsOrRowsToRepeat(string area, XLWorkbook xlWorkbook)
    {
        ParseReference(area, out var sheetName, out var sheetArea);
        sheetArea = sheetArea.Replace("$", "");

        if (sheetArea.Equals("#REF")) return;
        if (IsColReference(sheetArea))
            xlWorkbook.WorksheetsInternal.Worksheet(sheetName).PageSetup.SetColumnsToRepeatAtLeft(sheetArea);
        if (IsRowReference(sheetArea))
            xlWorkbook.WorksheetsInternal.Worksheet(sheetName).PageSetup.SetRowsToRepeatAtTop(sheetArea);
    }

    // either $A:$X => true or $1:$99 => false
    private static bool IsColReference(string sheetArea)
    {
        return sheetArea.All(c => c == ':' || char.IsLetter(c));
    }

    private static bool IsRowReference(string sheetArea)
    {
        return sheetArea.All(c => c == ':' || char.IsNumber(c));
    }

    internal static void ParseReference(string item, out string sheetName, out string sheetArea)
    {
        var sections = item.Trim().Split('!');
        if (sections.Length == 1)
        {
            sheetName = string.Empty;
            sheetArea = item;
        }
        else
        {
            sheetName = string.Join("!", sections.Take(sections.Length - 1)).UnescapeSheetName();
            sheetArea = sections[^1];
        }
    }
}
