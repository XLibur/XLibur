using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

        foreach (var definedName in workbook.DefinedNames.OfType<DefinedName>())
        {
            var name = definedName.Name;
            var visible = true;
            if (definedName.Hidden != null) visible = !BooleanValue.ToBoolean(definedName.Hidden);

            var localSheetId = -1;
            if (definedName.LocalSheetId?.HasValue ?? false)
                localSheetId = Convert.ToInt32(definedName.LocalSheetId!.Value);

            if (name == "_xlnm.Print_Area")
            {
                try
                {
                    LoadPrintAreas(definedName, xlWorkbook, localSheetId);
                }
                catch
                {
                    // The print area text is a formula (e.g. OFFSET) that can't be
                    // resolved to simple range references. Store the raw text so it
                    // can be round-tripped on save.
                    var ws = xlWorkbook.WorksheetsInternal.FirstOrDefault<XLWorksheet>(w => w.SheetId == (localSheetId + 1));
                    if (ws != null)
                        ((XLPrintAreas)ws.PageSetup.PrintAreas).FormulaReference = definedName.Text;
                }
            }
            else if (name == "_xlnm.Print_Titles")
            {
                LoadPrintTitles(definedName, xlWorkbook);
            }
            else
            {
                var text = definedName.Text;

                var comment = definedName.Comment;
                if (localSheetId == -1)
                {
                    if (xlWorkbook.DefinedNamesInternal.All<XLDefinedName>(nr => nr.Name != name))
                        xlWorkbook.DefinedNamesInternal.Add(name!, text, comment, validateName: false, validateRangeAddress: false)
                            .Visible = visible;
                }
                else
                {
                    if (xlWorkbook.Worksheet(localSheetId + 1).DefinedNames.All(nr => nr.Name != name))
                        ((XLDefinedNames)xlWorkbook.Worksheet(localSheetId + 1).DefinedNames).Add(name!, text, comment,
                            validateName: false, validateRangeAddress: false).Visible = visible;
                }
            }
        }
    }

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

    private static void LoadPrintAreas(DefinedName definedName, XLWorkbook xlWorkbook, int localSheetId)
    {
        var fixedNames = ValidateDefinedNames(definedName.Text.Split(','));
        foreach (var area in fixedNames)
        {
            if (area.Contains("["))
            {
                var ws = xlWorkbook.WorksheetsInternal.FirstOrDefault<XLWorksheet>(w => w.SheetId == (localSheetId + 1));
                if (ws != null)
                {
                    ws.PageSetup.PrintAreas.Add(area);
                }
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
        if (sections.Count() == 1)
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
