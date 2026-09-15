using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.DataValidations;

/// <summary>
/// Reads back the data validations a save wrote to one worksheet's part, in either of the two forms
/// Excel has for a rule.
/// </summary>
internal static class SavedDataValidations
{
    /// <summary>
    /// Each rule in <paramref name="sheetName"/>'s part: <c>standard</c> for one in
    /// <c>&lt;dataValidations&gt;</c>, <c>x14</c> for one in the extension, with its two formulas, an
    /// absent one read as empty.
    /// </summary>
    internal static List<(string Form, string Formula1, string Formula2)> Criteria(Stream package,
        string sheetName)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var worksheet = SheetPart(document, sheetName).Worksheet!;
        var standard = worksheet.Elements<S.DataValidations>()
            .SelectMany(d => d.Elements<S.DataValidation>())
            .Select(dv => ("standard", dv.Formula1?.Text ?? "", dv.Formula2?.Text ?? ""));
        var extension = worksheet.Descendants<X14.DataValidation>()
            .Select(dv => ("x14", dv.DataValidationForumla1?.InnerText ?? "", dv.DataValidationForumla2?.InnerText ?? ""));
        return standard.Concat(extension).ToList();
    }

    internal static WorksheetPart SheetPart(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().Single(s => s.Name?.Value == sheetName);
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }
}
