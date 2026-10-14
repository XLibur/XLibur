using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Utils;

namespace XLibur.Excel.IO;

internal static class SheetProtectionWriter
{
    internal static void WriteSheetProtection(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        // A <sheetProtection> that does not protect the sheet, as LibreOffice writes one, is written
        // again on an unprotected sheet: the load reads its settings, so they are written from the
        // model and keep the file's spelling while they hold (#722).
        var loaded = worksheet.Elements<SheetProtection>().FirstOrDefault();
        var protectsNothing = loaded is not null && !OpenXmlHelper.GetBooleanValueAsBool(loaded.Sheet, false);

        if (xlWorksheet.Protection.IsProtected || protectsNothing)
        {
            if (!worksheet.Elements<SheetProtection>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.SheetProtection);
                worksheet.InsertAfter(new SheetProtection(), previousElement);
            }

            var sheetProtection = worksheet.Elements<SheetProtection>().First();
            cm.SetElement(XLWorksheetContents.SheetProtection, sheetProtection);

            var protection = xlWorksheet.Protection;
            sheetProtection.Sheet = SchemaDefault.Bool(sheetProtection.Sheet, protection.IsProtected, false);

            sheetProtection.Password = null;
            sheetProtection.AlgorithmName = null;
            sheetProtection.HashValue = null;
            sheetProtection.SpinCount = null;
            sheetProtection.SaltValue = null;

            if (protection.Algorithm == XLProtectionAlgorithm.Algorithm.SimpleHash)
            {
                if (!string.IsNullOrWhiteSpace(protection.PasswordHash))
                    sheetProtection.Password = protection.PasswordHash;
            }
            else
            {
                sheetProtection.AlgorithmName =
                    DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.ToDescription(protection.Algorithm);
                sheetProtection.HashValue = protection.PasswordHash;
                sheetProtection.SpinCount = protection.SpinCount;
                sheetProtection.SaltValue = protection.Base64EncodedSalt;
            }

            // Each attribute is set when the element is protected, that is, not allowed. A value the
            // file spelled out stays, default or not (#722).
            var allowed = protection.AllowedElements;

            // default value of "1"
            sheetProtection.FormatCells = Protects(sheetProtection.FormatCells, XLSheetProtectionElements.FormatCells, true);
            sheetProtection.FormatColumns = Protects(sheetProtection.FormatColumns, XLSheetProtectionElements.FormatColumns, true);
            sheetProtection.FormatRows = Protects(sheetProtection.FormatRows, XLSheetProtectionElements.FormatRows, true);
            sheetProtection.InsertColumns = Protects(sheetProtection.InsertColumns, XLSheetProtectionElements.InsertColumns, true);
            sheetProtection.InsertRows = Protects(sheetProtection.InsertRows, XLSheetProtectionElements.InsertRows, true);
            sheetProtection.InsertHyperlinks = Protects(sheetProtection.InsertHyperlinks, XLSheetProtectionElements.InsertHyperlinks, true);
            sheetProtection.DeleteColumns = Protects(sheetProtection.DeleteColumns, XLSheetProtectionElements.DeleteColumns, true);
            sheetProtection.DeleteRows = Protects(sheetProtection.DeleteRows, XLSheetProtectionElements.DeleteRows, true);
            sheetProtection.Sort = Protects(sheetProtection.Sort, XLSheetProtectionElements.Sort, true);
            sheetProtection.AutoFilter = Protects(sheetProtection.AutoFilter, XLSheetProtectionElements.AutoFilter, true);
            sheetProtection.PivotTables = Protects(sheetProtection.PivotTables, XLSheetProtectionElements.PivotTables, true);
            sheetProtection.Scenarios = Protects(sheetProtection.Scenarios, XLSheetProtectionElements.EditScenarios, true);

            // default value of "0"
            sheetProtection.Objects = Protects(sheetProtection.Objects, XLSheetProtectionElements.EditObjects, false);
            sheetProtection.SelectLockedCells = Protects(sheetProtection.SelectLockedCells, XLSheetProtectionElements.SelectLockedCells, false);
            sheetProtection.SelectUnlockedCells = Protects(sheetProtection.SelectUnlockedCells, XLSheetProtectionElements.SelectUnlockedCells, false);

            BooleanValue? Protects(BooleanValue? loaded, XLSheetProtectionElements element, bool schemaDefault) =>
                SchemaDefault.Bool(loaded, !allowed.HasFlag(element), schemaDefault);
        }
        else
        {
            worksheet.RemoveAllChildren<SheetProtection>();
            cm.SetElement(XLWorksheetContents.SheetProtection, null);
        }
    }
}
