using XLibur.Excel.ContentManagers;
using XLibur.Utils;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace XLibur.Excel.IO;

internal sealed class SheetProtectionWriter
{
    internal static void WriteSheetProtection(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        if (xlWorksheet.Protection.IsProtected)
        {
            if (!worksheet.Elements<SheetProtection>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.SheetProtection);
                worksheet.InsertAfter(new SheetProtection(), previousElement);
            }

            var sheetProtection = worksheet.Elements<SheetProtection>().First();
            cm.SetElement(XLWorksheetContents.SheetProtection, sheetProtection);

            var protection = xlWorksheet.Protection;
            sheetProtection.Sheet = OpenXmlHelper.GetBooleanValue(protection.IsProtected, false);

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

            // default value of "1"
            sheetProtection.FormatCells =
                OpenXmlHelper.GetBooleanValue(
                    !protection.AllowedElements.HasFlag(XLSheetProtectionElements.FormatCells), true);
            sheetProtection.FormatColumns =
                OpenXmlHelper.GetBooleanValue(
                    !protection.AllowedElements.HasFlag(XLSheetProtectionElements.FormatColumns), true);
            sheetProtection.FormatRows =
                OpenXmlHelper.GetBooleanValue(!protection.AllowedElements.HasFlag(XLSheetProtectionElements.FormatRows),
                    true);
            sheetProtection.InsertColumns =
                OpenXmlHelper.GetBooleanValue(
                    !protection.AllowedElements.HasFlag(XLSheetProtectionElements.InsertColumns), true);
            sheetProtection.InsertRows =
                OpenXmlHelper.GetBooleanValue(!protection.AllowedElements.HasFlag(XLSheetProtectionElements.InsertRows),
                    true);
            sheetProtection.InsertHyperlinks =
                OpenXmlHelper.GetBooleanValue(
                    !protection.AllowedElements.HasFlag(XLSheetProtectionElements.InsertHyperlinks), true);
            sheetProtection.DeleteColumns =
                OpenXmlHelper.GetBooleanValue(
                    !protection.AllowedElements.HasFlag(XLSheetProtectionElements.DeleteColumns), true);
            sheetProtection.DeleteRows =
                OpenXmlHelper.GetBooleanValue(!protection.AllowedElements.HasFlag(XLSheetProtectionElements.DeleteRows),
                    true);
            sheetProtection.Sort =
                OpenXmlHelper.GetBooleanValue(!protection.AllowedElements.HasFlag(XLSheetProtectionElements.Sort),
                    true);
            sheetProtection.AutoFilter =
                OpenXmlHelper.GetBooleanValue(!protection.AllowedElements.HasFlag(XLSheetProtectionElements.AutoFilter),
                    true);
            sheetProtection.PivotTables =
                OpenXmlHelper.GetBooleanValue(
                    !protection.AllowedElements.HasFlag(XLSheetProtectionElements.PivotTables), true);
            sheetProtection.Scenarios =
                OpenXmlHelper.GetBooleanValue(
                    !protection.AllowedElements.HasFlag(XLSheetProtectionElements.EditScenarios), true);

            // default value of "0"
            sheetProtection.Objects =
                OpenXmlHelper.GetBooleanValue(
                    !protection.AllowedElements.HasFlag(XLSheetProtectionElements.EditObjects), false);
            sheetProtection.SelectLockedCells =
                OpenXmlHelper.GetBooleanValue(
                    !protection.AllowedElements.HasFlag(XLSheetProtectionElements.SelectLockedCells), false);
            sheetProtection.SelectUnlockedCells = OpenXmlHelper.GetBooleanValue(
                !protection.AllowedElements.HasFlag(XLSheetProtectionElements.SelectUnlockedCells), false);
        }
        else
        {
            worksheet.RemoveAllChildren<SheetProtection>();
            cm.SetElement(XLWorksheetContents.SheetProtection, null);
        }
    }
}
