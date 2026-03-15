using XLibur.Excel.ContentManagers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

internal sealed class DataValidationWriter
{
    internal static void WriteDataValidations(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        SaveOptions options)
    {
        // Saving of data validations happens in 2 phases because depending on the data validation
        // content, it gets saved into 1 of 2 possible locations in the XML structure.
        // First phase, save all the data validations that aren't references to other sheets into
        // the standard data validations section.
        var dataValidationsStandard = new List<(IXLDataValidation DataValidation, string MinValue, string MaxValue)>();
        var dataValidationsExtension = new List<(IXLDataValidation DataValidation, string MinValue, string MaxValue)>();
        if (options.ConsolidateDataValidationRanges)
        {
            xlWorksheet.DataValidations.Consolidate();
        }

        foreach (var dv in xlWorksheet.DataValidations)
        {
            var (minReferencesAnotherSheet, minValue) = UsesExternalSheet(xlWorksheet, dv.MinValue);
            var (maxReferencesAnotherSheet, maxValue) = UsesExternalSheet(xlWorksheet, dv.MaxValue);

            // Standard <dataValidation> element limits formula1/formula2 to 255 chars.
            // Longer formulas or formulas referencing another sheet must use X14 extension.
            var formulaTooLong = minValue.Length > 255 || maxValue.Length > 255;
            if (minReferencesAnotherSheet || maxReferencesAnotherSheet || formulaTooLong)
            {
                // We're dealing with a data validation that references another sheet or has long formulas, so has to be saved to extensions
                dataValidationsExtension.Add((dv, minValue, maxValue));
            }
            else
            {
                // We're dealing with a standard data validation
                dataValidationsStandard.Add((dv, minValue, maxValue));
            }
        }

        WriteStandardDataValidations(worksheet, cm, dataValidationsStandard);
        WriteExtensionDataValidations(worksheet, cm, dataValidationsExtension);
    }

    private static (bool, string) UsesExternalSheet(XLWorksheet sheet, string value)
    {
        if (!XLHelper.IsValidRangeAddress(value))
            return (false, value);

        var separatorIndex = value.LastIndexOf('!');
        var hasSheet = separatorIndex >= 0;
        if (!hasSheet)
            return (false, value);

        var sheetName = value[..separatorIndex].UnescapeSheetName();
        if (XLHelper.SheetComparer.Equals(sheet.Name, sheetName))
        {
            // The spec wants us to include references to ranges on the same worksheet without the sheet name
            return (false, value[(separatorIndex + 1)..]);
        }

        return (true, value);
    }

    private static void WriteStandardDataValidations(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        List<(IXLDataValidation DataValidation, string MinValue, string MaxValue)> dataValidationsStandard)
    {
        // Save validations that don't use another sheet. It must have at least 1 child, XML doesn't allow 0.
        if (!dataValidationsStandard.Any(d => d.DataValidation.IsDirty()))
        {
            worksheet.RemoveAllChildren<DataValidations>();
            cm.SetElement(XLWorksheetContents.DataValidations, null);
        }
        else
        {
            if (!worksheet.Elements<DataValidations>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.DataValidations);
                worksheet.InsertAfter(new DataValidations(), previousElement);
            }

            var dataValidations = worksheet.Elements<DataValidations>().First();
            cm.SetElement(XLWorksheetContents.DataValidations, dataValidations);
            dataValidations.RemoveAllChildren<DataValidation>();

            foreach (var (dv, minValue, maxValue) in dataValidationsStandard)
            {
                var sequence = string.Join(" ", dv.Ranges.Select(x => x.RangeAddress));
                var dataValidation = new DataValidation
                {
                    AllowBlank = dv.IgnoreBlanks,
                    Formula1 = new Formula1(minValue),
                    Formula2 = new Formula2(maxValue),
                    Type = dv.AllowedValues.ToOpenXml(),
                    ShowErrorMessage = dv.ShowErrorMessage,
                    Prompt = dv.InputMessage,
                    PromptTitle = dv.InputTitle,
                    ErrorTitle = dv.ErrorTitle,
                    Error = dv.ErrorMessage,
                    ShowDropDown = !dv.InCellDropdown,
                    ShowInputMessage = dv.ShowInputMessage,
                    ErrorStyle = dv.ErrorStyle.ToOpenXml(),
                    Operator = HasOperator(dv.AllowedValues) ? dv.Operator.ToOpenXml() : null,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = sequence }
                };

                dataValidations.AppendChild(dataValidation);
            }

            dataValidations.Count = (uint)dataValidationsStandard.Count;
        }
    }

    private static void WriteExtensionDataValidations(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        List<(IXLDataValidation DataValidation, string MinValue, string MaxValue)> dataValidationsExtension)
    {
        const string dataValidationsExtensionUri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";
        if (dataValidationsExtension.Count == 0)
        {
            RemoveExtensionDataValidations(worksheet, cm, dataValidationsExtensionUri);
        }
        else
        {
            WriteExtensionDataValidationElements(worksheet, cm, dataValidationsExtension, dataValidationsExtensionUri);
        }
    }

    private static void RemoveExtensionDataValidations(Worksheet worksheet, XLWorksheetContentManager cm,
        string dataValidationsExtensionUri)
    {
        var worksheetExtensionList = worksheet.Elements<WorksheetExtensionList>().FirstOrDefault();
        var worksheetExtension = worksheetExtensionList?.Elements<WorksheetExtension>()
            .FirstOrDefault(ext =>
                string.Equals(ext.Uri, dataValidationsExtensionUri, StringComparison.OrdinalIgnoreCase));

        worksheetExtension?.RemoveAllChildren<X14.DataValidations>();

        if (worksheetExtensionList == null)
            return;

        if (worksheetExtension is { HasChildren: false })
            worksheetExtensionList.RemoveChild(worksheetExtension);

        if (!worksheetExtensionList.HasChildren)
        {
            worksheet.RemoveChild(worksheetExtensionList);
            cm.SetElement(XLWorksheetContents.WorksheetExtensionList, null);
        }
    }

    private static void WriteExtensionDataValidationElements(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        List<(IXLDataValidation DataValidation, string MinValue, string MaxValue)> dataValidationsExtension,
        string dataValidationsExtensionUri)
    {
        if (!worksheet.Elements<WorksheetExtensionList>().Any())
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.WorksheetExtensionList);
            worksheet.InsertAfter(new WorksheetExtensionList(), previousElement);
        }

        var worksheetExtensionList = worksheet.Elements<WorksheetExtensionList>().First();
        cm.SetElement(XLWorksheetContents.WorksheetExtensionList, worksheetExtensionList);

        var extensionDataValidations = worksheetExtensionList.Descendants<X14.DataValidations>().SingleOrDefault();

        if (extensionDataValidations == null || !extensionDataValidations.Any())
        {
            var worksheetExtension = new WorksheetExtension() { Uri = dataValidationsExtensionUri };
            worksheetExtension.AddNamespaceDeclaration("x14", X14Main2009SsNs);
            worksheetExtensionList.Append(worksheetExtension);

            extensionDataValidations = new X14.DataValidations();
            extensionDataValidations.AddNamespaceDeclaration("xm", XmMain2006);
            worksheetExtension.Append(extensionDataValidations);
        }
        else
        {
            extensionDataValidations.RemoveAllChildren();
        }

        foreach (var (dv, minValue, maxValue) in dataValidationsExtension)
        {
            var sequence = string.Join(" ", dv.Ranges.Select(x => x.RangeAddress));
            var dataValidation = new X14.DataValidation
            {
                AllowBlank = dv.IgnoreBlanks,
                DataValidationForumla1 = !string.IsNullOrWhiteSpace(minValue)
                    ? new X14.DataValidationForumla1(new OfficeExcel.Formula(minValue))
                    : null,
                DataValidationForumla2 = !string.IsNullOrWhiteSpace(maxValue)
                    ? new X14.DataValidationForumla2(new OfficeExcel.Formula(maxValue))
                    : null,
                Type = dv.AllowedValues.ToOpenXml(),
                ShowErrorMessage = dv.ShowErrorMessage,
                Prompt = dv.InputMessage,
                PromptTitle = dv.InputTitle,
                ErrorTitle = dv.ErrorTitle,
                Error = dv.ErrorMessage,
                ShowDropDown = !dv.InCellDropdown,
                ShowInputMessage = dv.ShowInputMessage,
                ErrorStyle = dv.ErrorStyle.ToOpenXml(),
                Operator = HasOperator(dv.AllowedValues) ? dv.Operator.ToOpenXml() : null,
                ReferenceSequence = new OfficeExcel.ReferenceSequence() { Text = sequence }
            };
            extensionDataValidations.AppendChild(dataValidation);
        }

        extensionDataValidations.Count = (uint)dataValidationsExtension.Count;
    }

    /// <summary>
    /// Only validation types that compare values use the operator attribute.
    /// List, Custom, and AnyValue do not.
    /// </summary>
    private static bool HasOperator(XLAllowedValues allowedValues) => allowedValues switch
    {
        XLAllowedValues.WholeNumber or
        XLAllowedValues.Decimal or
        XLAllowedValues.Date or
        XLAllowedValues.Time or
        XLAllowedValues.TextLength => true,
        _ => false,
    };
}
