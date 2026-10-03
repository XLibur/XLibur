using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.ContentManagers;
using XLibur.Extensions;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

internal static class DataValidationWriter
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
            // A rule can be left covering nothing — ClearRanges and RemoveRange are public and
            // neither drops the rule. sqref is built by joining its ranges, so writing one
            // would emit sqref="", which the schema forbids and Excel repairs the file over.
            // A rule that applies to no cell has nothing to say, so skipping it loses nothing.
            if (((XLDataValidation)dv).Areas.Count == 0)
                continue;

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

    /// <summary>
    /// The text a criterion is written with, and whether it refers to another sheet, which Excel writes
    /// only in the <c>x14</c> extension.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A file stores a criterion as formula text, which has no leading <c>=</c>. A rule built in code can
    /// have one, as in <c>List("=Data!$A$1:$A$3")</c>, so it comes off first. Written as it was, the
    /// <c>=</c> went into <c>&lt;formula1&gt;</c>, and the address check below, which allows no <c>=</c>,
    /// sent a range on another sheet to the standard form (#523). Excel writes such a rule in the
    /// extension, without the <c>=</c>. The rule itself keeps its text, so
    /// <see cref="IXLDataValidation.MinValue"/> still returns what was set.
    /// </para>
    /// <para>
    /// A criterion that is not a plain range refers to another sheet when a reference anywhere in it
    /// names one, as in <c>OFFSET(Data!$A$1,0,0,3,1)</c> or <c>B1&lt;=MAX(Data!$A$1:$A$3)</c>. Excel writes
    /// such a rule in the extension too (#536). The parser finds the references, so a defined name and
    /// text in a string, as in <c>INDIRECT("Data!A1")</c>, do not count. Text the parser refuses is
    /// written in the standard form, as it was before.
    /// </para>
    /// <para>
    /// A bare <c>#REF!</c> names no sheet, so it stays in the standard form. Excel keeps a list that a
    /// sheet delete left as <c>#REF!</c> in the extension, and has saved a list with the same text in
    /// the standard form, so the text alone cannot say which form Excel would use.
    /// </para>
    /// </remarks>
    private static (bool, string) UsesExternalSheet(XLWorksheet sheet, string value)
    {
        var formula = CalcEngine.FormulaText.WithoutLeadingEquals(value);
        if (!XLHelper.IsValidRangeAddress(formula))
            return (OtherSheetReferenceFinder.RefersToAnotherSheet(formula, sheet.Name), formula);

        var separatorIndex = formula.LastIndexOf('!');
        var hasSheet = separatorIndex >= 0;
        if (!hasSheet)
            return (false, formula);

        var sheetName = formula[..separatorIndex].UnescapeSheetName();
        if (XLHelper.SheetComparer.Equals(sheet.Name, sheetName))
        {
            // The spec wants us to include references to ranges on the same worksheet without the sheet name
            return (false, formula[(separatorIndex + 1)..]);
        }

        return (true, formula);
    }

    private static void WriteStandardDataValidations(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        List<(IXLDataValidation DataValidation, string MinValue, string MaxValue)> dataValidationsStandard)
    {
        // Save validations that don't use another sheet. It must have at least 1 child, XML doesn't allow 0.
        if (!dataValidationsStandard.Any(d => HasContent(d.DataValidation)))
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

                // Built from the model, so there is nothing loaded to keep: a default is left out.
                var dataValidation = new DataValidation
                {
                    AllowBlank = SchemaDefault.Bool(null, dv.IgnoreBlanks, false),
                    Formula1 = string.IsNullOrEmpty(minValue) ? null : new Formula1(minValue),
                    Formula2 = string.IsNullOrEmpty(maxValue) ? null : new Formula2(maxValue),
                    Type = SchemaDefault.Enum(null, dv.AllowedValues.ToOpenXml(), DataValidationValues.None),
                    ShowErrorMessage = SchemaDefault.Bool(null, dv.ShowErrorMessage, false),
                    Prompt = NullIfEmpty(dv.InputMessage),
                    PromptTitle = NullIfEmpty(dv.InputTitle),
                    ErrorTitle = NullIfEmpty(dv.ErrorTitle),
                    Error = NullIfEmpty(dv.ErrorMessage),
                    ShowDropDown = SchemaDefault.Bool(null, !dv.InCellDropdown, false),
                    ShowInputMessage = SchemaDefault.Bool(null, dv.ShowInputMessage, false),
                    ErrorStyle = SchemaDefault.Enum(null, dv.ErrorStyle.ToOpenXml(), DataValidationErrorStyleValues.Stop),
                    Operator = HasOperator(dv.AllowedValues)
                        ? SchemaDefault.Enum(null, dv.Operator.ToOpenXml(), DataValidationOperatorValues.Between)
                        : null,
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
        var extensionDataValidations = GetOrAddEmptyExtensionDataValidations(worksheet, cm, dataValidationsExtensionUri);

        foreach (var (dv, minValue, maxValue) in dataValidationsExtension)
            extensionDataValidations.AppendChild(CreateExtensionDataValidation(dv, minValue, maxValue));

        extensionDataValidations.Count = (uint)dataValidationsExtension.Count;
    }

    /// <summary>
    /// The sheet's <c>x14:dataValidations</c>, emptied, or added in a new extension (and the extension
    /// list, if the sheet has none) when the sheet has none or only an empty one.
    /// </summary>
    private static X14.DataValidations GetOrAddEmptyExtensionDataValidations(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
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
            var worksheetExtension = new WorksheetExtension { Uri = dataValidationsExtensionUri };
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

        return extensionDataValidations;
    }

    private static X14.DataValidation CreateExtensionDataValidation(IXLDataValidation dv, string minValue,
        string maxValue)
    {
        var sequence = string.Join(" ", dv.Ranges.Select(x => x.RangeAddress));
        return new X14.DataValidation
        {
            AllowBlank = SchemaDefault.Bool(null, dv.IgnoreBlanks, false),
            DataValidationForumla1 = !string.IsNullOrWhiteSpace(minValue)
                ? new X14.DataValidationForumla1(new OfficeExcel.Formula(minValue))
                : null,
            DataValidationForumla2 = !string.IsNullOrWhiteSpace(maxValue)
                ? new X14.DataValidationForumla2(new OfficeExcel.Formula(maxValue))
                : null,
            Type = SchemaDefault.Enum(null, dv.AllowedValues.ToOpenXml(), DataValidationValues.None),
            ShowErrorMessage = SchemaDefault.Bool(null, dv.ShowErrorMessage, false),
            Prompt = NullIfEmpty(dv.InputMessage),
            PromptTitle = NullIfEmpty(dv.InputTitle),
            ErrorTitle = NullIfEmpty(dv.ErrorTitle),
            Error = NullIfEmpty(dv.ErrorMessage),
            ShowDropDown = SchemaDefault.Bool(null, !dv.InCellDropdown, false),
            ShowInputMessage = SchemaDefault.Bool(null, dv.ShowInputMessage, false),
            ErrorStyle = SchemaDefault.Enum(null, dv.ErrorStyle.ToOpenXml(), DataValidationErrorStyleValues.Stop),
            Operator = HasOperator(dv.AllowedValues)
                ? SchemaDefault.Enum(null, dv.Operator.ToOpenXml(), DataValidationOperatorValues.Between)
                : null,
            ReferenceSequence = new OfficeExcel.ReferenceSequence { Text = sequence }
        };
    }

    /// <summary>
    /// Does the rule say anything? <see cref="IXLDataValidation.IsDirty"/>, or a title or message,
    /// shown or not.
    /// </summary>
    /// <remarks>
    /// Excel writes the text of an input or error message whose display is turned off, and leaves
    /// out <c>showInputMessage</c> or <c>showErrorMessage</c>. Such a rule loads with the message
    /// hidden, and <see cref="IXLDataValidation.IsDirty"/> does not count hidden text, so an "Any
    /// value" rule that has only that text would be dropped (#709).
    /// </remarks>
    private static bool HasContent(IXLDataValidation dv) =>
        dv.IsDirty()
        || !string.IsNullOrWhiteSpace(dv.InputTitle) || !string.IsNullOrWhiteSpace(dv.InputMessage)
        || !string.IsNullOrWhiteSpace(dv.ErrorTitle) || !string.IsNullOrWhiteSpace(dv.ErrorMessage);

    /// <summary>A message or title, left out when empty, as Excel leaves it out.</summary>
    private static StringValue? NullIfEmpty(string? text) => string.IsNullOrEmpty(text) ? null : new StringValue(text);

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
