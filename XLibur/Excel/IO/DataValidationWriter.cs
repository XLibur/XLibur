using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes a worksheet's data validations, in the standard element or the <c>x14</c> extension. What a
/// rule is written as, and which of the two holds it, is decided by <see cref="DataValidationMapping"/>;
/// this class only places the elements in the worksheet.
/// </summary>
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
        var standard = new List<DataValidationMapping.Resolved>();
        var extension = new List<DataValidationMapping.Resolved>();
        if (options.ConsolidateDataValidationRanges)
        {
            xlWorksheet.DataValidations.Consolidate();
        }

        foreach (var dv in xlWorksheet.DataValidations.Cast<XLDataValidation>())
        {
            // The same rules decide whether a rule is written, whichever element it goes in.
            if (!DataValidationMapping.TryResolve(xlWorksheet, dv, out var resolved))
                continue;

            (resolved.NeedsExtension ? extension : standard).Add(resolved);
        }

        WriteStandardDataValidations(worksheet, cm, standard);
        WriteExtensionDataValidations(worksheet, cm, extension);
    }

    private static void WriteStandardDataValidations(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        List<DataValidationMapping.Resolved> rules)
    {
        // It must have at least 1 child, XML doesn't allow 0.
        if (rules.Count == 0)
        {
            worksheet.RemoveAllChildren<DataValidations>();
            cm.SetElement(XLWorksheetContents.DataValidations, null);
            return;
        }

        if (!worksheet.Elements<DataValidations>().Any())
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.DataValidations);
            worksheet.InsertAfter(new DataValidations(), previousElement);
        }

        var dataValidations = worksheet.Elements<DataValidations>().First();
        cm.SetElement(XLWorksheetContents.DataValidations, dataValidations);
        var loadedSqrefs = new SqrefSpellings(dataValidations.Elements<DataValidation>()
            .Select(DataValidationMapping.StandardSpelling.Sqref));
        dataValidations.RemoveAllChildren<DataValidation>();

        foreach (var rule in rules)
        {
            dataValidations.AppendChild(
                DataValidationMapping.StandardSpelling.Write(rule.Settings, loadedSqrefs.Spell(rule.Sqref)));
        }

        dataValidations.Count = (uint)rules.Count;
    }

    private static void WriteExtensionDataValidations(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        List<DataValidationMapping.Resolved> rules)
    {
        if (rules.Count == 0)
        {
            RemoveExtensionDataValidations(worksheet, cm);
        }
        else
        {
            WriteExtensionDataValidationElements(worksheet, cm, rules);
        }
    }

    private static void RemoveExtensionDataValidations(Worksheet worksheet, XLWorksheetContentManager cm)
    {
        var worksheetExtensionList = worksheet.Elements<WorksheetExtensionList>().FirstOrDefault();
        var worksheetExtension = worksheetExtensionList?.Elements<WorksheetExtension>()
            .FirstOrDefault(ext =>
                string.Equals(ext.Uri, DataValidationMapping.ExtensionUri, StringComparison.OrdinalIgnoreCase));

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
        List<DataValidationMapping.Resolved> rules)
    {
        var extensionDataValidations = GetOrAddEmptyExtensionDataValidations(worksheet, cm, out var loaded);
        var loadedSqrefs = new SqrefSpellings(loaded);

        foreach (var rule in rules)
        {
            extensionDataValidations.AppendChild(
                DataValidationMapping.ExtensionSpelling.Write(rule.Settings, loadedSqrefs.Spell(rule.Sqref)));
        }

        extensionDataValidations.Count = (uint)rules.Count;
    }

    /// <summary>
    /// The sheet's <c>x14:dataValidations</c>, emptied, or added in a new extension (and the extension
    /// list, if the sheet has none) when the sheet has none or only an empty one.
    /// </summary>
    /// <param name="worksheet">The worksheet to write to.</param>
    /// <param name="cm">Where each worksheet element is.</param>
    /// <param name="loaded">The <c>sqref</c> of each rule the element held before it was emptied.</param>
    private static X14.DataValidations GetOrAddEmptyExtensionDataValidations(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        out List<string?> loaded)
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
            loaded = [];

            var worksheetExtension = new WorksheetExtension { Uri = DataValidationMapping.ExtensionUri };
            worksheetExtension.AddNamespaceDeclaration("x14", X14Main2009SsNs);
            worksheetExtensionList.Append(worksheetExtension);

            extensionDataValidations = new X14.DataValidations();
            extensionDataValidations.AddNamespaceDeclaration("xm", XmMain2006);
            worksheetExtension.Append(extensionDataValidations);
        }
        else
        {
            loaded = extensionDataValidations.Elements<X14.DataValidation>()
                .Select(DataValidationMapping.ExtensionSpelling.Sqref)
                .ToList();
            extensionDataValidations.RemoveAllChildren();
        }

        return extensionDataValidations;
    }
}
