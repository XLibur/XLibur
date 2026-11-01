using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads data validation rules, from the standard <c>dataValidations</c> element and from the
/// <c>x14</c> extension. Both go through <see cref="DataValidationMapping"/>.
/// </summary>
internal static class DataValidationReader
{
    /// <summary>Loads the rules of the standard <c>dataValidations</c> element.</summary>
    internal static void LoadDataValidations(DataValidations dataValidations, XLWorksheet ws)
    {
        ArgumentNullException.ThrowIfNull(ws);

        Load<DataValidationMapping.StandardSpelling, DataValidation>(
            dataValidations.Elements<DataValidation>(), ws);
    }

    /// <summary>Loads the rules of the <c>x14</c> data validations a worksheet extension holds.</summary>
    internal static void LoadExtensionDataValidations(WorksheetExtensionList extensions, XLWorksheet ws)
    {
        Load<DataValidationMapping.ExtensionSpelling, X14.DataValidation>(
            extensions
                .Descendants<X14.DataValidations>()
                .SelectMany(dataValidations => dataValidations.Descendants<X14.DataValidation>()),
            ws);
    }

    /// <summary>
    /// One rule per range, as a file lists several ranges for the rules that share settings.
    /// </summary>
    private static void Load<TSpelling, TElement>(IEnumerable<TElement> elements, XLWorksheet ws)
        where TSpelling : IDataValidationSpelling<TElement>
        where TElement : OpenXmlElement
    {
        foreach (var element in elements)
        {
            var sqref = TSpelling.Sqref(element);
            if (string.IsNullOrWhiteSpace(sqref)) continue;

            var settings = TSpelling.Read(element);
            foreach (var rangeAddress in sqref.Split(' '))
            {
                var dvt = new XLDataValidation(ws.Range(rangeAddress)!);
                ws.DataValidations.Add(dvt, skipIntersectionsCheck: true);
                dvt.Settings = settings;
            }
        }
    }
}
