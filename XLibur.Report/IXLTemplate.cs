using System;
using System.IO;
using XLibur.Excel;

namespace XLibur.Report;

/// <summary>
/// An Excel workbook used as a report template: data is bound to it by name, then
/// <see cref="Generate"/> resolves the template's expressions and expands its ranges in place.
/// </summary>
public interface IXLTemplate : IDisposable
{
    /// <summary>The workbook being generated into.</summary>
    IXLWorkbook Workbook { get; }

    /// <summary>Whether <see cref="Generate"/> has run.</summary>
    bool IsGenerated { get; }

    /// <summary>Makes <paramref name="value"/> visible to the template under <paramref name="alias"/>.</summary>
    void AddVariable(string alias, object? value);

    /// <summary>
    /// Makes each public property and field of <paramref name="value"/> visible to the template
    /// under its own name. A dictionary contributes its entries instead.
    /// </summary>
    void AddVariable(object value);

    /// <summary>
    /// Resolves the template against the variables added so far.
    /// </summary>
    /// <returns>
    /// The errors encountered, if any. Generation does not throw on a bad expression — the
    /// offending cell records the problem and the rest of the report still builds.
    /// </returns>
    XLGenerateResult Generate();

    /// <summary>Saves the generated workbook to <paramref name="file"/>.</summary>
    void SaveAs(string file);

    /// <summary>Saves the generated workbook to <paramref name="stream"/>.</summary>
    void SaveAs(Stream stream);
}
