using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Extensions;
using XLibur.Utils;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// How a data validation rule is spelled in a file. A rule has two spellings, the standard
/// <c>dataValidation</c> element and the <c>x14</c> extension element, and each is an adapter at the
/// <see cref="IDataValidationSpelling{TElement}"/> seam. Both map between an element and a
/// <see cref="DataValidationSettings"/>; which one a rule needs, and what the rule is written as,
/// are decided here and not in the writer.
/// </summary>
internal static class DataValidationMapping
{
    /// <summary>The longest formula the standard element holds.</summary>
    internal const int MaxStandardFormulaLength = 255;

    /// <summary>The uri of the worksheet extension that holds the <c>x14</c> data validations.</summary>
    internal const string ExtensionUri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";

    /// <summary>A rule as it is written: what to write, where it applies, and in which spelling.</summary>
    /// <param name="Settings">The settings, with the criteria as the file stores them.</param>
    /// <param name="Sqref">The cells the rule applies to.</param>
    /// <param name="NeedsExtension">Whether only the extension spelling can hold the rule.</param>
    internal readonly record struct Resolved(DataValidationSettings Settings, string Sqref, bool NeedsExtension);

    /// <summary>
    /// Works out how <paramref name="dataValidation"/> is written, or says it is not written at all.
    /// The same decision is made for both spellings.
    /// </summary>
    /// <remarks>
    /// A rule is not written when it applies to no cell, or when it says nothing. A rule can be left
    /// covering nothing, because <c>ClearRanges</c> and <c>RemoveRange</c> are public and neither drops
    /// the rule. The ranges are joined into <c>sqref</c>, which the schema forbids empty and Excel
    /// repairs the file over, so there is nothing to write. A rule that is not
    /// <see cref="DataValidationSettings.HasContent"/> restricts nothing and shows nothing.
    /// </remarks>
    internal static bool TryResolve(XLWorksheet sheet, XLDataValidation dataValidation, out Resolved resolved)
    {
        resolved = default;
        var settings = dataValidation.Settings;
        if (dataValidation.Areas.Count == 0 || !settings.HasContent())
            return false;

        var (minOnAnotherSheet, minValue) = UsesExternalSheet(sheet, settings.MinValue);
        var (maxOnAnotherSheet, maxValue) = UsesExternalSheet(sheet, settings.MaxValue);
        settings = settings with { MinValue = minValue, MaxValue = maxValue };

        var sqref = string.Join(" ", dataValidation.Ranges.Select(x => x.RangeAddress));
        resolved = new Resolved(settings, sqref, NeedsExtension(settings, minOnAnotherSheet || maxOnAnotherSheet));
        return true;
    }

    /// <summary>
    /// The standard element holds neither a formula that refers to another sheet nor one longer than
    /// <see cref="MaxStandardFormulaLength"/>, so a rule with either is written as an extension.
    /// </summary>
    private static bool NeedsExtension(in DataValidationSettings settings, bool refersToAnotherSheet) =>
        refersToAnotherSheet
        || settings.MinValue.Length > MaxStandardFormulaLength
        || settings.MaxValue.Length > MaxStandardFormulaLength;

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

    /// <summary>
    /// The settings a file that leaves every setting out means: the value the schema gives each one.
    /// This is not <see cref="DataValidationSettings.Default"/>, because the schema defaults three
    /// flags to false, which a rule built in code has true (#709).
    /// </summary>
    private static readonly DataValidationSettings Absent = DataValidationSettings.Default with
    {
        IgnoreBlanks = false,
        ShowErrorMessage = false,
        ShowInputMessage = false,
    };

    /// <summary>A formula, or null for the element a rule with no formula leaves out, as Excel does.</summary>
    private static string? FormulaOrNull(string formula) => string.IsNullOrWhiteSpace(formula) ? null : formula;

    /// <summary>A message or title, left out when empty, as Excel leaves it out.</summary>
    private static StringValue? NullIfEmpty(string? text) => string.IsNullOrEmpty(text) ? null : new StringValue(text);

    /// <summary>The standard <c>dataValidation</c> element.</summary>
    internal sealed class StandardSpelling : IDataValidationSpelling<DataValidation>
    {
        public static string? Sqref(DataValidation element) => element.SequenceOfReferences?.InnerText;

        public static DataValidationSettings Read(DataValidation element)
        {
            var settings = Absent;
            settings = settings with
            {
                IgnoreBlanks = OpenXmlHelper.GetBooleanValueAsBool(element.AllowBlank, false),
                ShowErrorMessage = OpenXmlHelper.GetBooleanValueAsBool(element.ShowErrorMessage, false),
                ShowInputMessage = OpenXmlHelper.GetBooleanValueAsBool(element.ShowInputMessage, false),
            };
            if (element.ShowDropDown != null) settings = settings with { InCellDropdown = !element.ShowDropDown.Value };
            if (element.PromptTitle != null) settings = settings with { InputTitle = element.PromptTitle.Value! };
            if (element.Prompt != null) settings = settings with { InputMessage = element.Prompt.Value! };
            if (element.ErrorTitle != null) settings = settings with { ErrorTitle = element.ErrorTitle.Value! };
            if (element.Error != null) settings = settings with { ErrorMessage = element.Error.Value! };
            if (element.ErrorStyle != null) settings = settings with { ErrorStyle = element.ErrorStyle.Value.ToXLibur() };
            if (element.Type != null) settings = settings with { AllowedValues = element.Type.Value.ToXLibur() };
            if (element.Operator != null) settings = settings with { Operator = element.Operator.Value.ToXLibur() };
            if (element.Formula1 != null) settings = settings with { MinValue = element.Formula1.Text };
            if (element.Formula2 != null) settings = settings with { MaxValue = element.Formula2.Text };
            return settings;
        }

        public static DataValidation Write(in DataValidationSettings settings, string sqref)
        {
            var minValue = FormulaOrNull(settings.MinValue);
            var maxValue = FormulaOrNull(settings.MaxValue);

            // Built from the model, so there is nothing loaded to keep: a default is left out.
            return new DataValidation
            {
                AllowBlank = SchemaDefault.Bool(null, settings.IgnoreBlanks, false),
                Formula1 = minValue is null ? null : new Formula1(minValue),
                Formula2 = maxValue is null ? null : new Formula2(maxValue),
                Type = SchemaDefault.Enum(null, settings.AllowedValues.ToOpenXml(), DataValidationValues.None),
                ShowErrorMessage = SchemaDefault.Bool(null, settings.ShowErrorMessage, false),
                Prompt = NullIfEmpty(settings.InputMessage),
                PromptTitle = NullIfEmpty(settings.InputTitle),
                ErrorTitle = NullIfEmpty(settings.ErrorTitle),
                Error = NullIfEmpty(settings.ErrorMessage),
                ShowDropDown = SchemaDefault.Bool(null, !settings.InCellDropdown, false),
                ShowInputMessage = SchemaDefault.Bool(null, settings.ShowInputMessage, false),
                ErrorStyle = SchemaDefault.Enum(null, settings.ErrorStyle.ToOpenXml(), DataValidationErrorStyleValues.Stop),
                Operator = settings.UsesOperator()
                    ? SchemaDefault.Enum(null, settings.Operator.ToOpenXml(), DataValidationOperatorValues.Between)
                    : null,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = sqref }
            };
        }
    }

    /// <summary>The <c>x14:dataValidation</c> element of the worksheet extension.</summary>
    internal sealed class ExtensionSpelling : IDataValidationSpelling<X14.DataValidation>
    {
        public static string? Sqref(X14.DataValidation element) => element.ReferenceSequence?.InnerText;

        public static DataValidationSettings Read(X14.DataValidation element)
        {
            var settings = Absent;
            settings = settings with
            {
                IgnoreBlanks = OpenXmlHelper.GetBooleanValueAsBool(element.AllowBlank, false),
                ShowErrorMessage = OpenXmlHelper.GetBooleanValueAsBool(element.ShowErrorMessage, false),
                ShowInputMessage = OpenXmlHelper.GetBooleanValueAsBool(element.ShowInputMessage, false),
            };
            if (element.ShowDropDown != null) settings = settings with { InCellDropdown = !element.ShowDropDown.Value };
            if (element.PromptTitle != null) settings = settings with { InputTitle = element.PromptTitle.Value! };
            if (element.Prompt != null) settings = settings with { InputMessage = element.Prompt.Value! };
            if (element.ErrorTitle != null) settings = settings with { ErrorTitle = element.ErrorTitle.Value! };
            if (element.Error != null) settings = settings with { ErrorMessage = element.Error.Value! };
            if (element.ErrorStyle != null) settings = settings with { ErrorStyle = element.ErrorStyle.Value.ToXLibur() };
            if (element.Type != null) settings = settings with { AllowedValues = element.Type.Value.ToXLibur() };
            if (element.Operator != null) settings = settings with { Operator = element.Operator.Value.ToXLibur() };
            if (element.DataValidationForumla1 != null) settings = settings with { MinValue = element.DataValidationForumla1.InnerText };
            if (element.DataValidationForumla2 != null) settings = settings with { MaxValue = element.DataValidationForumla2.InnerText };
            return settings;
        }

        public static X14.DataValidation Write(in DataValidationSettings settings, string sqref)
        {
            var minValue = FormulaOrNull(settings.MinValue);
            var maxValue = FormulaOrNull(settings.MaxValue);

            return new X14.DataValidation
            {
                AllowBlank = SchemaDefault.Bool(null, settings.IgnoreBlanks, false),
                DataValidationForumla1 = minValue is null
                    ? null
                    : new X14.DataValidationForumla1(new OfficeExcel.Formula(minValue)),
                DataValidationForumla2 = maxValue is null
                    ? null
                    : new X14.DataValidationForumla2(new OfficeExcel.Formula(maxValue)),
                Type = SchemaDefault.Enum(null, settings.AllowedValues.ToOpenXml(), DataValidationValues.None),
                ShowErrorMessage = SchemaDefault.Bool(null, settings.ShowErrorMessage, false),
                Prompt = NullIfEmpty(settings.InputMessage),
                PromptTitle = NullIfEmpty(settings.InputTitle),
                ErrorTitle = NullIfEmpty(settings.ErrorTitle),
                Error = NullIfEmpty(settings.ErrorMessage),
                ShowDropDown = SchemaDefault.Bool(null, !settings.InCellDropdown, false),
                ShowInputMessage = SchemaDefault.Bool(null, settings.ShowInputMessage, false),
                ErrorStyle = SchemaDefault.Enum(null, settings.ErrorStyle.ToOpenXml(), DataValidationErrorStyleValues.Stop),
                Operator = settings.UsesOperator()
                    ? SchemaDefault.Enum(null, settings.Operator.ToOpenXml(), DataValidationOperatorValues.Between)
                    : null,
                ReferenceSequence = new OfficeExcel.ReferenceSequence { Text = sqref }
            };
        }
    }
}

/// <summary>
/// One spelling of a data validation rule in a file: the seam <see cref="DataValidationMapping"/> has
/// two adapters at.
/// </summary>
/// <typeparam name="TElement">The element that holds a rule in this spelling.</typeparam>
internal interface IDataValidationSpelling<TElement>
    where TElement : OpenXmlElement
{
    /// <summary>The cells the element applies to, as the file spells them, or null when it has none.</summary>
    static abstract string? Sqref(TElement element);

    /// <summary>The settings the element holds, with the schema's value for each one it leaves out.</summary>
    static abstract DataValidationSettings Read(TElement element);

    /// <summary>The element that holds <paramref name="settings"/>, leaving out what only restates the schema.</summary>
    static abstract TElement Write(in DataValidationSettings settings, string sqref);
}
