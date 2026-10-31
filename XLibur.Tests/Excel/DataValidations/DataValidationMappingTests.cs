using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.DataValidations;

/// <summary>
/// A validation rule has two spellings: the standard <c>dataValidation</c> element and the x14
/// extension element. The reader normalises both back to one model, so a reload cannot see a
/// spelling difference. These tests therefore do both jobs: round-trip every setting through each
/// spelling, and read the emitted elements where the two must agree.
/// </summary>
public class DataValidationMappingTests
{
    private const string Other = "Other";

    /// <summary>
    /// One row per setting: how to set it to a non-default value, and how to read it back. The
    /// formula rows are form-aware, because the extension form is only chosen when a formula refers
    /// to another sheet or is too long.
    /// </summary>
    private sealed record Setting(string Name, Action<IXLDataValidation, bool> Set, Func<IXLDataValidation, object?> Get);

    private static readonly Setting[] Settings =
    [
        new("AllowedValues", (dv, _) => dv.AllowedValues = XLAllowedValues.Decimal, dv => dv.AllowedValues),
        new("Operator", (dv, _) => dv.Operator = XLOperator.NotBetween, dv => dv.Operator),
        new("MinValue", (dv, ext) => dv.MinValue = ext ? $"{Other}!$B$1" : "5", dv => dv.MinValue),
        new("MaxValue", (dv, ext) => dv.MaxValue = ext ? $"{Other}!$B$2" : "50", dv => dv.MaxValue),
        new("IgnoreBlanks", (dv, _) => dv.IgnoreBlanks = false, dv => dv.IgnoreBlanks),
        new("InCellDropdown", (dv, _) => dv.InCellDropdown = false, dv => dv.InCellDropdown),
        new("ShowInputMessage", (dv, _) => dv.ShowInputMessage = false, dv => dv.ShowInputMessage),
        new("ShowErrorMessage", (dv, _) => dv.ShowErrorMessage = false, dv => dv.ShowErrorMessage),
        new("InputTitle", (dv, _) => dv.InputTitle = "in title", dv => dv.InputTitle),
        new("InputMessage", (dv, _) => dv.InputMessage = "in message", dv => dv.InputMessage),
        new("ErrorTitle", (dv, _) => dv.ErrorTitle = "err title", dv => dv.ErrorTitle),
        new("ErrorMessage", (dv, _) => dv.ErrorMessage = "err message", dv => dv.ErrorMessage),
        new("ErrorStyle", (dv, _) => dv.ErrorStyle = XLErrorStyle.Warning, dv => dv.ErrorStyle),
    ];

    public static IEnumerable<(string Name, bool Extension)> PropertyByForm() =>
        Settings.SelectMany(s => new[] { (s.Name, false), (s.Name, true) });

    /// <summary>
    /// A rule whose settings are all at their defaults except the type, which has to be one that
    /// uses an operator and two formulas, so every setting can be seen to survive.
    /// </summary>
    private static void SetBaseline(IXLDataValidation dv, bool extension)
    {
        dv.AllowedValues = XLAllowedValues.WholeNumber;
        dv.Operator = XLOperator.Between;
        dv.MinValue = extension ? $"{Other}!$A$1" : "1";
        dv.MaxValue = extension ? $"{Other}!$A$2" : "10";
    }

    private static MemoryStream Save(Action<XLWorkbook, IXLWorksheet> build)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        wb.AddWorksheet(Other);
        build(wb, ws);
        var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms;
    }

    private static IXLDataValidation Single(MemoryStream ms, out XLWorkbook wb)
    {
        ms.Position = 0;
        wb = new XLWorkbook(ms);
        return wb.Worksheet("Data").DataValidations.Single();
    }

    private static (List<DataValidation> Standard, List<X14.DataValidation> Extension) ReadElements(MemoryStream ms)
    {
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var worksheet = doc.WorkbookPart!.WorksheetParts.First().Worksheet!;
        return (worksheet.Descendants<DataValidation>().ToList(),
            worksheet.Descendants<X14.DataValidation>().ToList());
    }

    private static Dictionary<string, string> Attributes(OpenXmlElement element) =>
        element.GetAttributes().ToDictionary(a => a.LocalName, a => a.Value ?? string.Empty);

    [Test]
    [MethodDataSource(nameof(PropertyByForm))]
    public async Task EverySetting_RoundTripsInBothSpellings(string name, bool extension)
    {
        var setting = Settings.Single(s => s.Name == name);
        using var ms = Save((_, ws) =>
        {
            var dv = ws.Range("A1:A3").CreateDataValidation();
            SetBaseline(dv, extension);
            setting.Set(dv, extension);
        });

        var (standard, ext) = ReadElements(ms);
        await Assert.That(ext.Count).IsEqualTo(extension ? 1 : 0).Because("the spelling used was not the one under test");
        await Assert.That(standard.Count).IsEqualTo(extension ? 0 : 1);

        var loaded = Single(ms, out var wb);
        using (wb)
        {
            var expected = new XLWorkbook().AddWorksheet("x").Range("A1").CreateDataValidation();
            SetBaseline(expected, extension);
            setting.Set(expected, extension);
            await Assert.That(setting.Get(loaded)).IsEqualTo(setting.Get(expected));
        }
    }

    [Test]
    public async Task EveryProperty_OfTheMapping_IsOnTheValidationInterface()
    {
        // The settings list is the single source for read, write, copy, reset and equality, so a
        // setting on the public interface without a matching component (or the reverse) is a defect.
        // Value is an alias of MinValue, not a setting of its own.
        var onInterface = typeof(IXLDataValidation)
            .GetProperties()
            .Where(p => p.CanWrite && p.Name != nameof(IXLDataValidation.Value))
            .Select(p => p.Name)
            .OrderBy(n => n)
            .ToArray();

        var mapped = typeof(XLDataValidation).Assembly
            .GetType("XLibur.Excel.DataValidationSettings", throwOnError: true)!
            .GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance)
            .Select(p => p.Name)
            .OrderBy(n => n)
            .ToArray();

        var tested = Settings.Select(s => s.Name).OrderBy(n => n).ToArray();

        await Assert.That(onInterface.Length).IsEqualTo(13);
        await Assert.That(mapped).IsEquivalentTo(onInterface);
        await Assert.That(tested).IsEquivalentTo(onInterface);
    }

    [Test]
    public async Task CrossSheetList_SurvivesSaveAndReload()
    {
        using var ms = Save((_, ws) => ws.Range("A1:A3").CreateDataValidation().List($"{Other}!$A$1:$A$5"));

        var (standard, ext) = ReadElements(ms);
        await Assert.That(standard).IsEmpty();
        await Assert.That(ext.Count).IsEqualTo(1);

        var loaded = Single(ms, out var wb);
        using (wb)
        {
            await Assert.That(loaded.AllowedValues).IsEqualTo(XLAllowedValues.List);
            await Assert.That(loaded.MinValue).IsEqualTo($"{Other}!$A$1:$A$5");
        }
    }

    [Test]
    public async Task FormulaLongerThan255_SurvivesSaveAndReload()
    {
        var formula = "\"" + string.Join(",", Enumerable.Range(0, 80).Select(i => $"item{i}")) + "\"";
        await Assert.That(formula.Length).IsGreaterThan(255);

        using var ms = Save((_, ws) => ws.Range("A1").CreateDataValidation().List(formula));

        var (standard, ext) = ReadElements(ms);
        await Assert.That(standard).IsEmpty();
        await Assert.That(ext.Count).IsEqualTo(1);

        var loaded = Single(ms, out var wb);
        using (wb)
            await Assert.That(loaded.MinValue).IsEqualTo(formula);
    }

    [Test]
    public async Task ASameSheetReference_StaysInTheStandardForm()
    {
        using var ms = Save((_, ws) => ws.Range("A1:A3").CreateDataValidation().List("Data!$C$1:$C$5"));

        var (standard, ext) = ReadElements(ms);
        await Assert.That(ext).IsEmpty();
        await Assert.That(standard.Single().Formula1!.Text).IsEqualTo("$C$1:$C$5");
    }

    /// <summary>
    /// A reload cannot see an empty-formula difference, so this reads the elements. Excel omits a
    /// formula element it has nothing to put in, and so do both spellings now.
    /// </summary>
    [Test]
    public async Task EmptyFormula_IsOmittedByBothSpellings()
    {
        var longFormula = "\"" + new string('a', 300) + "\"";
        using var ms = Save((_, ws) =>
        {
            var standardRule = ws.Range("A1").CreateDataValidation();
            standardRule.TextLength.EqualOrGreaterThan(1);
            standardRule.MaxValue = string.Empty;

            var extensionRule = ws.Range("B1").CreateDataValidation();
            extensionRule.TextLength.EqualOrGreaterThan(1);
            extensionRule.MinValue = longFormula;
            extensionRule.MaxValue = string.Empty;
        });

        var (standard, ext) = ReadElements(ms);
        await Assert.That(standard.Single().Formula1).IsNotNull();
        await Assert.That(standard.Single().Formula2).IsNull();
        await Assert.That(ext.Single().DataValidationForumla1).IsNotNull();
        await Assert.That(ext.Single().DataValidationForumla2).IsNull();
    }

    [Test]
    public async Task EmptyFormulas_OnARuleWithNoFormulaType_AreOmittedFromTheStandardForm()
    {
        using var ms = Save((_, ws) =>
        {
            var dv = ws.Range("A1").CreateDataValidation();
            dv.InputMessage = "hello";
        });

        var (standard, _) = ReadElements(ms);
        var element = standard.Single();
        await Assert.That(element.Formula1).IsNull();
        await Assert.That(element.Formula2).IsNull();
    }

    /// <summary>
    /// For a rule expressible in either form, the two spellings emit the same attributes: only the
    /// element name and the range attribute differ.
    /// </summary>
    [Test]
    public async Task BothSpellings_EmitTheSameAttributesForTheSameSettings()
    {
        using var ms = Save((_, ws) =>
        {
            foreach (var (address, formula) in new[] { ("A1", "1"), ("B1", "Other!$A$1") })
            {
                var dv = ws.Range(address).CreateDataValidation();
                dv.AllowedValues = XLAllowedValues.WholeNumber;
                dv.Operator = XLOperator.GreaterThan;
                dv.MinValue = formula;
                dv.IgnoreBlanks = false;
                dv.InCellDropdown = false;
                dv.InputTitle = "t";
                dv.InputMessage = "m";
                dv.ErrorTitle = "et";
                dv.ErrorMessage = "em";
                dv.ErrorStyle = XLErrorStyle.Information;
            }
        });

        var (standard, ext) = ReadElements(ms);
        var standardAttributes = Attributes(standard.Single());
        var extensionAttributes = Attributes(ext.Single());
        standardAttributes.Remove("sqref");
        extensionAttributes.Remove("sqref");

        await Assert.That(extensionAttributes).IsEquivalentTo(standardAttributes);
        await Assert.That(standardAttributes.Count).IsGreaterThan(8);
    }

    /// <summary>
    /// A rule that does nothing (any value, no messages) is not written. It used to be skipped only
    /// when no rule on the sheet's standard path had anything to say, and never on the extension path.
    /// </summary>
    [Test]
    public async Task ARuleThatDoesNothing_IsNotWrittenInEitherSpelling()
    {
        using var ms = Save((_, ws) =>
        {
            ws.Range("A1").CreateDataValidation().List("\"a,b\"");

            ws.Range("B1").CreateDataValidation(); // dirty-free, standard-expressible

            var crossSheet = ws.Range("C1").CreateDataValidation(); // dirty-free, extension-shaped
            crossSheet.MinValue = $"{Other}!$A$1";
        });

        var (standard, ext) = ReadElements(ms);
        await Assert.That(standard.Select(s => s.SequenceOfReferences!.InnerText).ToArray())
            .IsEquivalentTo(new string?[] { "A1:A1" });
        await Assert.That(ext).IsEmpty();
    }

    [Test]
    public async Task ExtensionRuleWithMessages_IsWrittenEvenWhenTypeIsAnyValue()
    {
        using var ms = Save((_, ws) =>
        {
            var dv = ws.Range("A1").CreateDataValidation();
            dv.InputMessage = "hello";
            dv.MinValue = $"{Other}!$A$1";
        });

        var (_, ext) = ReadElements(ms);
        await Assert.That(ext.Count).IsEqualTo(1);
    }

    [Test]
    [MethodDataSource(nameof(PropertyNames))]
    public async Task ConsolidationKeepsRulesApart_WhenOnlyOneSettingDiffers(string name)
    {
        var setting = Settings.Single(s => s.Name == name);
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");

        var first = ws.Range("A1").CreateDataValidation();
        var second = ws.Range("A2").CreateDataValidation();
        SetBaseline(first, false);
        SetBaseline(second, false);
        setting.Set(second, false);

        ((XLDataValidations)ws.DataValidations).Consolidate();

        await Assert.That(ws.DataValidations.Count()).IsEqualTo(2);
    }

    [Test]
    public async Task ConsolidationMergesRules_WhoseSettingsAreAllEqual()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        foreach (var address in new[] { "A1", "A2" })
        {
            var dv = ws.Range(address).CreateDataValidation();
            SetBaseline(dv, false);
            foreach (var setting in Settings)
                setting.Set(dv, false);
        }

        ((XLDataValidations)ws.DataValidations).Consolidate();

        await Assert.That(ws.DataValidations.Count()).IsEqualTo(1);
        await Assert.That(ws.DataValidations.Single().Ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:A2");
    }

    [Test]
    public async Task Copy_CarriesEverySetting()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet("Source");
        var target = wb.AddWorksheet("Target");

        var original = source.Range("A1:A3").CreateDataValidation();
        SetBaseline(original, false);
        foreach (var setting in Settings)
            setting.Set(original, false);

        var copy = target.DataValidations.Add(original);

        foreach (var setting in Settings)
            await Assert.That(setting.Get(copy)).IsEqualTo(setting.Get(original)).Because(setting.Name);
    }

    [Test]
    public async Task Clear_ResetsEverySettingToItsDefault()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        var fresh = ws.Range("B1").CreateDataValidation();
        var dv = ws.Range("A1").CreateDataValidation();
        foreach (var setting in Settings)
            setting.Set(dv, false);

        dv.Clear();

        foreach (var setting in Settings)
            await Assert.That(setting.Get(dv)).IsEqualTo(setting.Get(fresh)).Because(setting.Name);
    }

    [Test]
    public async Task ConsolidatedRuleThatWasEmptied_NeverReachesTheFile()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            var emptied = ws.Range("A1:A3").CreateDataValidation();
            emptied.WholeNumber.EqualOrGreaterThan(1);
            ws.Range("C1:C3").CreateDataValidation().WholeNumber.EqualOrGreaterThan(1);
            emptied.ClearRanges();

            wb.SaveAs(ms, new SaveOptions { ConsolidateDataValidationRanges = true });
        }

        var (standard, _) = ReadElements(ms);
        await Assert.That(standard.Select(s => s.SequenceOfReferences!.InnerText).ToArray())
            .IsEquivalentTo(new string?[] { "C1:C3" });
    }

    public static IEnumerable<string> PropertyNames() => Settings.Select(s => s.Name);
}
