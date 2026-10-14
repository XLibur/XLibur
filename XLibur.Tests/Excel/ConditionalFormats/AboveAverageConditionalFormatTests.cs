using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// #743. An above- or below-average rule (<c>cfRule type="aboveAverage"</c>) loaded, but the save threw
/// <see cref="System.NotImplementedException"/>, because no converter wrote it. The load also dropped the
/// rule's <c>aboveAverage</c>, <c>equalAverage</c> and <c>stdDev</c> attributes, so a converter alone would
/// have written every such rule as "above average".
/// </summary>
/// <remarks>
/// LibreOffice's <c>new_cond_format_test.xlsx</c> and its re-export hold, on <c>Sheet1</c>, four top-10
/// rules and four average rules: above (<c>J2:J22</c>), below (<c>L2:L22</c>), equal or above
/// (<c>N2:N22</c>) and equal or below (<c>P2:P22</c>).
/// </remarks>
public class AboveAverageConditionalFormatTests
{
    private const string Folder = @"TryToLoad\LO\xlsx\";
    private const string Fixture = "new_cond_format_test.xlsx";
    private const string ExportFixture = "new_cond_format_test_export.xlsx";

    [Test]
    public async Task The_fixture_holds_the_average_rules_the_tests_expect()
    {
        using var source = TestHelper.OpenResource(Folder + Fixture);

        var averageRules = Rules(source).Where(r => r.Contains("aboveAverage")).ToList();

        await Assert.That(averageRules).IsEquivalentTo(new[]
        {
            "Sheet1!J2:J22 p4 aboveAverage dxf=0.00",
            "Sheet1!L2:L22 p3 aboveAverage aboveAverage=0 dxf=0.00",
            "Sheet1!N2:N22 p2 aboveAverage equalAverage=1 dxf=0.00",
            "Sheet1!P2:P22 p1 aboveAverage aboveAverage=0 equalAverage=1 dxf=0.00",
        });
    }

    [Test]
    [Arguments(Fixture)]
    [Arguments(ExportFixture)]
    public async Task A_load_and_save_writes_every_rule_as_the_file_has_it(string fileName)
    {
        using var source = TestHelper.OpenResource(Folder + fileName);
        var expected = Rules(source);

        using var saved = LoadAndSave(source);

        await Assert.That(Rules(saved)).IsEquivalentTo(expected);
    }

    /// <summary>
    /// The save writes an average rule's <c>aboveAverage</c> and <c>equalAverage</c> only when the file
    /// had them, as Excel and LibreOffice do: only when they differ from the schema default.
    /// </summary>
    [Test]
    public async Task The_save_writes_an_average_rule_with_the_attributes_the_file_has()
    {
        using var source = TestHelper.OpenResource(Folder + Fixture);
        var expected = AverageRuleAttributes(source);

        using var saved = LoadAndSave(source);

        await Assert.That(AverageRuleAttributes(saved)).IsEquivalentTo(expected);
    }

    [Test]
    [Arguments(Fixture)]
    [Arguments(ExportFixture)]
    public async Task The_saved_file_loads_and_saves_again_unchanged(string fileName)
    {
        using var source = TestHelper.OpenResource(Folder + fileName);
        using var saved = LoadAndSave(source);
        var expected = Rules(saved);

        using var resaved = LoadAndSave(saved);

        await Assert.That(Rules(resaved)).IsEquivalentTo(expected);
    }

    [Test]
    public async Task A_rule_with_a_standard_deviation_keeps_it_through_a_load_and_save()
    {
        using var source = TestHelper.OpenResource(Folder + Fixture);
        SetRuleAttributes(source, "L2:L22", rule => rule.StdDev = 2);

        using var saved = LoadAndSave(source);

        await Assert.That(Rules(saved))
            .Contains("Sheet1!L2:L22 p3 aboveAverage aboveAverage=0 stdDev=2 dxf=0.00");
    }

    [Test]
    public async Task A_load_keeps_which_values_each_average_rule_formats()
    {
        using var wb = new XLWorkbook(TestHelper.OpenResource(Folder + Fixture));

        var rules = AverageRules(wb.Worksheet("Sheet1"));

        await Assert.That(rules).IsEquivalentTo(new[]
        {
            "J2:J22 above=True equal=False stdDev=",
            "L2:L22 above=False equal=False stdDev=",
            "N2:N22 above=True equal=True stdDev=",
            "P2:P22 above=False equal=True stdDev=",
        });
    }

    [Test]
    public async Task A_copy_of_a_rule_to_another_sheet_keeps_which_values_it_formats()
    {
        using var source = TestHelper.OpenResource(Folder + Fixture);
        SetRuleAttributes(source, "P2:P22", rule => rule.StdDev = 3);
        using var wb = new XLWorkbook(source);
        var target = wb.AddWorksheet("Target");

        foreach (var cf in wb.Worksheet("Sheet1").ConditionalFormats
                     .Where(cf => cf.ConditionalFormatType == XLConditionalFormatType.AboveAverage))
            cf.CopyTo(target);

        await Assert.That(AverageRules(target)).IsEquivalentTo(new[]
        {
            "J2:J22 above=True equal=False stdDev=",
            "L2:L22 above=False equal=False stdDev=",
            "N2:N22 above=True equal=True stdDev=",
            "P2:P22 above=False equal=True stdDev=3",
        });
    }

    [Test]
    public async Task Two_average_rules_that_format_different_values_are_not_equal()
    {
        using var wb = new XLWorkbook(TestHelper.OpenResource(Folder + Fixture));
        var rules = wb.Worksheet("Sheet1").ConditionalFormats.Cast<XLConditionalFormat>()
            .Where(cf => cf.ConditionalFormatType == XLConditionalFormatType.AboveAverage)
            .ToList();

        var distinct = rules.Distinct(XLConditionalFormat.NoRangeComparer).Count();

        await Assert.That(distinct).IsEqualTo(4);
    }

    [Test]
    public async Task The_saved_file_passes_validation()
    {
        using var wb = new XLWorkbook(TestHelper.OpenResource(Folder + Fixture));

        using var saved = TestHelper.SaveValidated(wb);

        await Assert.That(Rules(saved).Count(r => r.Contains("aboveAverage"))).IsEqualTo(4);
    }

    /// <summary>
    /// The above- and below-average rules of <paramref name="sheet"/> as the model holds them, one line each.
    /// </summary>
    private static List<string> AverageRules(IXLWorksheet sheet)
        => sheet.ConditionalFormats.Cast<XLConditionalFormat>()
            .Where(cf => cf.ConditionalFormatType == XLConditionalFormatType.AboveAverage)
            .Select(cf => $"{cf.Ranges.Single().RangeAddress.ToStringRelative()} above={cf.AboveAverage} " +
                          $"equal={cf.EqualAverage} stdDev={cf.StdDev}")
            .ToList();

    /// <summary>Loads <paramref name="source"/>, saves it unchanged, and returns the saved package.</summary>
    private static MemoryStream LoadAndSave(Stream source)
    {
        var ms = new MemoryStream();
        source.Position = 0;
        using (var wb = new XLWorkbook(source))
            wb.SaveAs(ms);

        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// Sets attributes on the standard rule of <c>Sheet1</c> whose range is <paramref name="sqref"/>, in
    /// place in <paramref name="package"/>.
    /// </summary>
    private static void SetRuleAttributes(MemoryStream package, string sqref,
        System.Action<S.ConditionalFormattingRule> edit)
    {
        using (var document = SpreadsheetDocument.Open(package, true))
        {
            var rule = SheetParts(document)
                .Single(p => p.Name == "Sheet1").Part.Worksheet!
                .Elements<S.ConditionalFormatting>()
                .Single(cf => cf.SequenceOfReferences!.InnerText == sqref)
                .Elements<S.ConditionalFormattingRule>()
                .Single();
            edit(rule);
        }

        package.Position = 0;
    }

    /// <summary>
    /// Every standard conditional format rule in <paramref name="package"/>, one line each: its sheet and
    /// range, priority, type, the attributes that define it, its formulas, and the number format of its
    /// differential format (all the fixtures' formats are number formats). The format is named by what it
    /// holds, not its index, because a save may number the differential formats differently.
    /// </summary>
    private static List<string> Rules(Stream package)
    {
        package.Position = 0;
        var lines = new List<string>();
        using (var document = SpreadsheetDocument.Open(package, false))
        {
            var dxfs = document.WorkbookPart!.WorkbookStylesPart?.Stylesheet!.DifferentialFormats?
                .Elements<S.DifferentialFormat>().ToList() ?? [];

            foreach (var (name, part) in SheetParts(document))
            {
                foreach (var cf in part.Worksheet!.Elements<S.ConditionalFormatting>())
                {
                    foreach (var rule in cf.Elements<S.ConditionalFormattingRule>())
                        lines.Add($"{name}!{cf.SequenceOfReferences!.InnerText} {Describe(rule, dxfs)}");
                }
            }
        }

        package.Position = 0;
        return lines;
    }

    /// <summary>
    /// Each average rule on <c>Sheet1</c> of <paramref name="package"/>: its range and the attributes it is
    /// written with, other than its differential format's index, by name and value as written.
    /// </summary>
    private static List<string> AverageRuleAttributes(Stream package)
    {
        package.Position = 0;
        var lines = new List<string>();
        using (var document = SpreadsheetDocument.Open(package, false))
        {
            var sheet = SheetParts(document).Single(p => p.Name == "Sheet1").Part.Worksheet!;
            foreach (var cf in sheet.Elements<S.ConditionalFormatting>())
            {
                foreach (var rule in cf.Elements<S.ConditionalFormattingRule>()
                             .Where(r => r.Type?.InnerText == "aboveAverage"))
                {
                    var attributes = rule.GetAttributes()
                        .Where(a => a.LocalName != "dxfId")
                        .Select(a => $"{a.LocalName}={a.Value}")
                        .Order();
                    lines.Add($"{cf.SequenceOfReferences!.InnerText} {string.Join(" ", attributes)}" +
                              $" children={rule.ChildElements.Count}");
                }
            }
        }

        package.Position = 0;
        return lines;
    }

    private static string Describe(S.ConditionalFormattingRule rule, List<S.DifferentialFormat> dxfs)
    {
        // A flag is named only when it differs from its schema default, so that a rule written with the
        // default spelled out (XLibur writes a top-10 rule's percent="0") reads as the rule without it.
        var parts = new List<string> { $"p{rule.Priority?.Value}", rule.Type!.InnerText! };
        if (rule.AboveAverage?.Value == false) parts.Add("aboveAverage=0");
        if (rule.EqualAverage?.Value == true) parts.Add("equalAverage=1");
        if (rule.StdDev is not null) parts.Add($"stdDev={rule.StdDev.Value}");
        if (rule.Percent?.Value == true) parts.Add("percent=1");
        if (rule.Bottom?.Value == true) parts.Add("bottom=1");
        if (rule.Rank is not null) parts.Add($"rank={rule.Rank.Value}");
        if (rule.StopIfTrue?.Value == true) parts.Add("stopIfTrue=1");
        parts.AddRange(rule.Elements<S.Formula>().Select(f => $"formula={f.Text}"));
        if (rule.FormatId is not null)
            parts.Add($"dxf={dxfs[(int)rule.FormatId.Value].NumberingFormat?.FormatCode?.Value}");

        return string.Join(" ", parts);
    }

    private static IEnumerable<(string Name, WorksheetPart Part)> SheetParts(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        foreach (var sheet in workbookPart.Workbook!.Sheets!.Elements<S.Sheet>())
        {
            if (workbookPart.GetPartById(sheet.Id!.Value!) is WorksheetPart part)
                yield return (sheet.Name!.Value!, part);
        }
    }
}
