using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// The text rules put their text inside a formula string literal, where a <c>"</c> must be doubled
/// (issue #660). Measured over COM: Excel 16, given the text <c>say "hi"</c> on <c>A1:A5</c>, keeps
/// the raw text in the <c>text</c> attribute and doubles the quote in the formula only, e.g.
/// <c>NOT(ISERROR(SEARCH("say ""hi""",A1)))</c>. XLibur writes the length of the raw text where
/// Excel writes <c>LEN("say ""hi""")</c>; the two are equal.
/// </summary>
public class ConditionalFormatTextRuleQuoteTests
{
    private const string Text = "say \"hi\"";

    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Test]
    [Arguments(XLConditionalFormatType.ContainsText, "containsText", "NOT(ISERROR(SEARCH(\"say \"\"hi\"\"\",A1)))")]
    [Arguments(XLConditionalFormatType.NotContainsText, "notContainsText", "ISERROR(SEARCH(\"say \"\"hi\"\"\",A1))")]
    [Arguments(XLConditionalFormatType.StartsWith, "beginsWith", "LEFT(A1,8)=\"say \"\"hi\"\"\"")]
    [Arguments(XLConditionalFormatType.EndsWith, "endsWith", "RIGHT(A1,8)=\"say \"\"hi\"\"\"")]
    public async Task A_quote_in_the_text_is_doubled_in_the_formula_and_survives_a_round_trip(
        XLConditionalFormatType type, string xmlType, string expectedFormula)
    {
        using var firstSave = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            AddTextRule(ws.Range("A1:A5").AddConditionalFormat(), type, Text)
                .Fill.SetBackgroundColor(XLColor.Red);
            wb.SaveAs(firstSave);
        }

        var (savedType, savedText, savedFormula) = SavedRule(firstSave);
        await Assert.That(savedType).IsEqualTo(xmlType);
        await Assert.That(savedText).IsEqualTo(Text);
        await Assert.That(savedFormula).IsEqualTo(expectedFormula);

        using var secondSave = new MemoryStream();
        using (var wb = new XLWorkbook(firstSave))
        {
            var cf = wb.Worksheet(1).ConditionalFormats.Single();
            await Assert.That(cf.ConditionalFormatType).IsEqualTo(type);
            await Assert.That(cf.Values[1].Value).IsEqualTo(Text);
            await Assert.That(cf.Values[1].IsFormula).IsFalse();
            wb.SaveAs(secondSave);
        }

        await Assert.That(SavedRule(secondSave)).IsEqualTo((xmlType, Text, expectedFormula));
    }

    private static IXLStyle AddTextRule(IXLConditionalFormat cf, XLConditionalFormatType type, string text) =>
        type switch
        {
            XLConditionalFormatType.ContainsText => cf.WhenContains(text),
            XLConditionalFormatType.NotContainsText => cf.WhenNotContains(text),
            XLConditionalFormatType.StartsWith => cf.WhenStartsWith(text),
            XLConditionalFormatType.EndsWith => cf.WhenEndsWith(text),
            _ => throw new ArgumentOutOfRangeException(nameof(type), type, null),
        };

    private static (string? Type, string? Text, string Formula) SavedRule(Stream package)
    {
        var rule = XDocument.Parse(package.Sheet1Xml()).Descendants(Main + "cfRule").Single();
        return ((string?)rule.Attribute("type"), (string?)rule.Attribute("text"),
            rule.Elements(Main + "formula").Single().Value);
    }
}
