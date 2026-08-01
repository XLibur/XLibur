using System.Linq;
using XLibur.Excel;
using XLibur.Excel.RichText;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.RichText;

public class XLImmutableRichTextTests
{
    [Test]
    public async Task Equals_compares_text_runs_phonetic_runs_and_properties()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var richText = (XLRichText)ws.Cell("A1").CreateRichText();
        richText
            .AddText("こんにち").SetBold(true) // Hello in hiragana
            .AddText("は,").SetBold(false) // object marker
            .AddText("世界").SetFontSize(15); // world in kanji
        richText.Phonetics
            .SetAlignment(XLPhoneticAlignment.Distributed)
            .Add("konnichi wa", 0, 6); // world in hiragana

        // Assert equal
        var immutableRichText = XLImmutableRichText.Create(richText);
        var equalImmutableRichText = XLImmutableRichText.Create(richText);
        await Assert.That(equalImmutableRichText).IsEqualTo(immutableRichText);

        // Different font of a first run
        richText.ElementAt(0).SetBold(false);
        var withDifferentTextRunFont = XLImmutableRichText.Create(richText);
        await Assert.That(withDifferentTextRunFont).IsNotEqualTo(immutableRichText);
        richText.ElementAt(0).SetBold(true);

        // Different phonetic properties
        richText.Phonetics.SetAlignment(XLPhoneticAlignment.Left);
        var withDifferentPhoneticsProps = XLImmutableRichText.Create(richText);
        await Assert.That(withDifferentPhoneticsProps).IsNotEqualTo(immutableRichText);
        richText.Phonetics.SetAlignment(XLPhoneticAlignment.Distributed);

        // Different phonetic runs
        richText.Phonetics.Add("せかい", 6, 8);
        var withDifferentTextPhonetics = XLImmutableRichText.Create(richText);
        await Assert.That(withDifferentTextPhonetics).IsNotEqualTo(immutableRichText);
    }
}
