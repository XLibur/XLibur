using System;
using System.IO;
using System.Linq;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Loading;

/// <summary>
/// Covers reading of cell attribute and <c>&lt;v&gt;</c> content that does not fit the fixed-size
/// scratch buffer the sheet-data reader uses to avoid a string allocation per cell. Well-formed
/// files never exceed it, so without these tests the fallback path would be unexercised.
/// </summary>
[TestFixture]
public class SheetDataValueReadingTests
{
    private static XLWorkbook SaveAndReload(XLWorkbook workbook)
    {
        var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        return new XLWorkbook(ms);
    }

    private static XLWorkbook SaveAndReload(XLWorkbook workbook, SaveOptions options)
    {
        var ms = new MemoryStream();
        workbook.SaveAs(ms, options);
        ms.Position = 0;
        return new XLWorkbook(ms);
    }

    [TestCase(10)]
    [TestCase(63)]
    [TestCase(64)]
    [TestCase(65)]
    [TestCase(200)]
    [TestCase(5000)]
    public void Text_of_any_length_round_trips(int length)
    {
        var text = string.Concat(Enumerable.Repeat("abcdefghij", (length / 10) + 1))[..length];

        using var original = new XLWorkbook();
        var ws = original.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = text;

        using var reloaded = SaveAndReload(original);

        Assert.That(reloaded.Worksheet("Sheet1").Cell("A1").GetString(), Is.EqualTo(text));
    }

    [Test]
    public void Formula_string_result_longer_than_the_scratch_buffer_round_trips()
    {
        // A cached formula result is written as t="str" with the text inline in <v>, so it goes
        // through the same buffered read as a normal value but is not a shared string.
        var text = new string('x', 300);

        using var original = new XLWorkbook();
        var ws = original.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = text;
        ws.Cell("A2").FormulaA1 = "A1";

        // Cached formula values are only written when explicitly requested; without this the
        // formula cell would carry no <v> and the t="str" read path would not be exercised.
        using var reloaded = SaveAndReload(original, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        Assert.That(reloaded.Worksheet("Sheet1").Cell("A2").CachedValue.GetText(), Is.EqualTo(text));
    }

    /// <remarks>
    /// Restricted to values of at most 15 significant digits. Saving formats doubles with
    /// <c>"G15"</c> by deliberate policy (see <c>ObjectExtensions.ToInvariantString</c>), so wider
    /// values — and the extremes of <see cref="double"/> — lose precision on the way out. That is a
    /// property of the write path; what is asserted here is that the read path reconstructs exactly
    /// what was written.
    /// </remarks>
    [Test]
    public void High_precision_numbers_round_trip()
    {
        double[] numbers =
        [
            0d, 1d, -1d, 0.1d, -0.1d, 0.3d,
            1234567890.12345d,
            -123456.789012345d,
            1e-300, 1e300,
            123456789012345d,
            0.000123456789012d
        ];

        using var original = new XLWorkbook();
        var ws = original.AddWorksheet("Sheet1");
        for (var i = 0; i < numbers.Length; i++)
            ws.Cell(i + 1, 1).Value = numbers[i];

        using var reloaded = SaveAndReload(original);

        var loaded = reloaded.Worksheet("Sheet1");
        for (var i = 0; i < numbers.Length; i++)
        {
            Assert.That(loaded.Cell(i + 1, 1).GetDouble(), Is.EqualTo(numbers[i]),
                $"number at row {i + 1} did not round-trip");
        }
    }

    [Test]
    public void Shared_string_indexes_beyond_the_scratch_buffer_width_round_trip()
    {
        // Forces six-digit shared string indexes, so the <v> content of later cells is longer
        // than a short value but still well inside the buffer — guards the index parse path.
        const int count = 200_000;

        using var original = new XLWorkbook();
        var ws = original.AddWorksheet("Sheet1");
        for (var i = 0; i < count; i++)
            ws.Cell(i + 1, 1).Value = "s" + i;

        using var reloaded = SaveAndReload(original);

        var loaded = reloaded.Worksheet("Sheet1");
        Assert.Multiple(() =>
        {
            Assert.That(loaded.Cell(1, 1).GetString(), Is.EqualTo("s0"));
            Assert.That(loaded.Cell(count / 2, 1).GetString(), Is.EqualTo("s" + (count / 2 - 1)));
            Assert.That(loaded.Cell(count, 1).GetString(), Is.EqualTo("s" + (count - 1)));
        });
    }

    [Test]
    public void Whitespace_only_and_padded_text_is_preserved()
    {
        // <t xml:space="preserve"> content must survive the shared-string reader verbatim.
        string[] texts = ["   ", " leading", "trailing ", " both ", "\ttab\t", new string(' ', 100)];

        using var original = new XLWorkbook();
        var ws = original.AddWorksheet("Sheet1");
        for (var i = 0; i < texts.Length; i++)
            ws.Cell(i + 1, 1).Value = texts[i];

        using var reloaded = SaveAndReload(original);

        var loaded = reloaded.Worksheet("Sheet1");
        for (var i = 0; i < texts.Length; i++)
        {
            Assert.That(loaded.Cell(i + 1, 1).GetString(), Is.EqualTo(texts[i]),
                $"text at row {i + 1} was not preserved");
        }
    }

    [Test]
    public void Escaped_control_characters_are_decoded_once()
    {
        // _xHHHH_ escapes are decoded by the shared-string reader; a literal underscore run
        // must not be mistaken for an escape.
        string[] texts = ["ab", "_x0018_literal", "_Xceed_Something", "plain_underscore"];

        using var original = new XLWorkbook();
        var ws = original.AddWorksheet("Sheet1");
        for (var i = 0; i < texts.Length; i++)
            ws.Cell(i + 1, 1).Value = texts[i];

        using var reloaded = SaveAndReload(original);

        var loaded = reloaded.Worksheet("Sheet1");
        for (var i = 0; i < texts.Length; i++)
        {
            Assert.That(loaded.Cell(i + 1, 1).GetString(), Is.EqualTo(texts[i]),
                $"text at row {i + 1} was not decoded correctly");
        }
    }

    [Test]
    public void Rich_text_and_plain_text_can_share_one_table()
    {
        using var original = new XLWorkbook();
        var ws = original.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "plain before";

        var rich = ws.Cell("A2").GetRichText();
        rich.AddText("bold").Bold = true;
        rich.AddText(new string('n', 120));

        ws.Cell("A3").Value = "plain after";

        using var reloaded = SaveAndReload(original);

        var loaded = reloaded.Worksheet("Sheet1");
        Assert.Multiple(() =>
        {
            Assert.That(loaded.Cell("A1").GetString(), Is.EqualTo("plain before"));
            Assert.That(loaded.Cell("A3").GetString(), Is.EqualTo("plain after"));
            Assert.That(loaded.Cell("A2").GetString(), Is.EqualTo("bold" + new string('n', 120)));

            var loadedRich = loaded.Cell("A2").GetRichText();
            Assert.That(loadedRich.Count, Is.EqualTo(2));
            Assert.That(loadedRich.First().Bold, Is.True);
        });
    }
}
