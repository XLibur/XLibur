using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Cells;

public class SharedStringTableTests
{
    [Test]
    public async Task SameStringIsNotStoredTwice()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();
        const string txt1 = "Hello";
        var txt2 = new StringBuilder("Hel").Append("lo").ToString();
        await Assert.That(txt2).IsNotSameReferenceAs(txt1);

        ws1.Cell(1, 1).Value = txt1;
        ws2.Cell(1, 1).Value = txt2;

        await Assert.That(ws2.Cell(1, 1).Value.GetText()).IsSameReferenceAs(ws1.Cell(1, 1).Value.GetText());
    }

    [Test]
    public async Task CanAccessTextThroughId()
    {
        var sst = new SharedStringTable();
        var id = sst.IncreaseRef("test", false);
        await Assert.That(sst[id]).IsEqualTo("test");
        await Assert.That(sst.Count).IsEqualTo(1);
    }

    [Test]
    public async Task TextsWithoutReferenceAreRemoved()
    {
        var sst = new SharedStringTable();
        var id = sst.IncreaseRef("test", false);
        sst.DecreaseRef(id);

        await Assert.That(sst.Count).IsEqualTo(0);
        var ex = await Assert.That(() => _ = sst[id]).Throws<ArgumentException>();
        await Assert.That(ex!.Message).IsEqualTo("Id 0 has no text.");
    }

    [Test]
    public async Task TextReferencedByMultipleThingsIsNotFreedUntilAllAreRelease()
    {
        const string text = "test";
        var sst = new SharedStringTable();
        var id = sst.IncreaseRef(text, false);

        sst.IncreaseRef(text, false);
        await Assert.That(sst[id]).IsEqualTo(text);
        await Assert.That(sst.Count).IsEqualTo(1);

        sst.DecreaseRef(id);
        await Assert.That(sst[id]).IsEqualTo(text);
        await Assert.That(sst.Count).IsEqualTo(1);

        sst.IncreaseRef(text, false);
        await Assert.That(sst[id]).IsEqualTo(text);
        await Assert.That(sst.Count).IsEqualTo(1);

        sst.DecreaseRef(id);
        await Assert.That(sst[id]).IsEqualTo(text);
        await Assert.That(sst.Count).IsEqualTo(1);

        sst.DecreaseRef(id);
        await Assert.That(() => _ = sst[id]).Throws<ArgumentException>();
    }

    [Test]
    public async Task FreedIdCanBeReusedForDifferentText()
    {
        var sst = new SharedStringTable();
        sst.IncreaseRef("zero", false);
        var originalId = sst.IncreaseRef("original", false);
        var laterId = sst.IncreaseRef("two", false);

        await Assert.That(laterId).IsGreaterThan(originalId);

        sst.DecreaseRef(originalId);
        await Assert.That(() => _ = sst[originalId]).Throws<ArgumentException>();

        var replacementId = sst.IncreaseRef("replacement", false);
        await Assert.That(replacementId).IsEqualTo(originalId);
        await Assert.That(sst[replacementId]).IsEqualTo("replacement");
    }

    [Test]
    public async Task Texts_loaded_from_a_file_are_mapped_in_the_file_order_before_new_texts()
    {
        var sst = new SharedStringTable();
        var b = sst.IncreaseRef("b", false);
        var a = sst.IncreaseRef("a", false);
        sst.RecordFileIndex(b, 1);
        sst.RecordFileIndex(a, 0);
        var added = sst.IncreaseRef("added", false);

        var map = sst.GetConsecutiveMap();

        await Assert.That(map[a]).IsEqualTo(0);
        await Assert.That(map[b]).IsEqualTo(1);
        await Assert.That(map[added]).IsEqualTo(2);
    }

    [Test]
    public async Task A_freed_text_forgets_its_file_index()
    {
        var sst = new SharedStringTable();
        var a = sst.IncreaseRef("a", false);
        var b = sst.IncreaseRef("b", false);
        sst.RecordFileIndex(a, 0);
        sst.RecordFileIndex(b, 1);

        sst.DecreaseRef(a);
        var replacement = sst.IncreaseRef("replacement", false);
        await Assert.That(replacement).IsEqualTo(a);

        var map = sst.GetConsecutiveMap();

        await Assert.That(map[b]).IsEqualTo(0);
        await Assert.That(map[replacement]).IsEqualTo(1);
    }

    [Test]
    public async Task A_text_the_file_lists_twice_keeps_its_first_index()
    {
        var sst = new SharedStringTable();
        var later = sst.IncreaseRef("later", false);
        var twice = sst.IncreaseRef("twice", false);
        sst.RecordFileIndex(later, 2);
        sst.RecordFileIndex(twice, 3);
        sst.RecordFileIndex(twice, 1);

        var map = sst.GetConsecutiveMap();

        await Assert.That(map[twice]).IsEqualTo(0);
        await Assert.That(map[later]).IsEqualTo(1);
    }

    [Test]
    public async Task Inline_and_unused_texts_are_not_mapped_whatever_their_file_index()
    {
        var sst = new SharedStringTable();
        var inline = sst.IncreaseRef("inline", true);
        var freed = sst.IncreaseRef("freed", false);
        var kept = sst.IncreaseRef("kept", false);
        sst.RecordFileIndex(kept, 5);
        sst.DecreaseRef(freed);

        var map = sst.GetConsecutiveMap();

        await Assert.That(map[inline]).IsEqualTo(-1);
        await Assert.That(map[freed]).IsEqualTo(-1);
        await Assert.That(map[kept]).IsEqualTo(0);
    }

    [Test]
    public async Task A_loaded_workbook_saves_its_shared_strings_at_the_indices_the_file_had()
    {
        // XLibur would write the strings in the order the cells first use them: plain, rich, other.
        // The file lists them in another order, as Excel's table often does.
        using var package = SaveWorkbookWithThreeStrings();
        using var reordered = ReorderSharedStrings(package, newOrder: [2, 0, 1]);
        var fileStrings = SharedStringItems(reordered);
        var fileCells = StringCellIndices(reordered);

        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(reordered))
            wb.SaveAs(saved);

        await Assert.That(SharedStringItems(saved)).IsEquivalentTo(fileStrings, CollectionOrdering.Matching);
        await Assert.That(StringCellIndices(saved)).IsEquivalentTo(fileCells, CollectionOrdering.Matching);
    }

    [Test]
    public async Task A_text_added_after_the_load_is_written_after_the_texts_of_the_file()
    {
        using var package = SaveWorkbookWithThreeStrings();
        using var reordered = ReorderSharedStrings(package, newOrder: [2, 0, 1]);
        var fileStrings = SharedStringItems(reordered);

        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(reordered))
        {
            wb.Worksheet(1).Cell("A5").Value = "added";
            wb.SaveAs(saved);
        }

        var savedStrings = SharedStringItems(saved);
        await Assert.That(savedStrings.Length).IsEqualTo(fileStrings.Length + 1);
        await Assert.That(savedStrings[..fileStrings.Length]).IsEquivalentTo(fileStrings, CollectionOrdering.Matching);
        await Assert.That(savedStrings[^1]).IsEqualTo("<x:si><x:t>added</x:t></x:si>");

        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet(1).Cell("A1").GetText()).IsEqualTo("plain");
        await Assert.That(reloaded.Worksheet(1).Cell("A2").GetRichText().Text).IsEqualTo("rich text");
        await Assert.That(reloaded.Worksheet(1).Cell("A3").GetText()).IsEqualTo("other");
        await Assert.That(reloaded.Worksheet(1).Cell("A5").GetText()).IsEqualTo("added");
    }

    private static MemoryStream SaveWorkbookWithThreeStrings()
    {
        var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "plain";
            var rich = ws.Cell("A2").CreateRichText();
            rich.AddText("rich ").SetBold();
            rich.AddText("text");
            ws.Cell("A3").Value = "other";
            ws.Cell("A4").Value = "plain";
            wb.SaveAs(output);
        }

        output.Position = 0;
        return output;
    }

    /// <summary>
    /// A copy of <paramref name="package"/> whose shared-string item <c>i</c> is the original item
    /// <c>newOrder[i]</c>, with every shared-string cell renumbered to match.
    /// </summary>
    private static MemoryStream ReorderSharedStrings(MemoryStream package, int[] newOrder)
    {
        var items = SharedStringItems(package);
        if (items.Length != newOrder.Length)
            throw new InvalidOperationException($"The package has {items.Length} shared strings, not {newOrder.Length}.");

        var copy = new MemoryStream();
        copy.Write(package.GetBuffer(), 0, (int)package.Length);

        copy.RewritePart(SharedStringsPart, xml =>
        {
            var start = xml.IndexOf("<x:si>", StringComparison.Ordinal);
            var end = xml.LastIndexOf("</x:si>", StringComparison.Ordinal) + "</x:si>".Length;
            return xml[..start] + string.Concat(newOrder.Select(i => items[i])) + xml[end..];
        });

        var newIndexOf = new int[newOrder.Length];
        for (var i = 0; i < newOrder.Length; i++)
            newIndexOf[newOrder[i]] = i;

        copy.RewriteSheet1(xml => StringCellValue.Replace(xml,
            m => m.Groups["open"].Value + newIndexOf[int.Parse(m.Groups["index"].Value)] + m.Groups["close"].Value));

        copy.Position = 0;
        return copy;
    }

    private const string SharedStringsPart = "xl/sharedStrings.xml";

    private static readonly Regex SharedStringItem = new("<x:si>.*?</x:si>", RegexOptions.Singleline);

    private static readonly Regex StringCellValue =
        new("(?<open><x:c r=\"(?<ref>[A-Z]+[0-9]+)\"[^>]*t=\"s\"[^>]*><x:v>)(?<index>[0-9]+)(?<close></x:v>)");

    private static string[] SharedStringItems(Stream package) =>
        SharedStringItem.Matches(package.ReadPart(SharedStringsPart)).Select(m => m.Value).ToArray();

    private static string[] StringCellIndices(Stream package) =>
        StringCellValue.Matches(package.Sheet1Xml())
            .Select(m => m.Groups["ref"].Value + "=" + m.Groups["index"].Value)
            .ToArray();

    [Test]
    public async Task DereferencingFreedIdThrows()
    {
        var sst = new SharedStringTable();
        var id = sst.IncreaseRef("test", false);
        sst.DecreaseRef(id);
        await Assert.That(() => sst.DecreaseRef(id)).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task StringItem_without_text_is_loaded_as_empty_text()
    {
        // PR#2218: A text cell that references self-closed <si/> tag in SST is loaded without
        // an error and is loaded as type TEXT. Although it's not very common, an empty string is
        // a valid value of a cell.
        await TestHelper.LoadAndAssert(async (_, ws) =>
        {
            // Check that type is an empty string, just like in Excel.
            await Assert.That(ws.Evaluate("TYPE(B2)")).IsEqualTo(2);
            await Assert.That(ws.Cell("B2").GetText()).IsEmpty();
        }, @"Other\Cells\EmptySi.xlsx");
    }

    [Test]
    public async Task Empty_text_is_written_and_loaded_to_sst()
    {
        await TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.Cell("A1").Value = "Empty text cell (B1):";
                ws.Cell("B1").Value = string.Empty;

                ws.Cell("A2").Value = "Empty rich text";
                ws.Cell("B2").CreateRichText().AddText(string.Empty);
            },
            async (_, ws) =>
            {
                await Assert.That(ws.Cell("B1").CachedValue).IsEqualTo("");
                await Assert.That(ws.Cell("B2").GetRichText().Text).IsEqualTo("");
            },
            @"Other\Cells\EmptyText.xlsx");
    }
}
