using System;
using System.Text;
using XLibur.Excel;
using System.Threading.Tasks;

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
