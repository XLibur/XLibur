using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Rows;

public class RowHeightRoundTripTests
{
    private static XElement SavedRow(Stream xlsx, int rowNumber)
    {
        xlsx.Position = 0;
        using var doc = SpreadsheetDocument.Open(xlsx, isEditable: false);
        var part = doc.WorkbookPart!.WorksheetParts.Single();
        using var stream = part.GetStream();
        var xml = XDocument.Load(stream);
        var ns = xml.Root!.Name.Namespace;
        return xml.Root.Element(ns + "sheetData")!
            .Elements(ns + "row")
            .Single(r => r.Attribute("r")?.Value == rowNumber.ToString());
    }

    private static MemoryStream SaveEditedFixture(System.Action<IXLWorksheet> edit)
    {
        using var fixture = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EscapedApostrophe.xlsx"));
        var saved = new MemoryStream();
        using var wb = new XLWorkbook(fixture);
        edit(wb.Worksheet(1));
        wb.SaveAs(saved);
        return saved;
    }

    [Test]
    public async Task Auto_fitted_row_height_stays_auto_after_an_edit()
    {
        using var saved = SaveEditedFixture(ws => ws.Cell("K200").Value = "edit");

        var autoRow = SavedRow(saved, 121);
        await Assert.That(autoRow.Attribute("ht")?.Value).IsEqualTo("16.5");
        await Assert.That(autoRow.Attribute("customHeight")).IsNull();

        var customRow = SavedRow(saved, 1);
        await Assert.That(customRow.Attribute("ht")?.Value).IsEqualTo("48");
        await Assert.That(customRow.Attribute("customHeight")?.Value).IsEqualTo("1");
    }

    [Test]
    public async Task Auto_fitted_row_keeps_its_cached_height_in_the_model()
    {
        using var fixture = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EscapedApostrophe.xlsx"));
        using var wb = new XLWorkbook(fixture);

        await Assert.That(wb.Worksheet(1).Row(121).Height).IsEqualTo(16.5);
    }

    [Test]
    public async Task Setting_the_height_of_an_auto_fitted_row_pins_it()
    {
        using var saved = SaveEditedFixture(ws => ws.Row(121).Height = 16.5);

        var row = SavedRow(saved, 121);
        await Assert.That(row.Attribute("ht")?.Value).IsEqualTo("16.5");
        await Assert.That(row.Attribute("customHeight")?.Value).IsEqualTo("1");
    }

    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task Inserted_row_is_saved_with_the_height_of_the_row_it_copies(bool below)
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Row(1).Height = 30;
            ws.Row(2).Height = 40;
            if (below)
                ws.Row(1).InsertRowsBelow(1);
            else
                ws.Row(2).InsertRowsAbove(1);
            wb.SaveAs(saved);
        }

        var inserted = SavedRow(saved, 2);
        await Assert.That(inserted.Attribute("ht")?.Value).IsEqualTo("30");
        await Assert.That(inserted.Attribute("customHeight")?.Value).IsEqualTo("1");
    }

    [Test]
    public async Task Row_inserted_below_an_auto_fitted_row_keeps_the_height_auto()
    {
        using var saved = SaveEditedFixture(ws => ws.Row(121).InsertRowsBelow(1));

        var inserted = SavedRow(saved, 122);
        await Assert.That(inserted.Attribute("ht")?.Value).IsEqualTo("16.5");
        await Assert.That(inserted.Attribute("customHeight")).IsNull();
    }
}
