using System;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Excel.Cells;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A sheet with an in-cell image keeps its part, or its cells, on the first save of a loaded workbook
/// when neither the sheet nor the numbering of the workbook's images has changed (#702, #721).
/// </summary>
/// <remarks>
/// An image cell names its image through a <c>vm</c> index into the value metadata and holds the
/// index of its rich value. A save writes both again for every image in the workbook, numbered in
/// the order of the sheets and of the cells in them, before it writes any sheet. A change elsewhere
/// that renumbers the images changes the indices of this sheet's cells without anyone editing them,
/// and the sheet must then be written.
/// </remarks>
public class KeepCellImageSheetTests
{
    /// <summary>The part of sheet Data, the second sheet.</summary>
    private const string Sheet2 = "xl/worksheets/sheet2.xml";

    [Test]
    public async Task An_unchanged_sheet_with_a_cell_image_keeps_its_part()
    {
        using var source = Source();

        using var saved = source.LoadAndSave();

        await Assert.That(saved.IsKeptWhole(Sheet2)).IsTrue();
        await AssertValues(saved);
        await AssertImage(saved, "Before", "A1", 1);
        await AssertImage(saved, "Data", "C1", 2);
    }

    [Test]
    public async Task Editing_a_cell_beside_the_images_on_another_sheet_keeps_the_cells()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Before").Cell("D4").Value = 7);

        await Assert.That(saved.HasKeptCells(Sheet2)).IsTrue();
        await AssertImage(saved, "Data", "C1", 2);
    }

    [Test]
    public async Task An_image_added_on_an_earlier_sheet_makes_the_save_write_the_sheet()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb => SetImage(wb.Worksheet("Before").Cell("B1"), 3));

        await Assert.That(saved.IsWrittenFromModel(Sheet2)).IsTrue();
        await AssertValues(saved);
        await AssertImage(saved, "Data", "C1", 2);
        await AssertImage(saved, "Before", "B1", 3);
    }

    [Test]
    public async Task An_image_removed_from_an_earlier_sheet_makes_the_save_write_the_sheet()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Before").Cell("A1").RemoveCellImage());

        await Assert.That(saved.IsWrittenFromModel(Sheet2)).IsTrue();
        await AssertValues(saved);
        await AssertImage(saved, "Data", "C1", 2);
    }

    [Test]
    public async Task Deleting_an_earlier_sheet_with_an_image_makes_the_save_write_the_sheet()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Before").Delete());

        await Assert.That(saved.IsWrittenFromModel(Sheet2)).IsTrue();
        await AssertValues(saved);
        await AssertImage(saved, "Data", "C1", 2);
    }

    [Test]
    public async Task Replacing_the_image_makes_the_save_write_the_sheet()
    {
        using var source = Source();

        using var saved = source.LoadAndSave(wb => SetImage(wb.Worksheet("Data").Cell("C1"), 4));

        await Assert.That(saved.IsWrittenFromModel(Sheet2)).IsTrue();
        await AssertImage(saved, "Data", "C1", 4);
    }

    [Test]
    public async Task An_image_cell_holding_an_error_as_Excel_writes_it_makes_the_save_write_the_sheet()
    {
        // Excel writes an image cell as #VALUE!. The save gives the cell its rich value index as a
        // number, so a rewrite writes the cell otherwise than the file had it.
        using var source = Source();
        source.RewritePart(Sheet2, xml => xml.Replace("vm=\"2\"><x:v>1</x:v>", "t=\"e\" vm=\"2\"><x:v>#VALUE!</x:v>",
            StringComparison.Ordinal));
        await Assert.That(source.ReadPart(Sheet2)).Contains("#VALUE!");

        using var saved = source.LoadAndSave();

        await Assert.That(saved.IsWrittenFromModel(Sheet2)).IsTrue();
        await AssertImage(saved, "Data", "C1", 2);
    }

    /// <summary>
    /// Sheet Before holds image 1 in A1. Sheet Data holds "Header" in A1, 1.5 in B2 and image 2 in
    /// C1, so its cell has <c>vm="2"</c> and rich value 1. Data's part is marked by
    /// <see cref="KeptSheet.Mark"/>.
    /// </summary>
    private static MemoryStream Source()
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            SetImage(wb.AddWorksheet("Before").Cell("A1"), 1);

            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Header";
            ws.Cell("B2").Value = 1.5;
            SetImage(ws.Cell("C1"), 2);

            wb.SaveAs(stream);
        }

        return stream.Mark(Sheet2);
    }

    /// <summary>
    /// Gives <paramref name="cell"/> an image told apart from the others by its length: the test
    /// image followed by <paramref name="tag"/> zero bytes. The load does not read alt text back.
    /// </summary>
    private static void SetImage(IXLCell cell, int tag)
    {
        var png = XLCellImageTests.CreateTestPng();
        using var image = new MemoryStream(png.Length + tag);
        image.Write(png);
        image.Write(new byte[tag]);
        image.Position = 0;
        cell.SetCellImage(image, XLPictureFormat.Png);
    }

    private static async Task AssertValues(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A1").GetText()).IsEqualTo("Header");
        await Assert.That(ws.Cell("B2").GetDouble()).IsEqualTo(1.5);
    }

    private static async Task AssertImage(MemoryStream package, string sheet, string cell, int tag)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var image = ((XLCell)wb.Worksheet(sheet).Cell(cell)).CellImage;

        await Assert.That(image).IsNotNull();
        var length = wb.InCellImages.GetImage(image!.WorkbookImageIndex).Stream.Length;
        await Assert.That(length).IsEqualTo(XLCellImageTests.CreateTestPng().Length + tag);
    }
}
