using System.IO;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Cells;

public class XLCellImageTests
{
    /// <summary>
    /// Create a small valid PNG byte array (1×1 pixel, red).
    /// </summary>
    private static byte[] CreateTestPng()
    {
        // Minimal valid 1×1 red PNG
        return new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, // 8-bit RGB
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, // compressed data
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, // ...
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND chunk
            0x44, 0xAE, 0x42, 0x60, 0x82,
        };
    }

    [Test]
    public async Task SetCellImage_StoresImageInWorkbook()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");

        using var imgStream = new MemoryStream(CreateTestPng());
        cell.SetCellImage(imgStream, XLPictureFormat.Png);

        await Assert.That(wb.InCellImages.Count).IsEqualTo(1);
        await Assert.That(cell.HasCellImage).IsTrue();
    }

    [Test]
    public async Task SetCellImage_DeduplicatesSameImage()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var pngBytes = CreateTestPng();

        using (var s1 = new MemoryStream(pngBytes))
            ws.Cell("A1").SetCellImage(s1, XLPictureFormat.Png);

        using (var s2 = new MemoryStream(pngBytes))
            ws.Cell("B1").SetCellImage(s2, XLPictureFormat.Png);

        await Assert.That(wb.InCellImages.Count).IsEqualTo(1);
    }

    [Test]
    public async Task HasCellImage_ReturnsTrueAfterSet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");

        await Assert.That(cell.HasCellImage).IsFalse();

        using var imgStream = new MemoryStream(CreateTestPng());
        cell.SetCellImage(imgStream, XLPictureFormat.Png);

        await Assert.That(cell.HasCellImage).IsTrue();
    }

    [Test]
    public async Task RemoveCellImage_ClearsImage()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");

        using var imgStream = new MemoryStream(CreateTestPng());
        cell.SetCellImage(imgStream, XLPictureFormat.Png);
        cell.RemoveCellImage();

        await Assert.That(cell.HasCellImage).IsFalse();
    }

    [Test]
    public async Task Clear_Contents_RemovesCellImage()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");

        using var imgStream = new MemoryStream(CreateTestPng());
        cell.SetCellImage(imgStream, XLPictureFormat.Png);
        cell.Clear(XLClearOptions.Contents);

        await Assert.That(cell.HasCellImage).IsFalse();
    }

    [Test]
    public async Task CopyCell_CopiesCellImage()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var source = ws.Cell("A1");

        using var imgStream = new MemoryStream(CreateTestPng());
        source.SetCellImage(imgStream, XLPictureFormat.Png, "test alt");

        var target = ws.Cell("B1");
        target.CopyFrom(source);

        await Assert.That(target.HasCellImage).IsTrue();
    }

    [Test]
    public async Task IsEmpty_FalseWhenHasImage()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");

        await Assert.That(cell.IsEmpty()).IsTrue();

        using var imgStream = new MemoryStream(CreateTestPng());
        cell.SetCellImage(imgStream, XLPictureFormat.Png);

        await Assert.That(cell.IsEmpty()).IsFalse();
    }

    [Test]
    public async Task SaveAndReload_PreservesCellImage()
    {
        await TestHelper.CreateSaveLoadAssert(
            (wb, ws) =>
            {
                using var imgStream = new MemoryStream(CreateTestPng());
                ws.Cell("A1").SetCellImage(imgStream, XLPictureFormat.Png, "red pixel");
            },
            async (wb, ws) =>
            {
                var cell = ws.Cell("A1");
                await Assert.That(cell.HasCellImage).IsTrue();
                await Assert.That(wb.InCellImages.Count).IsEqualTo(1);
            },
            validate: false);
    }

    [Test]
    public async Task SaveAndReload_MultipleCellsSameImage()
    {
        var pngBytes = CreateTestPng();

        await TestHelper.CreateSaveLoadAssert(
            (wb, ws) =>
            {
                using (var s1 = new MemoryStream(pngBytes))
                    ws.Cell("A1").SetCellImage(s1, XLPictureFormat.Png, "img1");
                using (var s2 = new MemoryStream(pngBytes))
                    ws.Cell("B1").SetCellImage(s2, XLPictureFormat.Png, "img2");
            },
            async (wb, ws) =>
            {
                await Assert.That(ws.Cell("A1").HasCellImage).IsTrue();
                await Assert.That(ws.Cell("B1").HasCellImage).IsTrue();
            },
            validate: false);
    }

    [Test]
    public async Task SaveAndReload_CellStyleSurvives()
    {
        await TestHelper.CreateSaveLoadAssert(
            (wb, ws) =>
            {
                var cell = ws.Cell("A1");
                cell.Style.Font.Bold = true;
                using var imgStream = new MemoryStream(CreateTestPng());
                cell.SetCellImage(imgStream, XLPictureFormat.Png);
            },
            async (wb, ws) =>
            {
                var cell = ws.Cell("A1");
                await Assert.That(cell.HasCellImage).IsTrue();
                await Assert.That(cell.Style.Font.Bold).IsTrue();
            },
            validate: false);
    }
}
