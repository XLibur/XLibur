using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.ImageHandling;

public class PictureTests
{
    [Test]
    public async Task CanAddPictureFromStream()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        using var resourceStream = System.Reflection.Assembly.GetAssembly(typeof(XLibur.Examples.BasicTable)).GetManifestResourceStream("XLibur.Examples.Resources.SampleImage.jpg");
        var picture = ws.AddPicture(resourceStream, "MyPicture")
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .MoveTo(50, 50)
            .WithSize(200, 200);

        await Assert.That(picture.Format).IsEqualTo(XLPictureFormat.Jpeg);
        await Assert.That(picture.Width).IsEqualTo(200);
        await Assert.That(picture.Height).IsEqualTo(200);
    }

    [Test]
    public async Task CanAddPictureFromFile()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

        try
        {
            using (var resourceStream = System.Reflection.Assembly.GetAssembly(typeof(XLibur.Examples.BasicTable)).GetManifestResourceStream("XLibur.Examples.Resources.SampleImage.jpg"))
            using (var fileStream = File.Create(path))
            {
                resourceStream.Seek(0, SeekOrigin.Begin);
                resourceStream.CopyTo(fileStream);
                fileStream.Close();
            }

            var picture = ws.AddPicture(path)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(50, 50);

            await Assert.That(picture.Format).IsEqualTo(XLPictureFormat.Jpeg);
            await Assert.That(picture.Width).IsEqualTo(400);
            await Assert.That(picture.Height).IsEqualTo(400);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    [Test]
    public async Task CanAddPictureConcurrentlyFromFile()
    {
        var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

        try
        {
            using (var resourceStream = System.Reflection.Assembly.GetAssembly(typeof(XLibur.Examples.BasicTable)).GetManifestResourceStream("XLibur.Examples.Resources.SampleImage.jpg"))
            using (var fileStream = File.Create(path))
            {
                resourceStream.Seek(0, SeekOrigin.Begin);
                resourceStream.CopyTo(fileStream);
                fileStream.Close();
            }

            await Task.WhenAll(
                Task.Run(() => verifyAddImageFromFile(path)),
                Task.Run(() => verifyAddImageFromFile(path)));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    private static async Task verifyAddImageFromFile(string filePath)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var picture = ws.AddPicture(filePath)
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .MoveTo(50, 50);

        await Assert.That(picture.Format).IsEqualTo(XLPictureFormat.Jpeg);
        await Assert.That(picture.Width).IsEqualTo(400);
        await Assert.That(picture.Top).IsEqualTo(50);
    }

    [Test]
    public async Task CanScaleImage()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        using var resourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png");
        var pic = ws.AddPicture(resourceStream, "MyPicture")
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .MoveTo(50, 50);

        await Assert.That(pic.OriginalWidth).IsEqualTo(252);
        await Assert.That(pic.OriginalHeight).IsEqualTo(152);
        await Assert.That(pic.Width).IsEqualTo(252);
        await Assert.That(pic.Height).IsEqualTo(152);

        pic.ScaleHeight(0.7);
        pic.ScaleWidth(1.2);

        await Assert.That(pic.OriginalWidth).IsEqualTo(252);
        await Assert.That(pic.OriginalHeight).IsEqualTo(152);
        await Assert.That(pic.Width).IsEqualTo(302);
        await Assert.That(pic.Height).IsEqualTo(106);

        pic.ScaleHeight(0.7);
        pic.ScaleWidth(1.2);

        await Assert.That(pic.OriginalWidth).IsEqualTo(252);
        await Assert.That(pic.OriginalHeight).IsEqualTo(152);
        await Assert.That(pic.Width).IsEqualTo(362);
        await Assert.That(pic.Height).IsEqualTo(74);

        pic.ScaleHeight(0.8, true);
        pic.ScaleWidth(1.1, true);

        await Assert.That(pic.OriginalWidth).IsEqualTo(252);
        await Assert.That(pic.OriginalHeight).IsEqualTo(152);
        await Assert.That(pic.Width).IsEqualTo(277);
        await Assert.That(pic.Height).IsEqualTo(122);
    }

    [Test]
    public async Task TestDefaultPictureNames()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        using (var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png"))
        {
            ws.AddPicture(stream, XLPictureFormat.Png);
            stream.Position = 0;

            ws.AddPicture(stream, XLPictureFormat.Png);
            stream.Position = 0;

            ws.AddPicture(stream, XLPictureFormat.Png).Name = "Picture 4";
            stream.Position = 0;

            ws.AddPicture(stream, XLPictureFormat.Png);
            stream.Position = 0;
        }

        await Assert.That(ws.Pictures.Skip(0).First().Name).IsEqualTo("Picture 1");
        await Assert.That(ws.Pictures.Skip(1).First().Name).IsEqualTo("Picture 2");
        await Assert.That(ws.Pictures.Skip(2).First().Name).IsEqualTo("Picture 4");
        await Assert.That(ws.Pictures.Skip(3).First().Name).IsEqualTo("Picture 5");
    }

    [Test]
    public async Task TestDefaultIds()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        using (var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png"))
        {
            ws.AddPicture(stream, XLPictureFormat.Png);
            stream.Position = 0;

            ws.AddPicture(stream, XLPictureFormat.Png);
            stream.Position = 0;

            ws.AddPicture(stream, XLPictureFormat.Png).Name = "Picture 4";
            stream.Position = 0;

            ws.AddPicture(stream, XLPictureFormat.Png);
            stream.Position = 0;
        }

        await Assert.That(ws.Pictures.Skip(0).First().Id).IsEqualTo(1);
        await Assert.That(ws.Pictures.Skip(1).First().Id).IsEqualTo(2);
        await Assert.That(ws.Pictures.Skip(2).First().Id).IsEqualTo(3);
        await Assert.That(ws.Pictures.Skip(3).First().Id).IsEqualTo(4);
    }

    [Test]
    public async Task XLMarkerTests()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        XLMarker firstMarker = new XLMarker(ws.Cell(1, 10), new Point(100, 0));

        await Assert.That(firstMarker.ColumnNumber).IsEqualTo(10);
        await Assert.That(firstMarker.RowNumber).IsEqualTo(1);
        await Assert.That(firstMarker.Offset.X).IsEqualTo(100);
        await Assert.That(firstMarker.Offset.Y).IsEqualTo(0);
    }

    [Test]
    public async Task XLPictureTests()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        using var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png");
        var pic = ws.AddPicture(stream, XLPictureFormat.Png, "Image1")
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .MoveTo(220, 155);

        await Assert.That(pic.Placement).IsEqualTo(XLPicturePlacement.FreeFloating);
        await Assert.That(pic.Name).IsEqualTo("Image1");
        await Assert.That(pic.Format).IsEqualTo(XLPictureFormat.Png);
        await Assert.That(pic.OriginalWidth).IsEqualTo(252);
        await Assert.That(pic.OriginalHeight).IsEqualTo(152);
        await Assert.That(pic.Width).IsEqualTo(252);
        await Assert.That(pic.Height).IsEqualTo(152);
        await Assert.That(pic.Left).IsEqualTo(220);
        await Assert.That(pic.Top).IsEqualTo(155);
    }

    [Test]
    public async Task CanLoadFileWithImagesAndCopyImagesToNewSheet()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        await Assert.That(ws.Pictures.Count).IsEqualTo(2);

        var copy = ws.CopyTo("NewSheet");
        await Assert.That(copy.Pictures.Count).IsEqualTo(2);
    }

    [Test]
    public async Task CanDeletePictureOnlyOne()
    {
        using var ms = new MemoryStream();
        int originalCount;

        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheets.First();
            originalCount = ws.Pictures.Count;
            ws.Pictures.Delete(ws.Pictures.First());

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.Pictures.Count).IsEqualTo(originalCount - 1);
        }
    }

    [Test]
    public async Task CanDeletePictures()
    {
        using var ms = new MemoryStream();
        int originalCount;

        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheets.First();
            originalCount = ws.Pictures.Count;
            ws.Pictures.Delete(ws.Pictures.First());

            var pictureName = ws.Pictures.First().Name;
            ws.Pictures.Delete(pictureName);

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.Pictures.Count).IsEqualTo(originalCount - 2);
        }
    }

    [Test]
    public async Task PictureRenameTests()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheet("Images3");
        var picture = ws.Pictures.First();
        await Assert.That(picture.Name).IsEqualTo("Picture 1");

        picture.Name = "picture 1";
        picture.Name = "pICture 1";
        picture.Name = "Picture 1";

        picture = ws.Pictures.Last();
        picture.Name = "new name";

        await Assert.That(() => picture.Name = "Picture 1").Throws<ArgumentException>();
        await Assert.That(() => picture.Name = "picTURE 1").Throws<ArgumentException>();
    }

    [Test]
    public async Task HandleDuplicatePictureIdsAcrossWorksheets()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet1");
        var ws2 = wb.AddWorksheet("Sheet2");

        using var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png");
        (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2);
        (ws1 as XLWorksheet).AddPicture(stream, "Picture 2", 3);

        //Internal method - used for loading files
        var pic = (ws2 as XLWorksheet).AddPicture(stream, "Picture 1", 2)
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .MoveTo(220, 155) as XLPicture;

        var id = pic.Id;

        pic.Id = id;
        await Assert.That(pic.Id).IsEqualTo(id);

        pic.Id = 3;
        await Assert.That(pic.Id).IsEqualTo(3);

        pic.Id = id;

        var pic2 = (ws2 as XLWorksheet).AddPicture(stream, "Picture 2", 3)
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .MoveTo(440, 300) as XLPicture;
    }

    [Test]
    public async Task CopyImageSameWorksheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");

        IXLPicture original;
        using (var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png"))
        {
            original = (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(220, 155) as XLPicture;
        }

        var copy = original.Duplicate()
            .MoveTo(300, 200) as XLPicture;

        await Assert.That(ws1.Pictures.Count).IsEqualTo(2);
        await Assert.That(copy.Worksheet).IsEqualTo(ws1);
        await Assert.That(copy.Format).IsEqualTo(original.Format);
        await Assert.That(copy.Height).IsEqualTo(original.Height);
        await Assert.That(copy.Placement).IsEqualTo(original.Placement);
        await Assert.That(copy.TopLeftCell.ToString()).IsEqualTo(original.TopLeftCell.ToString());
        await Assert.That(copy.Width).IsEqualTo(original.Width);
        await Assert.That(copy.ImageStream.ToArray()).IsEquivalentTo(original.ImageStream.ToArray(), CollectionOrdering.Matching).Because("Image streams differ");

        await Assert.That(copy.Top).IsEqualTo(200);
        await Assert.That(copy.Left).IsEqualTo(300);
        await Assert.That(copy.Id).IsNotEqualTo(original.Id);
        await Assert.That(copy.Name).IsNotEqualTo(original.Name);
    }

    [Test]
    public async Task CopyImageDifferentWorksheets()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        IXLPicture original;
        using (var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png"))
        {
            original = (ws1 as XLWorksheet).AddPicture(stream, "Picture 1", 2)
                .WithPlacement(XLPicturePlacement.FreeFloating)
                .MoveTo(220, 155) as XLPicture;
        }
        var ws2 = wb.Worksheets.Add("Sheet2");

        var copy = original.CopyTo(ws2);

        await Assert.That(ws1.Pictures.Count).IsEqualTo(1);
        await Assert.That(ws2.Pictures.Count).IsEqualTo(1);

        await Assert.That(copy.Worksheet).IsEqualTo(ws2);

        await Assert.That(copy.Format).IsEqualTo(original.Format);
        await Assert.That(copy.Height).IsEqualTo(original.Height);
        await Assert.That(copy.Left).IsEqualTo(original.Left);
        await Assert.That(copy.Name).IsEqualTo(original.Name);
        await Assert.That(copy.Placement).IsEqualTo(original.Placement);
        await Assert.That(copy.Top).IsEqualTo(original.Top);
        await Assert.That(copy.TopLeftCell.ToString()).IsEqualTo(original.TopLeftCell.ToString());
        await Assert.That(copy.Width).IsEqualTo(original.Width);
        await Assert.That(copy.ImageStream.ToArray()).IsEquivalentTo(original.ImageStream.ToArray(), CollectionOrdering.Matching).Because("Image streams differ");

        await Assert.That(copy.Id).IsNotEqualTo(original.Id);
    }

    [Test]
    public async Task PictureShiftsWhenInsertingRows()
    {
        using var wb = new XLWorkbook();
        using var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png");
        var ws = wb.Worksheets.Add("ImageShift");
        var picture = ws.AddPicture(stream, XLPictureFormat.Png, "PngImage")
            .MoveTo(ws.Cell(5, 2))
            .WithPlacement(XLPicturePlacement.Move);

        ws.Row(2).InsertRowsBelow(20);

        await Assert.That(picture.TopLeftCell.Address.RowNumber).IsEqualTo(25);
    }

    [Test]
    public async Task PictureNotFound()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        await Assert.That(() => ws.Picture("dummy")).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => ws.Pictures.Delete("dummy")).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task CanCopyEmfPicture()
    {
        // #1621 - There are 2 Bmp Guids: ImageFormat.Bmp and ImageFormat.MemoryBmp
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Pictures\EmfPicture.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws1 = wb.Worksheets.First();
        var img1 = ws1.Pictures.First();

        var ws2 = wb.AddWorksheet();

        var img2 = img1.CopyTo(ws2);

        await Assert.That(img2.Format).IsEqualTo(XLPictureFormat.Emf);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        ms.Seek(0, SeekOrigin.Begin);

        using var wb2 = new XLWorkbook(ms);
        ws2 = wb2.Worksheet("Sheet2");
        img2 = ws2.Pictures.First();
        await Assert.That(img2.Format).IsEqualTo(XLPictureFormat.Emf);
    }

    [Test]
    [Arguments("Picture:With:Colons")]
    [Arguments("Picture/With/Slashes")]
    [Arguments(@"Picture\With\Backslashes")]
    [Arguments("Picture?With?Questions")]
    [Arguments("Picture*With*Stars")]
    [Arguments("Picture[With]Brackets")]
    public async Task Picture_name_can_contain_special_characters(string name)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        using var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("XLibur.Tests.Resource.Images.ImageHandling.png");
        var pic = ws.AddPicture(stream, XLPictureFormat.Png, "temp");
        pic.Name = name;

        await Assert.That(pic.Name).IsEqualTo(name);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        ms.Seek(0, SeekOrigin.Begin);
        using var wb2 = new XLWorkbook(ms);
        var ws2 = wb2.Worksheet("Sheet1");
        await Assert.That(ws2.Pictures.First().Name).IsEqualTo(name);
    }

    [Test]
    public async Task CanAddSvgPictureFromStream()
    {
        var svgContent = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24\" height=\"24\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M20 6 9 17l-5-5\"/></svg>";
        using var svgStream = new MemoryStream(Encoding.UTF8.GetBytes(svgContent));

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        var picture = ws.AddPicture(svgStream, "img.svg")
            .MoveTo(ws.Cell(1, 1))
            .WithSize(120, 120);

        await Assert.That(picture.Format).IsEqualTo(XLPictureFormat.Svg);
        await Assert.That(picture.Width).IsEqualTo(120);
        await Assert.That(picture.Height).IsEqualTo(120);
        await Assert.That(picture.OriginalWidth).IsEqualTo(24);
        await Assert.That(picture.OriginalHeight).IsEqualTo(24);
    }

    [Test]
    public async Task CanSaveAndLoadSvgPicture()
    {
        var svgContent = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24\" height=\"24\" viewBox=\"0 0 24 24\"><path d=\"M20 6 9 17l-5-5\"/></svg>";
        using var svgStream = new MemoryStream(Encoding.UTF8.GetBytes(svgContent));

        using var ms = new MemoryStream();

        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            ws.AddPicture(svgStream, "CheckIcon")
                .MoveTo(ws.Cell(1, 1))
                .WithSize(120, 120);

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            await Assert.That(ws.Pictures.Count).IsEqualTo(1);

            var pic = ws.Pictures.First();
            await Assert.That(pic.Format).IsEqualTo(XLPictureFormat.Svg);
            await Assert.That(pic.Name).IsEqualTo("CheckIcon");
            await Assert.That(pic.Width).IsEqualTo(120);
            await Assert.That(pic.Height).IsEqualTo(120);
        }
    }

    [Test]
    public async Task KeepOriginalDrawingShapesZOrder()
    {
        // File contains shapes and a picture in a mixed order.
        using var stream = TestHelper.GetStreamFromResource("Other.Pictures.ImageShapeZOrder-Input.xlsx");
        await TestHelper.CreateAndCompare(
            () => new XLWorkbook(stream),
            @"Other\Pictures\ImageShapeZOrder-Output.xlsx");
    }
}
