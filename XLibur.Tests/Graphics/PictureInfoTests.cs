using System.Drawing;
using System.IO;
using System.Reflection;
using XLibur.Excel.Drawings;
using XLibur.Fonts.SixLabors.V1;
using XLibur.Graphics;
using System.Threading.Tasks;

namespace XLibur.Tests.Graphics;

public class PictureInfoTests
{
    [Test]
    public async Task CanReadPng()
    {
        await AssertRasterImage("SampleImagePng.png", XLPictureFormat.Png, new Size(252, 152), 96, 96);
    }

    [Test]
    [Arguments("SampleImageJfif.jpg", 176, 270, 96, 96)]
    [Arguments("jpeg-rgb.jpg", 200, 200, 0, 0)] // Adobe JPG, has APP14 marker right after SOI instead of APP0
    [Arguments("jpeg-icc-profile.jpg", 4, 4, 0, 0)] // JPEG with ICC profile (APP2) as first marker
    [Arguments("jpeg-xmp.jpg", 4, 4, 0, 0)] // JPEG with XMP metadata (APP1/XMP) as first marker
    [Arguments("jpeg-dqt-first.jpg", 4, 4, 0, 0)] // JPEG with DQT as first marker (no APP segment)
    public async Task CanReadJfif(string filename, int widthPx, int heightPx, int dpiX, int dpiY)
    {
        await AssertRasterImage($"Jpg.{filename}", XLPictureFormat.Jpeg, new Size(widthPx, heightPx), dpiX, dpiY);
    }

    [Test]
    public async Task CanReadExif()
    {
        await AssertRasterImage("SampleImageExif.jpg", XLPictureFormat.Jpeg, new Size(252, 152), 0, 0);
    }

    [Test]
    public async Task CanReadGif87Image()
    {
        await AssertRasterImage("SampleImageGif87a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);
    }

    [Test]
    public async Task CanReadGif89Image()
    {
        await AssertRasterImage("SampleImageGif89a.gif", XLPictureFormat.Gif, new Size(500, 200), 0, 0);
    }

    [Test]
    [Arguments("SampleImageBmpWin24bit.bmp")]
    [Arguments("SampleImageBmpWin8bit.bmp")]
    [Arguments("SampleImageBmpWin4bit.bmp")]
    [Arguments("SampleImageBmpWin24bit.bmp")]
    public async Task CanReadBmpImageV3AndFurther(string imageName)
    {
        await AssertRasterImage(imageName, XLPictureFormat.Bmp, new Size(167, 51), 80.645d, 80.645d);
    }

    [Test]
    public async Task CanReadBmpV1()
    {
        await AssertRasterImage("SampleImageBmpV1.bmp", XLPictureFormat.Bmp, new Size(150, 50), 0, 0);
    }

    [Test]
    public async Task CanReadOs2BitmapArray()
    {
        // OS/2 "BA" file: a 14-byte BITMAPARRAYHEADER wrapping a plain BM bitmap
        // with a 12-byte BITMAPCOREHEADER (V1) reporting a 120x80 image.
        using var stream = new MemoryStream(BuildBitmapArrayV1(width: 120, height: 80));
        var read = new BmpInfoReader().TryGetInfo(stream, out var info);

        await Assert.That(read).IsTrue();
        await Assert.That(info.Format).IsEqualTo(XLPictureFormat.Bmp);
        await Assert.That(info.SizePx).IsEqualTo(new Size(120, 80));
    }

    [Test]
    public async Task Os2IconArrayIsRejected()
    {
        // Same wrapper, but the inner entry is an OS/2 icon ("IC"), which Excel can't read.
        var bytes = BuildBitmapArrayV1(width: 120, height: 80);
        bytes[14] = (byte)'I';
        bytes[15] = (byte)'C';
        using var stream = new MemoryStream(bytes);

        await Assert.That(new BmpInfoReader().TryGetInfo(stream, out _)).IsFalse();
    }

    private static byte[] BuildBitmapArrayV1(ushort width, ushort height)
    {
        var file = new byte[40];
        file[0] = (byte)'B';
        file[1] = (byte)'A'; // BITMAPARRAYHEADER (offsets 0..13)
        file[14] = (byte)'B';
        file[15] = (byte)'M'; // inner BITMAPFILEHEADER (offsets 14..27)
        WriteU16LE(file, 28, 12); // BITMAPCOREHEADER size (low word), high word stays 0
        WriteU16LE(file, 32, width);
        WriteU16LE(file, 34, height);
        return file;
    }

    [Test]
    public async Task CanReadTiffWithBigEndianEncoding()
    {
        await AssertRasterImage("SampleImageTiffBigEndian.tiff", XLPictureFormat.Tiff, new Size(130, 45), 96, 96);
    }

    [Test]
    public async Task CanReadTiffWithLittleEndianEncoding()
    {
        await AssertRasterImage("SampleImageTiffLittleEndian.tiff", XLPictureFormat.Tiff, new Size(130, 45), 96, 96);
    }

    [Test]
    public async Task CanReadPcx()
    {
        await AssertRasterImage("SampleImagePcx.pcx", XLPictureFormat.Pcx, new Size(100, 50), 96, 96);
    }

    [Test]
    public async Task PcxWithValidWindowBoundsIsRead()
    {
        // Sanity check that a hand-built header with sensible bounds is accepted.
        using var stream = new MemoryStream(BuildPcxHeader(xMin: 0, yMin: 0, xMax: 99, yMax: 49));
        var read = new PcxInfoReader().TryGetInfo(stream, out var info);

        await Assert.That(read).IsTrue();
        await Assert.That(info.SizePx).IsEqualTo(new Size(100, 50));
    }

    [Test]
    [Arguments(99, 0, 0, 49, DisplayName = "PcxRejectsXMaxBelowXMin")]
    [Arguments(0, 49, 99, 0, DisplayName = "PcxRejectsYMaxBelowYMin")]
    public async Task PcxWithMalformedWindowBoundsIsRejected(int xMin, int yMin, int xMax, int yMax)
    {
        // Otherwise valid PCX signature, but Max < Min would yield a zero/negative size.
        using var stream = new MemoryStream(BuildPcxHeader(xMin, yMin, xMax, yMax));
        var read = new PcxInfoReader().TryGetInfo(stream, out _);

        await Assert.That(read).IsFalse();
    }

    private static byte[] BuildPcxHeader(int xMin, int yMin, int xMax, int yMax)
    {
        var header = new byte[16];
        header[0] = 0x0A; // Manufacturer
        header[1] = 5; // Version
        header[2] = 1; // Encoding (RLE)
        header[3] = 8; // BitsPerPixel
        WriteU16LE(header, 4, xMin);
        WriteU16LE(header, 6, yMin);
        WriteU16LE(header, 8, xMax);
        WriteU16LE(header, 10, yMax);
        WriteU16LE(header, 12, 96); // HDpi
        WriteU16LE(header, 14, 96); // VDpi
        return header;
    }

    private static void WriteU16LE(byte[] buffer, int offset, int value)
    {
        buffer[offset] = (byte)(value & 0xFF);
        buffer[offset + 1] = (byte)((value >> 8) & 0xFF);
    }

    [Test]
    public async Task CanReadWmfWithPlaceableHeader()
    {
        await AssertVectorImage("SampleImagePlaceableWmf.wmf", XLPictureFormat.Wmf, new Size(1000, 500));
    }

    [Test]
    public async Task CanReadWmfWithOriginalHeader()
    {
        await AssertVectorImage("SampleImageOriginalWmf.wmf", XLPictureFormat.Wmf, new Size(12496, 6247));
    }

    [Test]
    public async Task CanReadEmf()
    {
        await AssertVectorImage("SampleImageEmf.emf", XLPictureFormat.Emf, new Size(28844, 28938));
    }

    [Test]
    public async Task CanReadExtendedWebp()
    {
        await AssertRasterImage("SampleImageWebpExtendedFormat.webp", XLPictureFormat.Webp, new Size(188, 231), 72, 72);
    }

    [Test]
    public async Task CanReadLossyWebp()
    {
        await AssertRasterImage("SampleImageWebpLossy.webp", XLPictureFormat.Webp, new Size(278, 90), 72, 72);
    }

    [Test]
    public async Task CanReadLosslessWebp()
    {
        await AssertRasterImage("SampleImageWebpLossless.webp", XLPictureFormat.Webp, new Size(395, 136), 72, 72);
    }

    [Test]
    public async Task CanReadSvgWithWidthAndHeight()
    {
        await AssertRasterImage("SampleImageSvg.svg", XLPictureFormat.Svg, new Size(24, 24), 96, 96);
    }

    [Test]
    public async Task CanReadSvgWithViewBoxOnly()
    {
        await AssertRasterImage("SampleImageSvgViewBox.svg", XLPictureFormat.Svg, new Size(100, 50), 96, 96);
    }

    private static async Task AssertRasterImage(string imageName, XLPictureFormat expectedFormat, Size expectedPxSize, double expectedDpiX, double expectedDpiY)
    {
        await AssertImage(imageName, expectedFormat, expectedPxSize, Size.Empty, expectedDpiX, expectedDpiY);
    }

    private static async Task AssertVectorImage(string imageName, XLPictureFormat expectedFormat, Size expectedHiMetricSize)
    {
        await AssertImage(imageName, expectedFormat, Size.Empty, expectedHiMetricSize, 0, 0);
    }

    private static async Task AssertImage(string imageName, XLPictureFormat expectedFormat, Size expectedPxSize, Size expectedHiMetricSize, double expectedDpiX, double expectedDpiY)
    {
        using var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream($"XLibur.Tests.Resource.Images.{imageName}");
        var engine = new DefaultGraphicEngine(DefaultFontEngine.Instance.Value);
        var info = engine.GetPictureInfo(stream, XLPictureFormat.Unknown);

        await Assert.That(info.Format).IsEqualTo(expectedFormat);
        await Assert.That(info.SizePx).IsEqualTo(expectedPxSize);
        await Assert.That(info.SizePhys).IsEqualTo(expectedHiMetricSize);

        // Some DPI is stored as pixels per meter, causing a rounding errors.
        await Assert.That(info.DpiX).IsEqualTo(expectedDpiX).Within(0.02);
        await Assert.That(info.DpiY).IsEqualTo(expectedDpiY).Within(0.02);
    }
}
