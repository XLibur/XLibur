using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Excel.Drawings;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Drawings;

/// <summary>
/// An <c>Add</c> that throws because the name is taken leaves the pictures as they were (issue #806).
/// </summary>
public class PictureNameClashTests
{
    [Test]
    public async Task Pictures_Add_stream_with_a_taken_name_adds_no_picture()
    {
        await AssertRejectedNameAddsNoPicture(ws => ws.Pictures.Add(new MemoryStream(Png), "logo"));
    }

    [Test]
    public async Task Pictures_Add_stream_and_format_with_a_taken_name_adds_no_picture()
    {
        await AssertRejectedNameAddsNoPicture(ws => ws.Pictures.Add(new MemoryStream(Png), XLPictureFormat.Png, "LOGO"));
    }

    [Test]
    public async Task Pictures_Add_file_with_a_taken_name_adds_no_picture()
    {
        var file = PngFile();
        try
        {
            await AssertRejectedNameAddsNoPicture(ws => ws.Pictures.Add(file, "Logo"));
        }
        finally
        {
            File.Delete(file);
        }
    }

    [Test]
    public async Task AddPicture_stream_with_a_taken_name_adds_no_picture()
    {
        await AssertRejectedNameAddsNoPicture(ws => ws.AddPicture(new MemoryStream(Png), "logo"));
    }

    [Test]
    public async Task AddPicture_stream_and_format_with_a_taken_name_adds_no_picture()
    {
        await AssertRejectedNameAddsNoPicture(ws => ws.AddPicture(new MemoryStream(Png), XLPictureFormat.Png, "LOGO"));
    }

    [Test]
    public async Task AddPicture_file_with_a_taken_name_adds_no_picture()
    {
        var file = PngFile();
        try
        {
            await AssertRejectedNameAddsNoPicture(ws => ws.AddPicture(file, "Logo"));
        }
        finally
        {
            File.Delete(file);
        }
    }

    [Test]
    public async Task A_whitespace_name_adds_no_picture()
    {
        await AssertRejectedNameAddsNoPicture(ws => ws.Pictures.Add(new MemoryStream(Png), "   "));
        await AssertRejectedNameAddsNoPicture(ws => ws.Pictures.Add(new MemoryStream(Png), XLPictureFormat.Png, " "));
        await AssertRejectedNameAddsNoPicture(ws => ws.AddPicture(new MemoryStream(Png), "\t"));
    }

    [Test]
    public async Task The_next_picture_after_a_failed_add_takes_the_name_it_would_have_had()
    {
        // The failed add takes no generated name either. Beside "Logo" alone, a new picture is
        // "Picture 2"; a picture the failed add left behind would have taken that name.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.AddPicture(new MemoryStream(Png), "Logo");
        await Assert.That(() => ws.AddPicture(new MemoryStream(Png), "logo")).Throws<ArgumentException>();

        await Assert.That(ws.AddPicture(new MemoryStream(Png)).Name).IsEqualTo("Picture 2");
    }

    /// <summary>
    /// Beside a picture named "Logo", <paramref name="addSecond"/> throws, because the name it gives
    /// is taken or blank, and leaves "Logo" the sheet's only picture.
    /// </summary>
    private static async Task AssertRejectedNameAddsNoPicture(Func<IXLWorksheet, IXLPicture> addSecond)
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.AddPicture(new MemoryStream(Png), "Logo");

            await Assert.That(() => addSecond(ws)).Throws<ArgumentException>();

            await Assert.That(ws.Pictures.Count).IsEqualTo(1);
            await Assert.That(ws.Pictures.Single().Name).IsEqualTo("Logo");
            wb.SaveAs(saved);
        }

        await Assert.That(SavedPictureNames(saved)).IsEqualTo("Logo");
    }

    private static readonly XNamespace XdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    /// <summary>The names of the pictures in the saved drawing, in order, comma separated.</summary>
    private static string SavedPictureNames(Stream saved)
    {
        var names = XDocument.Parse(saved.ReadPart("xl/drawings/drawing1.xml")).Descendants(XdrNs + "pic")
            .Select(picture => (string?)picture.Descendants(XdrNs + "cNvPr").Single().Attribute("name"));
        return string.Join(", ", names);
    }

    /// <summary>Writes <see cref="Png"/> to a new temporary file and returns its path.</summary>
    private static string PngFile()
    {
        var file = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.png");
        File.WriteAllBytes(file, Png);
        return file;
    }

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");
}
