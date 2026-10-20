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
/// The ids and names a sheet's pictures take, as they are added, renamed, renumbered and deleted.
/// </summary>
/// <remarks>
/// Issue #804 keeps the highest id and the names taken as the pictures change, rather than reading
/// every picture each time a picture is added, so a load of n pictures no longer costs n². These pin
/// what the collection says while it changes under them.
/// </remarks>
public class PictureIdAndNameTests
{
    private const string Drawing1 = "xl/drawings/drawing1.xml";

    private static readonly XNamespace XdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    [Test]
    public async Task A_deleted_picture_name_is_free_again_in_any_case()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.AddPicture(new MemoryStream(Png), "Logo").Delete();

        await Assert.That(ws.AddPicture(new MemoryStream(Png), "LOGO").Name).IsEqualTo("LOGO");
    }

    [Test]
    public async Task A_renamed_picture_frees_its_old_name_and_takes_its_new_one_in_any_case()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.AddPicture(new MemoryStream(Png), "Before").Name = "After";
        var other = ws.AddPicture(new MemoryStream(Png), "Other");

        await Assert.That(() => ws.AddPicture(new MemoryStream(Png), "after")).Throws<ArgumentException>();
        await Assert.That(() => other.Name = "AFTER").Throws<ArgumentException>();
        await Assert.That(ws.AddPicture(new MemoryStream(Png), "before").Name).IsEqualTo("before");
    }

    [Test]
    public async Task A_generated_name_skips_the_names_taken_and_reuses_a_freed_one()
    {
        // A generated name starts from the number of pictures. Picture 3, deleted, is free for the
        // next picture; with Picture 2 renamed "picture 4", the picture after that starts at
        // "Picture 4", which is taken, and takes "Picture 5".
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.AddPicture(new MemoryStream(Png));
        var second = ws.AddPicture(new MemoryStream(Png));
        ws.AddPicture(new MemoryStream(Png)).Delete();

        await Assert.That(ws.AddPicture(new MemoryStream(Png)).Name).IsEqualTo("Picture 3");

        second.Name = "picture 4";
        await Assert.That(ws.AddPicture(new MemoryStream(Png)).Name).IsEqualTo("Picture 5");
    }

    [Test]
    public async Task A_loaded_picture_name_is_taken_in_any_case_until_the_picture_is_renamed()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet().AddPicture(new MemoryStream(Png), "Logo");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var loaded = new XLWorkbook(saved);
        var ws = loaded.Worksheet(1);
        await Assert.That(() => ws.AddPicture(new MemoryStream(Png), "LOGO")).Throws<ArgumentException>();

        ws.Picture("Logo").Name = "Renamed";
        await Assert.That(ws.AddPicture(new MemoryStream(Png), "logo").Name).IsEqualTo("logo");
    }

    [Test]
    public async Task A_new_picture_takes_the_id_after_the_highest_any_sheet_holds_now()
    {
        // Sheet 1 holds 1 to 3, and Sheet 2's first picture takes 4. With 3 and 4 deleted the
        // highest is 2, so the next picture takes 3.
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();
        ws1.AddPicture(new MemoryStream(Png));
        ws1.AddPicture(new MemoryStream(Png));
        var third = ws1.AddPicture(new MemoryStream(Png));
        var fourth = ws2.AddPicture(new MemoryStream(Png));
        await Assert.That(fourth.Id).IsEqualTo(4);

        fourth.Delete();
        third.Delete();

        await Assert.That(ws2.AddPicture(new MemoryStream(Png)).Id).IsEqualTo(3);
    }

    [Test]
    public async Task A_new_picture_after_a_save_takes_the_id_after_the_highest_the_save_wrote()
    {
        // Pictures 2 and 3 are left of three, and the save renumbers them 1 and 2, so the next picture
        // takes 3, not 4.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var first = ws.AddPicture(new MemoryStream(Png));
        ws.AddPicture(new MemoryStream(Png));
        ws.AddPicture(new MemoryStream(Png));
        first.Delete();
        await Assert.That(string.Join(", ", ws.Pictures.Select(p => p.Id))).IsEqualTo("2, 3");

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(string.Join(", ", ws.Pictures.Select(p => p.Id))).IsEqualTo("1, 2");
        await Assert.That(ws.AddPicture(new MemoryStream(Png)).Id).IsEqualTo(3);
    }

    [Test]
    public async Task A_picture_takes_an_id_another_picture_on_the_sheet_holds_only_once_that_one_is_gone()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        var first = (XLPicture)ws.AddPicture(new MemoryStream(Png));
        var second = (XLPicture)ws.AddPicture(new MemoryStream(Png));

        await Assert.That(() => second.Id = 1).Throws<ArgumentException>();

        first.Delete();
        second.Id = 1;

        // The highest id is now 1, so the next picture takes 2.
        await Assert.That(second.Id).IsEqualTo(1);
        await Assert.That(ws.AddPicture(new MemoryStream(Png)).Id).IsEqualTo(2);
    }

    [Test]
    public async Task A_load_reads_each_sheets_picture_ids_for_the_highest_at_most_once()
    {
        // Each loaded picture first takes the next free id, then the file's. Before the fix, giving
        // up the next free id, the highest, lost the highest, and each picture after it read every
        // id again to find it: a load of n pictures read n² ids.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            foreach (var sheet in new[] { wb.AddWorksheet(), wb.AddWorksheet() })
            {
                for (var i = 0; i < 10; i++)
                    sheet.AddPicture(new MemoryStream(Png)).MoveTo(sheet.Cell(i * 5 + 1, 1));
            }

            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var loaded = new XLWorkbook(saved);
        foreach (var sheet in loaded.Worksheets)
        {
            var pictures = (XLPictures)sheet.Pictures;
            await Assert.That(pictures.Count).IsEqualTo(10);
            await Assert.That(pictures.HighestIdReads).IsLessThanOrEqualTo(1);
        }

        // The highest is still right: the save numbers each drawing's pictures from 1, so each
        // sheet holds 1 to 10.
        await Assert.That(loaded.Worksheet(1).AddPicture(new MemoryStream(Png)).Id).IsEqualTo(11);
    }

    [Test]
    public async Task A_new_picture_takes_the_id_after_the_highest_through_any_run_of_changes()
    {
        // Adds, deletes and new ids in a fixed random order. After each, a new picture takes one
        // more than the highest id the sheet's pictures hold, read from the pictures themselves.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var random = new Random(804);
        for (var step = 0; step < 300; step++)
        {
            var pictures = ws.Pictures.Cast<XLPicture>().ToList();
            var choice = random.Next(3);
            if (choice == 0 || pictures.Count < 2)
            {
                var expected = pictures.Select(p => p.Id).DefaultIfEmpty(0).Max() + 1;
                await Assert.That(ws.AddPicture(new MemoryStream(Png)).Id).IsEqualTo(expected);
            }
            else if (choice == 1)
            {
                pictures[random.Next(pictures.Count)].Delete();
            }
            else
            {
                var free = Enumerable.Range(1, 40).Except(pictures.Select(p => p.Id)).ToList();
                pictures[random.Next(pictures.Count)].Id = free[random.Next(free.Count)];
            }
        }
    }

    [Test]
    public async Task Pictures_grouped_in_a_save_that_removes_a_grouped_picture_are_written_with_their_new_names()
    {
        // The save removes Picture 2 from the loaded group, then builds the new group, then writes
        // each grouped picture's name into the drawing: the new group's members are found there.
        using var withFreePictures = LoadAndSave(Resource("GroupedPictures.xlsx"), ws =>
        {
            ws.AddPicture(new MemoryStream(Png), "Free A").MoveTo(600, 20);
            ws.AddPicture(new MemoryStream(Png), "Free B").MoveTo(600, 300);
        });

        using var saved = LoadAndSave(withFreePictures, ws =>
        {
            ws.Picture("Picture 2").Delete();
            ws.Pictures.Group(ws.Picture("Free A"), ws.Picture("Free B"));
            ws.Picture("Free A").Name = "Grouped A";
            ws.Picture("Free B").Name = "Grouped B";
        });

        await Assert.That(GroupedPictureNames(saved)).IsEqualTo("Picture 1 | Grouped A, Grouped B");
    }

    private static MemoryStream LoadAndSave(MemoryStream source, Action<IXLWorksheet> edit)
    {
        var saved = new MemoryStream();
        source.Position = 0;
        using (var wb = new XLWorkbook(source))
        {
            edit(wb.Worksheet(1));
            wb.SaveAs(saved);
        }

        return saved;
    }

    private static MemoryStream Resource(string name)
    {
        using var resource = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath($@"Other\Drawings\{name}"));
        var package = new MemoryStream();
        resource.CopyTo(package);
        package.Position = 0;
        return package;
    }

    /// <summary>The names of the pictures in each group of the saved drawing, a group to a "|".</summary>
    private static string GroupedPictureNames(Stream saved) =>
        string.Join(" | ", XDocument.Parse(saved.ReadPart(Drawing1)).Descendants(XdrNs + "grpSp")
            .Select(group => string.Join(", ", group.Elements(XdrNs + "pic")
                .Select(pic => (string?)pic.Descendants(XdrNs + "cNvPr").Single().Attribute("name")))));
}
