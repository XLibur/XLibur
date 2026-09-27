using System;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// Saving a workbook XLibur created, then loading and saving that output unchanged, must give the
/// same worksheet parts, byte for byte. A save that can skip an unchanged part (#702) compares
/// what it would write with what the file holds, and needs XLibur's own output to be stable.
/// </summary>
public class ResaveStabilityTests
{
    [Test]
    public async Task A_new_sheet_saves_the_same_bytes_when_loaded_and_saved_again()
    {
        using var created = Save(wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Header";
            ws.Cell("A2").Value = 12.5;
        });

        using var resaved = LoadAndSave(created);

        await AssertWorksheetPartsEqual(created, resaved);
    }

    [Test]
    public async Task A_new_sheet_with_structure_saves_the_same_bytes_when_loaded_and_saved_again()
    {
        using var created = Save(wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Column(2).Style.Font.Italic = true;
            ws.Cell("A1").Value = "Header";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("B2").Value = 2.5;
            ws.Cell("C3").Value = new DateTime(2026, 9, 27);
            ws.Cell("D4").Value = true;
            ws.Cell("D4").Style.Fill.BackgroundColor = XLColor.Yellow;
            ws.Row(6).Height = 30;
            ws.Range("A5:C5").Merge();
            ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            wb.AddWorksheet("Empty");
        });

        using var resaved = LoadAndSave(created);

        await AssertWorksheetPartsEqual(created, resaved);
    }

    [Test]
    public async Task A_sheet_written_in_the_default_namespace_saves_the_same_bytes_from_its_first_save_on()
    {
        // Excel writes <worksheet xmlns="…main">, with no prefix. XLibur saves it with an x: prefix,
        // and that first save has to be one the next save repeats.
        using var created = Save(wb =>
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Header";
            ws.Cell("A2").Value = 12.5;
        });
        using var defaultNamespace = Copy(created).RewriteSheet1(xml => xml
            .Replace("<x:", "<", StringComparison.Ordinal)
            .Replace("</x:", "</", StringComparison.Ordinal)
            .Replace("xmlns:x=", "xmlns=", StringComparison.Ordinal));
        await Assert.That(defaultNamespace.Sheet1Xml()).DoesNotContain("<x:");

        using var firstSave = LoadAndSave(defaultNamespace);
        using var secondSave = LoadAndSave(firstSave);

        await AssertWorksheetPartsEqual(firstSave, secondSave);
    }

    private static async Task AssertWorksheetPartsEqual(MemoryStream expected, MemoryStream actual)
    {
        var parts = expected.PartsUnder("xl/worksheets/sheet");
        await Assert.That(parts.Length).IsGreaterThan(0);
        await Assert.That(actual.PartsUnder("xl/worksheets/sheet")).IsEquivalentTo(parts);

        foreach (var part in parts)
            await Assert.That(actual.ReadPart(part)).IsEqualTo(expected.ReadPart(part));
    }

    private static MemoryStream Save(Action<XLWorkbook> build)
    {
        var output = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            build(wb);
            wb.SaveAs(output);
        }

        output.Position = 0;
        return output;
    }

    /// <summary>An expandable copy, which a package rewrite needs.</summary>
    private static MemoryStream Copy(MemoryStream package)
    {
        var copy = new MemoryStream();
        copy.Write(package.GetBuffer(), 0, (int)package.Length);
        copy.Position = 0;
        return copy;
    }

    private static MemoryStream LoadAndSave(MemoryStream package)
    {
        var output = new MemoryStream();
        using (var wb = new XLWorkbook(new MemoryStream(package.ToArray())))
            wb.SaveAs(output);

        output.Position = 0;
        return output;
    }
}
