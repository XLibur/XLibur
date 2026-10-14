using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A sheet whose cells carry a <c>cm</c> index keeps its part, or its cells, on the first save of a
/// loaded workbook when the sheet has not changed (#702, #721).
/// </summary>
/// <remarks>
/// A <c>cm</c> index names a block of the cell metadata in the metadata part. Only the load sets one,
/// so no change to the model moves it without editing the cell. The save writes the metadata part
/// back as the file had it, adding the dynamic-array block when it needs one and the file has none.
/// The tests check that each such save leaves the kept index naming the block it named in the file,
/// as <c>KeepSheetDataTests</c> does for shared strings. The cell is a plain value: a sheet with a
/// formula is excluded from the fast path for its own reasons.
/// </remarks>
public class KeepCellMetadataSheetTests
{
    private const string Sheet1 = "xl/worksheets/sheet1.xml";

    [Test]
    public async Task An_unchanged_sheet_with_a_cm_index_keeps_its_part()
    {
        using var source = Source(renameDynamicArrayType: false);

        using var saved = source.LoadAndSave();

        await Assert.That(saved.IsKeptWhole(Sheet1)).IsTrue();
        await Assert.That(saved.Sheet1Xml()).Contains("cm=\"1\"");
        await Assert.That(CellMetadataTypeName(saved, 1)).IsEqualTo("XLDAPR");
        await AssertValues(saved);
    }

    [Test]
    public async Task A_dynamic_array_block_the_save_adds_leaves_the_kept_index_naming_its_block()
    {
        // The file's only cell metadata type is not XLDAPR, so the new dynamic array makes the save
        // append an XLDAPR type and a block for it, after the block the kept cell names.
        using var source = Source(renameDynamicArrayType: true);

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Other").Cell("C1").SetDynamicFormulaA1("SEQUENCE(2)"));

        await Assert.That(saved.HasKeptCells(Sheet1)).IsTrue();
        await Assert.That(saved.Sheet1Xml()).Contains("cm=\"1\"");
        await Assert.That(CellMetadataTypeName(saved, 1)).IsEqualTo("XLFAKE");
        await Assert.That(CellMetadataTypeName(saved, 2)).IsEqualTo("XLDAPR");
        await AssertValues(saved);
    }

    [Test]
    public async Task Removing_the_last_dynamic_array_leaves_the_kept_index_naming_its_block()
    {
        using var source = Source(renameDynamicArrayType: false);

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Other").Cell("A1").Clear());

        await Assert.That(saved.HasKeptCells(Sheet1)).IsTrue();
        await Assert.That(CellMetadataTypeName(saved, 1)).IsEqualTo("XLDAPR");
        await AssertValues(saved);
    }

    [Test]
    public async Task An_edited_sheet_with_a_cm_index_is_written_from_the_model()
    {
        using var source = Source(renameDynamicArrayType: false);

        using var saved = source.LoadAndSave(wb => wb.Worksheet("Data").Cell("A3").Value = "new");

        await Assert.That(saved.IsWrittenFromModel(Sheet1)).IsTrue();
        await Assert.That(saved.Sheet1Xml()).Contains("cm=\"1\"");
    }

    /// <summary>
    /// Sheet Data holds "Header" in A1, 1.5 in B2 and 7 in C1, which has <c>cm="1"</c>; sheet Other
    /// holds a dynamic array, for which the metadata part has its XLDAPR block first. Data's part is
    /// marked by <see cref="KeptSheet.Mark"/>.
    /// </summary>
    /// <param name="renameDynamicArrayType">
    /// Renames the XLDAPR type of the metadata part to XLFAKE, so the load takes the formula for an
    /// array formula and the file has no dynamic-array metadata.
    /// </param>
    private static MemoryStream Source(bool renameDynamicArrayType)
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Header";
            ws.Cell("B2").Value = 1.5;
            ws.Cell("C1").Value = 7;

            var other = wb.AddWorksheet("Other");
            other.Cell("A1").SetDynamicFormulaA1("SEQUENCE(3)");
            wb.RecalculateAllFormulas();
            wb.SaveAs(stream);
        }

        stream.RewriteSheet1(xml => ReplaceOnce(xml, "<x:c r=\"C1\"", "<x:c r=\"C1\" cm=\"1\""));
        if (renameDynamicArrayType)
            stream.RewritePart(MetadataPart(stream), xml => xml.Replace("XLDAPR", "XLFAKE", StringComparison.Ordinal));

        return stream.Mark(Sheet1);
    }

    private static string ReplaceOnce(string text, string oldValue, string newValue)
    {
        var index = text.IndexOf(oldValue, StringComparison.Ordinal);
        if (index < 0)
            throw new InvalidOperationException($"'{oldValue}' not found.");

        return text[..index] + newValue + text[(index + oldValue.Length)..];
    }

    private static string MetadataPart(Stream package) => package.PartsUnder("xl/metadata").Single();

    /// <summary>The name of the metadata type the 1-based cell metadata block <paramref name="cm"/> refers to.</summary>
    private static string CellMetadataTypeName(Stream package, int cm)
    {
        XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var metadata = XDocument.Parse(package.ReadPart(MetadataPart(package))).Root!;
        var types = metadata.Element(main + "metadataTypes")!.Elements(main + "metadataType").ToList();
        var block = metadata.Element(main + "cellMetadata")!.Elements(main + "bk").ElementAt(cm - 1);
        var typeIndex = (int)block.Element(main + "rc")!.Attribute("t")!;
        return (string)types[typeIndex - 1].Attribute("name")!;
    }

    private static async Task AssertValues(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Data");

        await Assert.That(ws.Cell("A1").GetText()).IsEqualTo("Header");
        await Assert.That(ws.Cell("B2").GetDouble()).IsEqualTo(1.5);
        await Assert.That(ws.Cell("C1").GetDouble()).IsEqualTo(7);
    }
}
