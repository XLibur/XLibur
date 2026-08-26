using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Rows;

/// <summary>
/// Row outline levels were counted into the column counter (XLRow.OutlineLevel, copied verbatim from
/// XLColumn.OutlineLevel), so sheetFormatPr/@outlineLevelRow was never emitted and row groups inflated
/// @outlineLevelCol instead. Spec 26 task 1. Nothing asserted either attribute before this file.
/// </summary>
public class OutlineRoundTripTests
{
    private static XElement SheetFormatPr(Stream xlsx)
    {
        xlsx.Position = 0;
        using var doc = SpreadsheetDocument.Open(xlsx, isEditable: false);
        var part = doc.WorkbookPart!.WorksheetParts.Single();
        using var stream = part.GetStream();
        var xml = XDocument.Load(stream);
        var ns = xml.Root!.Name.Namespace;
        return xml.Root.Element(ns + "sheetFormatPr")!;
    }

    [Test]
    public async Task Grouping_rows_emits_outlineLevelRow_and_not_outlineLevelCol()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "x";
            ws.Rows(2, 4).Group();
            ws.Rows(3, 3).Group();     // level 2
            wb.SaveAs(ms);
        }

        var sfp = SheetFormatPr(ms);
        await Assert.That(sfp.Attribute("outlineLevelRow")?.Value).IsEqualTo("2");
        await Assert.That(sfp.Attribute("outlineLevelCol")).IsNull();
    }

    [Test]
    public async Task Grouping_columns_emits_outlineLevelCol_and_not_outlineLevelRow()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "x";
            ws.Columns(2, 4).Group();
            wb.SaveAs(ms);
        }

        var sfp = SheetFormatPr(ms);
        await Assert.That(sfp.Attribute("outlineLevelCol")?.Value).IsEqualTo("1");
        await Assert.That(sfp.Attribute("outlineLevelRow")).IsNull();
    }

    /// <summary>
    /// The load path sets XLRow.OutlineLevel (WorksheetSheetDataReader), so before the fix, re-saving
    /// a file with row groups inflated that file's @outlineLevelCol. This is the round-trip half of
    /// the defect.
    /// </summary>
    [Test]
    public async Task Reloading_and_resaving_row_groups_does_not_inflate_outlineLevelCol()
    {
        using var first = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "x";
            ws.Rows(2, 4).Group();
            wb.SaveAs(first);
        }

        using var second = new MemoryStream();
        first.Position = 0;
        using (var wb = new XLWorkbook(first))
            wb.SaveAs(second);

        var sfp = SheetFormatPr(second);
        await Assert.That(sfp.Attribute("outlineLevelRow")?.Value).IsEqualTo("1");
        await Assert.That(sfp.Attribute("outlineLevelCol")).IsNull();
    }

    /// <summary>
    /// GetMaxRowOutline guarded the dictionary's size rather than the filtered sequence's, so a
    /// dictionary holding only zero counts made .Max() throw on an empty sequence. Unreachable while
    /// row outlines were never counted; reachable the moment task 1's first change lands.
    /// RowTests.UngroupFromAll performs exactly this sequence but never saves.
    /// </summary>
    [Test]
    public async Task Grouping_then_ungrouping_every_row_still_saves()
    {
        using var ms = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "x";
        ws.Rows(1, 2).Group();
        ws.Rows(1, 2).Ungroup(true);

        wb.SaveAs(ms);

        var sfp = SheetFormatPr(ms);
        await Assert.That(sfp.Attribute("outlineLevelRow")).IsNull();
    }
}
