using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// #585, following #577. A conditional format's <c>pivotAreas</c> — both the 2007
/// <c>conditionalFormats</c> list and the <c>x14</c> one — can hold a reference on the 'data' field
/// that names a value by its <em>position</em> in <c>dataFields</c>, exactly like the style formats
/// #577 fixed. Removing a value must renumber or drop those references too, and a copy of the sheet
/// must not let a later removal on one table renumber the other's.
/// </summary>
/// <remarks>
/// No Excel-written fixture holds a pivot table with a conditional format that names a data field,
/// so these build the pivot table and its formats through the object model instead, the way
/// <see cref="PivotConditionalFormatLinkTests"/> does, and assert on the saved
/// <c>xl/pivotTables/pivotTable1.xml</c> the way <see cref="XLPivotStyleFormatPruningTests"/> asserts
/// on the style formats.
/// </remarks>
public class XLPivotConditionalFormatPruningTests
{
    /// <summary>
    /// What a reference's <c>field</c> holds when it is on the 'data' field rather than a pivot
    /// field: -2 unsigned, written as 4294967294.
    /// </summary>
    private const uint DataFieldReferenceIndex = unchecked((uint)-2);

    [Test]
    [Property("Description", "#585: removing the first of several values must renumber the conditional formats naming the values after it, in both lists")]
    public async Task Removing_a_value_renumbers_conditional_formats_naming_the_values_after_it()
    {
        var (wb, sheet, pt) = BuildPivotTableWithThreeValues();
        using (wb)
        {
            // 2007 list: one format naming only the removed value, one naming the value after it,
            // one naming all three.
            AddPivotConditionalFormat(sheet, pt, 0);
            AddPivotConditionalFormat(sheet, pt, 1);
            AddPivotConditionalFormat(sheet, pt, 0, 1, 2);

            // x14 list: same shape, so the two lists are exercised by the one removal together.
            AddPivotExtensionConditionalFormat(pt, "{11111111-1111-1111-1111-111111111111}", 101, 0);
            AddPivotExtensionConditionalFormat(pt, "{22222222-2222-2222-2222-222222222222}", 102, 2);

            pt.Values.Remove(pt.Values.First().CustomName);

            using var saved = new MemoryStream();
            wb.SaveAs(saved);

            var mainList = ConditionalFormatPositions(saved, sheet.Name);
            await Assert.That(mainList.Count).IsEqualTo(2)
                .Because("the format naming only the removed value goes with it");
            await Assert.That(mainList).IsEquivalentTo(new List<List<uint>> { new() { 0 }, new() { 0, 1 } })
                .Because("a surviving reference follows its own value down a position");

            var extensionList = ExtensionConditionalFormatPositions(saved, sheet.Name);
            await Assert.That(extensionList.Count).IsEqualTo(1)
                .Because("the x14 list is pruned in the same operation as the 2007 list, not separately");
            await Assert.That(extensionList).IsEquivalentTo(new List<List<uint>> { new() { 1 } });
        }
    }

    /// <summary>
    /// <see cref="XLPivotTable.CopyConditionalFormatsTo"/> is exercised directly, on two
    /// independently built pivot tables, rather than through a sheet copy
    /// (<c>IXLWorksheet.CopyTo</c>). A sheet copy also copies each pivot field through
    /// <c>XLPivotTable.CopyTo</c>, which throws on the <c>Values</c> sentinel field that a table of
    /// three values puts on an axis (<see cref="XLPivotFieldBase.SortType"/> cannot be set on a
    /// data field) — a pre-existing defect unrelated to #585. Calling the method under test
    /// directly keeps this test independent of that one.
    /// </summary>
    [Test]
    [Property("Description", "#585: a copy's conditional formats must not be renumbered when a value is removed from the original")]
    public async Task Removing_a_value_from_the_original_does_not_renumber_the_copys_conditional_formats()
    {
        using var wb = new XLWorkbook();
        var data = BuildDataSheet(wb);
        var pt = AddPivotTableWithThreeValues(wb, data, "Pivots");
        var copy = AddPivotTableWithThreeValues(wb, data, "Copy");

        AddPivotConditionalFormat((XLWorksheet)pt.Worksheet, pt, 0, 1);
        AddPivotExtensionConditionalFormat(pt, "{33333333-3333-3333-3333-333333333333}", 201, 0, 1);

        pt.CopyConditionalFormatsTo(copy);

        pt.Values.Remove(pt.Values.First().CustomName);

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var originalMain = ConditionalFormatPositions(saved, "Pivots");
        await Assert.That(originalMain).IsEquivalentTo(new List<List<uint>> { new() { 0 } })
            .Because("the original's own format follows its value down a position");

        var copyMain = ConditionalFormatPositions(saved, "Copy");
        await Assert.That(copyMain).IsEquivalentTo(new List<List<uint>> { new() { 0, 1 } })
            .Because("the copy never lost a value, so its format must still name both of them");

        var originalExtension = ExtensionConditionalFormatPositions(saved, "Pivots");
        await Assert.That(originalExtension).IsEquivalentTo(new List<List<uint>> { new() { 0 } });

        var copyExtension = ExtensionConditionalFormatPositions(saved, "Copy");
        await Assert.That(copyExtension).IsEquivalentTo(new List<List<uint>> { new() { 0, 1 } });
    }

    /// <inheritdoc cref="Removing_a_value_from_the_original_does_not_renumber_the_copys_conditional_formats"/>
    [Test]
    [Property("Description", "#585: a value removed from a copy must not renumber the original's conditional formats")]
    public async Task Removing_a_value_from_the_copy_does_not_renumber_the_originals_conditional_formats()
    {
        using var wb = new XLWorkbook();
        var data = BuildDataSheet(wb);
        var pt = AddPivotTableWithThreeValues(wb, data, "Pivots");
        var copy = AddPivotTableWithThreeValues(wb, data, "Copy");

        AddPivotConditionalFormat((XLWorksheet)pt.Worksheet, pt, 0, 1);

        pt.CopyConditionalFormatsTo(copy);

        copy.Values.Remove(copy.Values.First().CustomName);

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var originalMain = ConditionalFormatPositions(saved, "Pivots");
        await Assert.That(originalMain).IsEquivalentTo(new List<List<uint>> { new() { 0, 1 } })
            .Because("the original never lost a value, so its own format must be untouched");

        var copyMain = ConditionalFormatPositions(saved, "Copy");
        await Assert.That(copyMain).IsEquivalentTo(new List<List<uint>> { new() { 0 } })
            .Because("the copy's own format follows its value down a position");
    }

    private static (XLWorkbook Workbook, XLWorksheet Sheet, XLPivotTable PivotTable) BuildPivotTableWithThreeValues()
    {
        var wb = new XLWorkbook();
        var data = BuildDataSheet(wb);
        var pt = AddPivotTableWithThreeValues(wb, data, "Pivots");
        return (wb, (XLWorksheet)pt.Worksheet, pt);
    }

    private static XLWorksheet BuildDataSheet(XLWorkbook wb)
    {
        var data = wb.AddWorksheet("Data");
        data.Cell("A1").Value = "Label";
        data.Cell("B1").Value = "V1";
        data.Cell("C1").Value = "V2";
        data.Cell("D1").Value = "V3";
        data.Cell("A2").Value = "Row1";
        data.Cell("B2").Value = 1;
        data.Cell("C2").Value = 2;
        data.Cell("D2").Value = 3;
        return (XLWorksheet)data;
    }

    private static XLPivotTable AddPivotTableWithThreeValues(XLWorkbook wb, XLWorksheet data, string pivotSheetName)
    {
        var pivots = (XLWorksheet)wb.AddWorksheet(pivotSheetName);
        var pt = (XLPivotTable)pivots.PivotTables.Add("pt", pivots.Cell("A1")!, data.Range("A1:D2")!);
        pt.RowLabels.Add("Label");
        pt.Values.Add("V1");
        pt.Values.Add("V2");
        pt.Values.Add("V3");
        return pt;
    }

    private static XLPivotArea DataFieldArea(params uint[] positions)
    {
        var area = new XLPivotArea();
        var reference = new XLPivotReference { Field = DataFieldReferenceIndex };
        foreach (var position in positions)
            reference.AddFieldItem(position);

        area.AddReference(reference);
        return area;
    }

    private static XLPivotConditionalFormat AddPivotConditionalFormat(XLWorksheet sheet, XLPivotTable pt, params uint[] positions)
    {
        var rule = (XLConditionalFormat)sheet.Range("A1:A1")!.AddConditionalFormat();
        rule.WhenIsTrue("TRUE");

        // Kept only on the pivot table's own list, the way a pivot table's rule really is (mirrors
        // PivotConditionalFormatLinkTests.The_main_list_names_its_rule_by_the_priority_the_sheet_wrote).
        sheet.ConditionalFormats.Remove(f => f == rule);

        var pivotFormat = new XLPivotConditionalFormat(rule);
        pivotFormat.AddArea(DataFieldArea(positions));
        pt.AddConditionalFormat(pivotFormat);
        return pivotFormat;
    }

    private static XLPivotExtensionConditionalFormat AddPivotExtensionConditionalFormat(XLPivotTable pt, string ruleId, uint priority, params uint[] positions)
    {
        var format = new XLPivotExtensionConditionalFormat(ruleId, priority);
        format.AddArea(DataFieldArea(positions));
        pt.AddExtensionConditionalFormat(format);
        return format;
    }

    /// <summary>
    /// The 'data' field positions named by <paramref name="element"/>'s <c>pivotArea</c>
    /// reference(s), read off the saved XML. The <c>pivotArea</c> and its
    /// <c>references</c>/<c>reference</c> children are in the main namespace whichever list holds
    /// them (the writer's own note on <c>WriteConditionalFormatAreas</c>), so the plain
    /// <c>DocumentFormat.OpenXml.Spreadsheet</c> types find them under an <c>x14:conditionalFormat</c>
    /// parent too.
    /// </summary>
    private static List<uint> DataFieldPositions(OpenXmlElement element)
    {
        return element.Descendants<S.PivotAreaReference>()
            .Where(reference => reference.Field?.Value == DataFieldReferenceIndex)
            .SelectMany(reference => reference.Elements<S.FieldItem>())
            .Select(fieldItem => fieldItem.Val!.Value)
            .ToList();
    }

    private static List<List<uint>> ConditionalFormatPositions(Stream saved, string sheetName)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var definition = Definition(doc, sheetName);
        return definition.Elements<S.ConditionalFormats>()
            .SelectMany(list => list.Elements<S.ConditionalFormat>())
            .Select(DataFieldPositions)
            .ToList();
    }

    private static List<List<uint>> ExtensionConditionalFormatPositions(Stream saved, string sheetName)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var definition = Definition(doc, sheetName);
        return definition.Descendants<X14.ConditionalFormats>()
            .SelectMany(list => list.Elements<X14.ConditionalFormat>())
            .Select(DataFieldPositions)
            .ToList();
    }

    private static S.PivotTableDefinition Definition(SpreadsheetDocument doc, string sheetName)
    {
        var workbookPart = doc.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().Single(s => s.Name == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return worksheetPart.GetPartsOfType<PivotTablePart>().Single().PivotTableDefinition!;
    }
}
