using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// #585, following #577. A chart format's <c>pivotArea</c> can hold a reference on the 'data' field
/// that names a value by its <em>position</em> in <c>dataFields</c>, exactly like the style formats
/// #577 fixed. Removing a value must renumber or drop those references too, not only the style
/// formats'.
/// </summary>
/// <remarks>
/// No Excel-written fixture holds a pivot table with a chart format that names a data field, so this
/// builds the pivot table and its chart formats through the object model instead, the way
/// <see cref="PivotChartFormatTests"/> does, and asserts on the saved
/// <c>xl/pivotTables/pivotTable1.xml</c> the way <see cref="XLPivotStyleFormatPruningTests"/> asserts
/// on the style formats.
/// </remarks>
public class XLPivotChartFormatPruningTests
{
    /// <summary>
    /// What a reference's <c>field</c> holds when it is on the 'data' field rather than a pivot
    /// field: -2 unsigned, written as 4294967294.
    /// </summary>
    private const uint DataFieldReferenceIndex = unchecked((uint)-2);

    [Test]
    [Property("Description", "#585: removing the first of several values must renumber the chart formats naming the values after it")]
    public async Task Removing_a_value_renumbers_chart_formats_naming_the_values_after_it()
    {
        var (wb, sheet, pt) = BuildPivotTableWithThreeValues();
        using (wb)
        {
            AddPivotChartFormat(pt, 0, 3, 0);
            AddPivotChartFormat(pt, 1, 7, 1);
            AddPivotChartFormat(pt, 2, 9, 0, 1, 2);

            pt.Values.Remove(pt.Values.First().CustomName);

            using var saved = new MemoryStream();
            wb.SaveAs(saved);

            var chartFormats = ChartFormatPositions(saved, sheet.Name);
            await Assert.That(chartFormats.Count).IsEqualTo(2)
                .Because("the chart format naming only the removed value goes with it, exactly as a style format would");
            await Assert.That(chartFormats.Select(f => f.Chart)).IsEquivalentTo(new uint[] { 1, 2 })
                .Because("chart 0's format named only the removed value");
            await Assert.That(chartFormats.Single(f => f.Chart == 1).Positions).IsEquivalentTo(new List<uint> { 0 });
            await Assert.That(chartFormats.Single(f => f.Chart == 2).Positions).IsEquivalentTo(new List<uint> { 0, 1 });
        }
    }

    private static (XLWorkbook Workbook, XLWorksheet Sheet, XLPivotTable PivotTable) BuildPivotTableWithThreeValues()
    {
        var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        data.Cell("A1").Value = "Label";
        data.Cell("B1").Value = "V1";
        data.Cell("C1").Value = "V2";
        data.Cell("D1").Value = "V3";
        data.Cell("A2").Value = "Row1";
        data.Cell("B2").Value = 1;
        data.Cell("C2").Value = 2;
        data.Cell("D2").Value = 3;

        var pivots = (XLWorksheet)wb.AddWorksheet("Pivots");
        var pt = (XLPivotTable)pivots.PivotTables.Add("pt", pivots.Cell("A1")!, data.Range("A1:D2")!);
        pt.RowLabels.Add("Label");
        pt.Values.Add("V1");
        pt.Values.Add("V2");
        pt.Values.Add("V3");

        return (wb, pivots, pt);
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

    private static XLPivotChartFormat AddPivotChartFormat(XLPivotTable pt, uint chart, uint formatId, params uint[] positions)
    {
        var chartFormat = new XLPivotChartFormat(DataFieldArea(positions)) { Chart = chart, Format = formatId };
        pt.AddChartFormat(chartFormat);
        return chartFormat;
    }

    /// <summary>
    /// The 'data' field positions named by <paramref name="element"/>'s <c>pivotArea</c>
    /// reference(s), read off the saved XML.
    /// </summary>
    private static List<uint> DataFieldPositions(OpenXmlElement element)
    {
        return element.Descendants<S.PivotAreaReference>()
            .Where(reference => reference.Field?.Value == DataFieldReferenceIndex)
            .SelectMany(reference => reference.Elements<S.FieldItem>())
            .Select(fieldItem => fieldItem.Val!.Value)
            .ToList();
    }

    private static List<(uint Chart, List<uint> Positions)> ChartFormatPositions(Stream saved, string sheetName)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var definition = Definition(doc, sheetName);
        return definition.Elements<S.ChartFormats>()
            .SelectMany(list => list.Elements<S.ChartFormat>())
            .Select(chartFormat => (chartFormat.Chart!.Value, DataFieldPositions(chartFormat)))
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
