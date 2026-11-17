using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.PivotTables;

public class PivotTableSaveFidelityTests
{
    private const uint FirstCustomNumberFormatId = 164;
    private const string X14PivotTableDefinitionUri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}";

    [Test]
    public async Task Save_keeps_a_custom_pivot_table_style_name()
    {
        await AssertSameAfterSave(@"Other\Lion\PivotTables\CustomPivotTheme.xlsx",
            definition => definition.PivotTableStyle?.Name?.Value);
    }

    [Test]
    public async Task A_theme_set_in_code_replaces_a_custom_style_name()
    {
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"Other\Lion\PivotTables\CustomPivotTheme.xlsx"));
        var saved = new MemoryStream();
        using (var wb = new XLWorkbook(stream))
        {
            wb.Worksheets.SelectMany(ws => ws.PivotTables).Single().Theme = XLPivotTableTheme.PivotStyleMedium2;
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        var names = ReadPivotTables(saved, (definition, _) => definition.PivotTableStyle?.Name?.Value);
        await Assert.That(names.Values.Single()).IsEqualTo("PivotStyleMedium2");
    }

    [Test]
    [Arguments(@"Other\Lion\PivotTables\PivotWithStyles.xlsx")]
    [Arguments(@"TryToLoad\TemplateWithTableSourcePivotTables.xlsx")]
    public async Task Save_keeps_formats_whose_dxf_clears_a_property(string resource)
    {
        var source = ReadPivotTables(resource, FormatDxfs);
        var saved = ReadPivotTables(LoadAndSave(resource), FormatDxfs);

        await Assert.That(saved.Keys).IsEquivalentTo(source.Keys);
        foreach (var (key, sourceDxfs) in source)
        {
            var savedDxfs = saved[key];
            await Assert.That(savedDxfs.Count).IsEqualTo(sourceDxfs.Count);
            for (var i = 0; i < sourceDxfs.Count; i++)
            {
                await Assert.That(savedDxfs[i] is null).IsEqualTo(sourceDxfs[i] is null);
                if (sourceDxfs[i] is { } sourceDxf && Clears(sourceDxf))
                    await Assert.That(savedDxfs[i]).IsEqualTo(sourceDxf);
            }
        }
    }

    [Test]
    public async Task Save_keeps_extensions_xlibur_does_not_write()
    {
        await AssertSameAfterSave(@"TryToLoad\load_pivottable_renamedcolumns_2591.xlsx",
            definition => string.Join("\n", definition.GetFirstChild<PivotTableDefinitionExtensionList>()?
                .Elements<PivotTableDefinitionExtension>()
                .Where(ext => ext.Uri?.Value != X14PivotTableDefinitionUri)
                .Select(ext => WithoutNamespaceDeclarations(ext.OuterXml)) ?? []));
    }

    [Test]
    public async Task Save_keeps_the_fill_down_labels_default()
    {
        await AssertSameAfterSave(@"TryToLoad\TemplateWithTableSourcePivotTables.xlsx",
            definition => definition.Descendants<X14.PivotTableDefinition>().SingleOrDefault()?
                .FillDownLabelsDefault?.Value);
    }

    [Test]
    [Arguments(@"TryToLoad\Timelines_Missing_21232.xlsx")]
    [Arguments(@"TryToLoad\DialogSheet.xlsx")]
    [Arguments(@"Other\ExcelAuthored\PivotFilters.xlsx")]
    [Arguments(@"TryToLoad\load_pivottable_renamedcolumns_2591.xlsx")]
    public async Task Save_keeps_cache_field_number_formats(string resource)
    {
        var source = ReadCacheFieldNumberFormats(TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource)));
        var saved = ReadCacheFieldNumberFormats(LoadAndSave(resource));

        await Assert.That(source.Values.SelectMany(f => f).Any(id => id is not null)).IsTrue();
        await Assert.That(saved.Keys).IsEquivalentTo(source.Keys);
        foreach (var (relId, sourceFormats) in source)
            await Assert.That(saved[relId]).IsEquivalentTo(sourceFormats, TUnit.Assertions.Enums.CollectionOrdering.Matching);
    }

    private static async Task AssertSameAfterSave<T>(string resource, System.Func<PivotTableDefinition, T> read)
    {
        var source = ReadPivotTables(resource, (definition, _) => read(definition));
        var saved = ReadPivotTables(LoadAndSave(resource), (definition, _) => read(definition));

        await Assert.That(source.Values.Any(v => v is string { Length: > 0 } or true)).IsTrue();
        await Assert.That(saved).IsEquivalentTo(source);
    }

    private static List<string?> FormatDxfs(PivotTableDefinition definition, DifferentialFormat[] dxfs)
    {
        return definition.Formats?.Elements<Format>()
            .Select(format => format.FormatId?.Value is { } id ? WithoutNamespaceDeclarations(dxfs[id].OuterXml) : null)
            .ToList() ?? [];
    }

    private static bool Clears(string dxf)
        => dxf.Contains("patternType=\"none\"") || dxf.Contains("<right />") || dxf.Contains("<left />");

    private static Dictionary<string, T> ReadPivotTables<T>(string resource,
        System.Func<PivotTableDefinition, DifferentialFormat[], T> read)
        => ReadPivotTables(TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource)), read);

    private static Dictionary<string, T> ReadPivotTables<T>(Stream stream,
        System.Func<PivotTableDefinition, DifferentialFormat[], T> read)
    {
        var result = new Dictionary<string, T>();
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var dxfs = workbookPart.WorkbookStylesPart?.Stylesheet?.DifferentialFormats?
            .Elements<DifferentialFormat>().ToArray() ?? [];
        foreach (var sheet in workbookPart.Workbook!.Sheets!.Elements<Sheet>())
        {
            if (workbookPart.GetPartById(sheet.Id!.Value!) is not WorksheetPart worksheetPart)
                continue;

            foreach (var pivotTablePart in worksheetPart.PivotTableParts)
            {
                var definition = pivotTablePart.PivotTableDefinition!;
                result.Add($"{sheet.Name!.Value}!{definition.Name!.Value}", read(definition, dxfs));
            }
        }

        return result;
    }

    private static Dictionary<string, List<string?>> ReadCacheFieldNumberFormats(Stream stream)
    {
        var result = new Dictionary<string, List<string?>>();
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var customFormats = workbookPart.WorkbookStylesPart?.Stylesheet?.NumberingFormats?
            .Elements<NumberingFormat>()
            .ToDictionary(nf => nf.NumberFormatId!.Value, nf => nf.FormatCode!.Value) ?? [];
        foreach (var pivotCache in workbookPart.Workbook!.PivotCaches?.Elements<PivotCache>() ?? [])
        {
            var part = (PivotTableCacheDefinitionPart)workbookPart.GetPartById(pivotCache.Id!.Value!);
            result.Add(pivotCache.Id.Value!, part.PivotCacheDefinition!.CacheFields!.Elements<CacheField>()
                .Select(f => f.NumberFormatId?.Value switch
                {
                    null => null,
                    >= FirstCustomNumberFormatId and var id => customFormats[id],
                    var id => id.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                })
                .ToList());
        }

        return result;
    }

    private static MemoryStream LoadAndSave(string resource)
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource));
        using var wb = new XLWorkbook(stream);
        var saved = new MemoryStream();
        wb.SaveAs(saved);
        saved.Position = 0;
        return saved;
    }

    private static string WithoutNamespaceDeclarations(string xml)
    {
        var element = XElement.Parse(xml);
        element.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        return element.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }
}
