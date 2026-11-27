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
    public async Task Save_keeps_the_number_format_of_each_pivot_format()
    {
        const string resource = @"TryToLoad\TemplateWithTableSourcePivotTables.xlsx";
        var source = ReadPivotTables(resource, FormatNumberFormats);
        var saved = ReadPivotTables(LoadAndSave(resource), FormatNumberFormats);

        await Assert.That(source.Values.SelectMany(f => f).Any(code => code == "0.00%")).IsTrue();
        await Assert.That(saved.Keys).IsEquivalentTo(source.Keys);
        foreach (var (key, sourceFormats) in source)
            await Assert.That(saved[key]).IsEquivalentTo(sourceFormats, TUnit.Assertions.Enums.CollectionOrdering.Matching)
                .Because(key);
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

    [Test]
    public async Task Save_keeps_the_unused_flag_of_each_shared_item_and_contains_non_date()
    {
        const string resource = @"TryToLoad\load_pivottable_renamedcolumns_2591.xlsx";
        var source = ReadCacheDefinitions(TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource)), SharedItemFlags);
        var saved = ReadCacheDefinitions(LoadAndSave(resource), SharedItemFlags);

        await Assert.That(source.Values.SelectMany(fields => fields).Count(field => field.Contains("@u"))).IsGreaterThan(0);
        await Assert.That(saved.Keys).IsEquivalentTo(source.Keys);
        foreach (var (relId, sourceFields) in source)
            await Assert.That(saved[relId]).IsEquivalentTo(sourceFields, TUnit.Assertions.Enums.CollectionOrdering.Matching);
    }

    [Test]
    public async Task Save_keeps_the_shared_item_each_record_indexes()
    {
        const string resource = @"TryToLoad\load_pivottable_renamedcolumns_2591.xlsx";
        var source = ReadCacheDefinitions(TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource)), RecordItems);
        var saved = ReadCacheDefinitions(LoadAndSave(resource), RecordItems);

        await Assert.That(source.Values.SelectMany(records => records).Any()).IsTrue();
        await Assert.That(saved.Keys).IsEquivalentTo(source.Keys);
        foreach (var (relId, sourceRecords) in source)
            await Assert.That(saved[relId]).IsEquivalentTo(sourceRecords, TUnit.Assertions.Enums.CollectionOrdering.Matching);
    }

    // One line per cache field: its name, containsNonDate as written, and each shared item with its u flag.
    private static List<string> SharedItemFlags(PivotTableCacheDefinitionPart part)
    {
        return part.PivotCacheDefinition!.CacheFields!.Elements<CacheField>()
            .Select(field => $"{field.Name!.Value} containsNonDate={field.SharedItems?.ContainsNonDate?.Value}: " +
                             string.Join(" ", field.SharedItems?.Elements().Select(item =>
                                 item.LocalName + (item.GetAttributes().Any(a => a.LocalName == "u" && a.Value is "1" or "true") ? "@u" : "")) ?? []))
            .ToList();
    }

    // One line per record: the value each field holds, an index resolved to the shared item it names.
    private static List<string> RecordItems(PivotTableCacheDefinitionPart part)
    {
        var sharedItems = part.PivotCacheDefinition!.CacheFields!.Elements<CacheField>()
            .Select(field => field.SharedItems?.Elements().Select(ItemValue).ToArray() ?? [])
            .ToArray();
        return part.PivotTableCacheRecordsPart?.PivotCacheRecords?.Elements<PivotCacheRecord>()
            .Select(record => string.Join(" ", record.Elements().Select((item, fieldIdx) => item is FieldItem index
                ? sharedItems[fieldIdx][index.Val!.Value]
                : ItemValue(item))))
            .ToList() ?? [];

        static string ItemValue(DocumentFormat.OpenXml.OpenXmlElement item)
            => $"{item.LocalName}:{item.GetAttributes().FirstOrDefault(a => a.LocalName == "v").Value}";
    }

    private static Dictionary<string, List<string>> ReadCacheDefinitions(Stream stream,
        System.Func<PivotTableCacheDefinitionPart, List<string>> read)
    {
        var result = new Dictionary<string, List<string>>();
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        foreach (var pivotCache in workbookPart.Workbook!.PivotCaches?.Elements<PivotCache>() ?? [])
            result.Add(pivotCache.Id!.Value!, read((PivotTableCacheDefinitionPart)workbookPart.GetPartById(pivotCache.Id.Value!)));

        return result;
    }

    [Test]
    public async Task Every_dxf_reference_names_the_dxf_it_had_after_two_saves()
    {
        var source = WithCustomTableStyle(@"TryToLoad\TemplateWithTableSourcePivotTables.xlsx", dxfIds: [4, 5, 365]);
        var expected = ReadDxfReferences(source);
        var first = new MemoryStream();
        var second = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            wb.SaveAs(first);
            wb.SaveAs(second);
        }

        var firstSaved = ReadDxfReferences(first);
        var verbatim = expected.Where(reference => IsCarriedVerbatim(reference.Key, reference.Value))
            .ToDictionary(reference => reference.Key, reference => reference.Value);

        await Assert.That(verbatim.Keys.Select(key => key.Split(' ')[0]).Distinct())
            .IsEquivalentTo(["pivotFormat", "cfRule", "tableStyle", "table"]);
        await Assert.That(expected.Keys.Except(firstSaved.Keys)).IsEmpty();
        await Assert.That(Mismatches(verbatim, firstSaved.Where(r => verbatim.ContainsKey(r.Key)).ToDictionary())).IsEmpty();
        await Assert.That(Mismatches(firstSaved, ReadDxfReferences(second))).IsEmpty();
    }

    // A pivot format that sets a property, and a table column's dataDxfId, are written again from
    // the style model, which renumbers custom number formats and drops readingOrder="0". The save
    // also adds a dataDxfId to a column whose cells share one style.
    private static bool IsCarriedVerbatim(string key, string dxf) => key.Split(' ')[0] switch
    {
        "pivotFormat" => Clears(dxf),
        "table" => !key.EndsWith(" dataDxfId", System.StringComparison.Ordinal),
        _ => true,
    };

    private static List<string> Mismatches(Dictionary<string, string> expected, Dictionary<string, string> actual)
    {
        return expected.Keys.Union(actual.Keys)
            .Where(key => expected.GetValueOrDefault(key) != actual.GetValueOrDefault(key))
            .Select(key => $"{key}: {expected.GetValueOrDefault(key)} -> {actual.GetValueOrDefault(key)}")
            .ToList();
    }

    private static MemoryStream WithCustomTableStyle(string resource, uint[] dxfIds)
    {
        var stream = new MemoryStream();
        using (var resourceStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource)))
            resourceStream.CopyTo(stream);

        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            var style = new TableStyle { Name = "Custom", Pivot = false, Count = (uint)dxfIds.Length };
            TableStyleValues[] types = [TableStyleValues.WholeTable, TableStyleValues.HeaderRow, TableStyleValues.TotalRow];
            for (var i = 0; i < dxfIds.Length; i++)
                style.Append(new TableStyleElement { Type = types[i], FormatId = dxfIds[i] });

            stylesheet.TableStyles!.Append(style);
            stylesheet.TableStyles.Count = 1;
        }

        stream.Position = 0;
        return stream;
    }

    private static Dictionary<string, string> ReadDxfReferences(Stream stream)
    {
        stream.Position = 0;
        var result = new Dictionary<string, string>();
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
        var dxfs = stylesheet.DifferentialFormats?.Elements<DifferentialFormat>().ToArray() ?? [];
        string Dxf(uint id) => WithoutNamespaceDeclarations(dxfs[id].OuterXml);

        foreach (var style in stylesheet.TableStyles?.Elements<TableStyle>() ?? [])
            foreach (var element in style.Elements<TableStyleElement>())
                result.Add($"tableStyle {style.Name!.Value} {element.Type!.InnerText}", Dxf(element.FormatId!.Value));

        foreach (var sheet in workbookPart.Workbook!.Sheets!.Elements<Sheet>())
        {
            if (workbookPart.GetPartById(sheet.Id!.Value!) is not WorksheetPart worksheetPart)
                continue;

            var rules = worksheetPart.Worksheet!.Elements<ConditionalFormatting>()
                .SelectMany(cf => cf.Elements<ConditionalFormattingRule>()
                    .Select((rule, i) => (Key: $"{cf.SequenceOfReferences!.InnerText} {i}", rule.FormatId)));
            foreach (var (key, formatId) in rules)
                result.Add($"cfRule {sheet.Name!.Value}!{key}", Dxf(formatId!.Value));

            foreach (var pivotTablePart in worksheetPart.PivotTableParts)
            {
                var definition = pivotTablePart.PivotTableDefinition!;
                var formats = definition.Formats?.Elements<Format>().ToList() ?? [];
                for (var i = 0; i < formats.Count; i++)
                    if (formats[i].FormatId?.Value is { } id)
                        result.Add($"pivotFormat {sheet.Name!.Value}!{definition.Name!.Value} {i}", Dxf(id));
            }

            foreach (var tablePart in worksheetPart.TableDefinitionParts)
            {
                var table = tablePart.Table!;
                foreach (var element in table.Descendants().Prepend(table))
                    foreach (var attribute in element.GetAttributes().Where(a => a.LocalName.EndsWith("DxfId", System.StringComparison.Ordinal)))
                        result.Add($"table {sheet.Name!.Value}!{table.Name!.Value} {element.LocalName} {element.GetAttribute("name", "").Value} {attribute.LocalName}",
                            Dxf(uint.Parse(attribute.Value!, System.Globalization.CultureInfo.InvariantCulture)));
            }
        }

        return result;
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

    private static List<string?> FormatNumberFormats(PivotTableDefinition definition, DifferentialFormat[] dxfs)
    {
        return definition.Formats?.Elements<Format>()
            .Select(format => format.FormatId?.Value is { } id ? dxfs[id].NumberingFormat?.FormatCode?.Value : null)
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
