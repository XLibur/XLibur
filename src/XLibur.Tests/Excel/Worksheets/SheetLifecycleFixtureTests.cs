using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// Spec 55 design §5. Each test loads a workbook the owner saved in Excel before an edit, makes the
/// same edit through XLibur, saves, and compares the text of every holder with the workbook Excel
/// saved after the edit. The fixtures are in <c>Resource/Other/SheetLifecycle</c>, and spec 55's
/// Results record what each one holds.
/// </summary>
/// <remarks>
/// Two differences from Excel are deliberate, both calls recorded in spec 55's Results, so the
/// comparison allows for them:
/// <list type="bullet">
/// <item>On a delete, Excel moves a chart series' name and categories into <c>c15:filtered*</c>
/// extensions, each <c>#REF!</c>. XLibur rewrites each <c>c:f</c> to <c>#REF!</c> where it stands. The
/// references are compared as text, wherever they stand.</item>
/// <item>Excel shrinks a spilled array formula's range when its result becomes a single error. That
/// is recalculation, not a rewrite, so only formula text is compared.</item>
/// </list>
/// </remarks>
public class SheetLifecycleFixtureTests
{
    private const string Folder = @"Other\SheetLifecycle\";
    private const string Workbook = "(workbook)";

    [Test]
    public async Task A_rename_matches_Excel()
    {
        var saved = EditAndSave("rename-before.xlsx", wb => wb.Worksheet("Data").Name = "Renamed");
        var excel = Read(Resource("rename-after.xlsx"));

        await AssertSameText(saved, excel, excelChartMovesReferences: false);
    }

    /// <summary>
    /// Every holder matches Excel, the names scoped to the deleted sheet included: <c>Local</c>, which
    /// <c>Q</c> refers to, is kept at workbook scope as <c>#REF!</c>, and <c>Q</c> becomes
    /// <c>[0]!Local</c>. <c>Data</c>'s print area goes with <c>Data</c>.
    /// </summary>
    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task Deleting_two_sheets_matches_Excel(bool throughCollection)
    {
        var saved = EditAndSave("delete-before.xlsx", wb =>
        {
            Delete(wb, "First", throughCollection);
            Delete(wb, "Data", throughCollection);
        });
        var excel = Read(Resource("delete-after.xlsx"));

        await AssertSameText(saved, excel, excelChartMovesReferences: true);
    }

    /// <summary>
    /// <c>delete-before</c> holds its sheets out of <c>sheetId</c> order. <c>Data</c>'s print area has
    /// <c>localSheetId="1"</c>, the second tab, and loaded onto <c>Other</c>, the sheet whose
    /// <c>sheetId</c> is 2 (D75).
    /// </summary>
    [Test]
    public async Task A_print_area_loads_onto_the_sheet_at_its_position()
    {
        using var wb = new XLWorkbook(Resource("delete-before.xlsx"));

        var printAreas = wb.Worksheets
            .Select(w => (w.Name, ((XLPrintAreas)w.PageSetup.PrintAreas).FormulaReference))
            .Where(p => p.FormulaReference is not null)
            .ToList();
        await Assert.That(printAreas).IsEquivalentTo([("Data", (string?)"OFFSET(Data!$A$1,0,0,4,2)")]);
    }

    [Test]
    public async Task Deleting_a_sheet_with_a_broken_name_matches_Excel()
    {
        var saved = EditAndSave("refdelete-before.xlsx", wb => wb.Worksheet("Data").Delete());
        var excel = Read(Resource("refdelete-after.xlsx"));

        await Assert.That(Lines(saved.Names)).IsEqualTo(Lines(excel.Names));
        await Assert.That(saved.Names).Contains($"{Workbook}|Broken = #REF!");
    }

    [Test]
    public async Task Deleting_a_sheet_with_names_scoped_to_it_matches_Excel()
    {
        var saved = EditAndSave("scoped-delete-before.xlsx", wb => wb.Worksheet("Data").Delete());
        var excel = Read(Resource("scoped-delete-after.xlsx"));

        await Assert.That(Lines(saved.Names)).IsEqualTo(Lines(excel.Names));
        await Assert.That(Lines(saved.CellFormulas)).IsEqualTo(Lines(excel.CellFormulas));
        await Assert.That(saved.CellFormulas).Contains("A1 = [0]!Used");
    }

    /// <summary>
    /// Excel keeps a pivot cache whose source was on the deleted sheet, with its records, and the pivot
    /// table on <c>Other</c> that uses it. XLibur's save used to delete the cache part. The delete adds
    /// no <c>OpenXmlValidator</c> error: the package has exactly the errors a save of the untouched
    /// workbook has.
    /// </summary>
    /// <remarks>
    /// Neither package passes outright. Excel's own <c>delete-before.xlsx</c> fails validation, with a
    /// <c>pageSetup</c> whose dpi is 0 and an extension inside <c>c:chart</c>, and XLibur writes the
    /// chart part back as it loaded it.
    /// </remarks>
    [Test]
    public async Task Deleting_the_sheet_a_pivot_cache_reads_keeps_the_cache_and_the_pivot_table()
    {
        using var untouched = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("delete-before.xlsx")))
            wb.SaveAs(untouched);

        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("delete-before.xlsx")))
        {
            wb.Worksheet("Data").Delete();
            wb.SaveAs(ms);
        }

        var saved = Read(ms);
        await Assert.That(saved.PivotSource).IsEqualTo("Data!A1:B4");
        await Assert.That(saved.PivotTablesOnOther).IsEqualTo(1);
        await Assert.That(saved.PivotCacheRecords).IsGreaterThan(0);
        await Assert.That(Lines(ValidationErrors(ms))).IsEqualTo(Lines(ValidationErrors(untouched)));

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet("Other").PivotTables.Count()).IsEqualTo(1);
    }

    /// <summary>
    /// A pivot cache's source stopped resolving when its sheet was renamed (D66). It names the sheet
    /// by its new name now, so it resolves again.
    /// </summary>
    [Test]
    public async Task After_a_rename_a_pivot_cache_source_resolves_again()
    {
        using var wb = new XLWorkbook(Resource("rename-before.xlsx"));

        wb.Worksheet("Data").Name = "Renamed";

        var source = wb.PivotCaches.Single().SourceRange;
        await Assert.That(source).IsNotNull();
        await Assert.That(source!.RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Renamed!A1:B4");
    }

    /// <summary>
    /// A loaded chart is patched in place, and a rewritten series reference reaches the part only
    /// because the rewrite marks it as assigned. The series name keeps its cached text, which is the
    /// name XLibur reads back.
    /// </summary>
    [Test]
    public async Task A_loaded_charts_series_keeps_its_name_through_a_rename()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("rename-before.xlsx")))
        {
            wb.Worksheet("Data").Name = "Renamed";
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        var series = (XLChartSeries)reloaded.Worksheet("Other").Charts.Single().Series.Single();
        await Assert.That(series.Name).IsEqualTo("S");
        await Assert.That(series.NameReference).IsEqualTo("Renamed!$D$1");
        await Assert.That(series.ValueReferences).IsEqualTo("Renamed!$A$2:$A$4");
    }

    /// <summary>
    /// A rename rewrites a loaded chart's references and keeps the cached values that go with them,
    /// as Excel does: its <c>rename-after.xlsx</c> keeps the <c>c:numCache</c> and both
    /// <c>c:strCache</c> elements. A viewer that draws from the cache would otherwise show an empty
    /// chart.
    /// </summary>
    [Test]
    public async Task A_rename_keeps_a_loaded_charts_cached_values()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("rename-before.xlsx")))
        {
            wb.Worksheet("Data").Name = "Renamed";
            wb.SaveAs(ms);
        }

        await Assert.That(ChartCaches(ms)).IsEqualTo(ChartCaches(Resource("rename-after.xlsx")));
        await Assert.That(ChartCaches(ms)).IsEqualTo((1, 2));
    }

    /// <summary>
    /// A delete keeps the caches too, while each <c>c:f</c> reads <c>#REF!</c>. Excel's
    /// <c>delete-after.xlsx</c> keeps the values' <c>c:numCache</c>; the name's and the categories'
    /// caches follow the rule for a rename, since XLibur does not move them into <c>c15:filtered*</c>.
    /// </summary>
    [Test]
    public async Task A_delete_keeps_a_loaded_charts_cached_values()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("delete-before.xlsx")))
        {
            wb.Worksheet("Data").Delete();
            wb.SaveAs(ms);
        }

        await Assert.That(ChartCaches(ms)).IsEqualTo((1, 2));
    }

    /// <summary>
    /// A caller who re-points a loaded series still drops the cache that described the old range, so
    /// that a chart re-pointed at other cells does not open showing the old values. Re-pointing wins
    /// over a rename's rewrite of the same reference.
    /// </summary>
    [Test]
    public async Task Re_pointing_a_loaded_series_still_drops_its_cached_values()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource("rename-before.xlsx")))
        {
            wb.Worksheet("Data").Name = "Renamed";
            wb.Worksheet("Other").Charts.Single().Series.Single().ValueReferences = "Renamed!$A$2:$A$3";
            wb.SaveAs(ms);
        }

        await Assert.That(ChartCaches(ms)).IsEqualTo((0, 2));
    }

    /// <summary>
    /// The number of <c>c:numCache</c> and <c>c:strCache</c> elements in a package's chart parts.
    /// </summary>
    private static (int NumberCaches, int StringCaches) ChartCaches(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var chartSpaces = document.WorkbookPart!.WorksheetParts
            .SelectMany(w => w.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())
            .Select(p => p.ChartSpace!)
            .ToList();
        return (chartSpaces.Sum(c => c.Descendants<C.NumberingCache>().Count()),
            chartSpaces.Sum(c => c.Descendants<C.StringCache>().Count()));
    }

    /// <summary>
    /// Compares each holder as one line per item, so that a failure shows both sides in full.
    /// </summary>
    private static async Task AssertSameText(Holders saved, Holders excel, bool excelChartMovesReferences)
    {
        await Assert.That(Lines(saved.Names)).IsEqualTo(Lines(excel.Names));
        await Assert.That(Lines(saved.CellFormulas)).IsEqualTo(Lines(excel.CellFormulas));
        await Assert.That(Lines(saved.ConditionalFormatFormulas)).IsEqualTo(Lines(excel.ConditionalFormatFormulas));
        await Assert.That(saved.PivotSource).IsEqualTo(excel.PivotSource);
        await Assert.That(saved.PivotTablesOnOther).IsEqualTo(excel.PivotTablesOnOther);
        await Assert.That(Lines(saved.Hyperlinks)).IsEqualTo(Lines(excel.Hyperlinks));

        // XLibur never writes the c15:filtered* form, so a reference Excel moved there is compared
        // with the c:f XLibur left where it was.
        await Assert.That(saved.FilteredChartReferences).IsEmpty();
        var excelChartReferences = excelChartMovesReferences
            ? excel.ChartReferences.Concat(excel.FilteredChartReferences).Order(StringComparer.Ordinal).ToList()
            : excel.ChartReferences;
        var savedChartReferences = excelChartMovesReferences
            ? saved.ChartReferences.Order(StringComparer.Ordinal).ToList()
            : saved.ChartReferences;
        await Assert.That(Lines(savedChartReferences)).IsEqualTo(Lines(excelChartReferences));
    }

    private static string Lines(IEnumerable<string> items) => string.Join(Environment.NewLine, items);

    private static Holders EditAndSave(string before, Action<XLWorkbook> edit)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(before)))
        {
            edit(wb);
            wb.SaveAs(ms);
        }

        return Read(ms);
    }

    private static void Delete(XLWorkbook wb, string sheetName, bool throughCollection)
    {
        if (throughCollection)
            wb.Worksheets.Delete(sheetName);
        else
            wb.Worksheet(sheetName).Delete();
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));

    /// <summary>
    /// Every <c>OpenXmlValidator</c> error of a saved package, one line each.
    /// </summary>
    private static List<string> ValidationErrors(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}")
            .Order(StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>
    /// The text of every holder the fixtures exercise, read from a saved package. Names are keyed by
    /// their scope, which is a sheet name or <see cref="Workbook"/>; the other holders are all on the
    /// sheet <c>Other</c>.
    /// </summary>
    private static Holders Read(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        var sheets = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().ToList();

        var names = (workbookPart.Workbook.DefinedNames?.Elements<S.DefinedName>() ?? [])
            .Select(n =>
            {
                var scope = n.LocalSheetId?.Value is { } id ? sheets[(int)id].Name!.Value! : Workbook;
                return $"{scope}|{n.Name!.Value} = {n.Text}";
            })
            .Order(StringComparer.Ordinal)
            .ToList();

        var other = (WorksheetPart)workbookPart.GetPartById(sheets.Single(s => s.Name == "Other").Id!.Value!);
        var worksheet = other.Worksheet!;

        var cellFormulas = worksheet.Descendants<S.Cell>()
            .Where(c => c.CellFormula is not null)
            .Select(c => $"{c.CellReference!.Value} = {c.CellFormula!.Text}")
            .Order(StringComparer.Ordinal)
            .ToList();

        var conditionalFormatFormulas = worksheet.Descendants<X14.ConditionalFormattingRule>()
            .Select(r => $"{r.Type?.InnerText} = {string.Join(" | ", r.Descendants<OfficeExcel.Formula>().Select(f => f.Text))}")
            .Order(StringComparer.Ordinal)
            .ToList();

        var chartSpaces = other.DrawingsPart?.ChartParts.Select(p => p.ChartSpace!).ToList() ?? [];
        var chartReferences = chartSpaces
            .SelectMany(c => c.Descendants<C.Formula>())
            .Select(f => f.Text)
            .ToList();
        var filteredChartReferences = chartSpaces
            .SelectMany(c => c.Descendants())
            .Where(e => e.LocalName == "sqref" && e.Prefix == "c15")
            .Select(e => e.InnerText)
            .ToList();

        var cacheDefinition = workbookPart.PivotTableCacheDefinitionParts.SingleOrDefault();
        var cacheSource = cacheDefinition?.PivotCacheDefinition?.CacheSource?.WorksheetSource;
        var pivotSource = cacheSource is null ? null : $"{cacheSource.Sheet?.Value}!{cacheSource.Reference?.Value}";

        // The rows, not the optional count attribute, which XLibur's records writer leaves out.
        var pivotCacheRecords = cacheDefinition?.PivotTableCacheRecordsPart?.PivotCacheRecords?
            .Elements<S.PivotCacheRecord>().Count() ?? 0;

        var hyperlinks = worksheet.Descendants<S.Hyperlink>()
            .Select(h => $"{h.Reference!.Value} = {h.Location?.Value}")
            .Order(StringComparer.Ordinal)
            .ToList();

        return new Holders(names, cellFormulas, conditionalFormatFormulas, chartReferences, filteredChartReferences,
            pivotSource, other.PivotTableParts.Count(), pivotCacheRecords, hyperlinks);
    }

    private sealed record Holders(
        List<string> Names,
        List<string> CellFormulas,
        List<string> ConditionalFormatFormulas,
        List<string> ChartReferences,
        List<string> FilteredChartReferences,
        string? PivotSource,
        int PivotTablesOnOther,
        int PivotCacheRecords,
        List<string> Hyperlinks);
}
