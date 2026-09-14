using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Tests.Excel.Worksheets;

/// <summary>
/// #497 and #498, against the <c>chartex-pivotcf-*</c> fixtures, as <see cref="SheetLifecycleFixtureTests"/>
/// does for spec 55's: load the workbook the owner saved in Excel before an edit, make the same edit
/// through XLibur, save, and compare the text of every holder with the workbook Excel saved after it.
/// </summary>
/// <remarks>
/// <para>
/// <c>Other</c> holds a waterfall chart over <c>Data</c>, which is a ChartEx chart, and a pivot table
/// whose source is <c>Data!A1:B4</c>, with a conditional format whose rule is <c>Data!$A$2&gt;0</c>.
/// Excel keeps that rule only in the sheet's <c>x14</c> extension, marked <c>pivot="1"</c>.
/// </para>
/// <para>
/// A ChartEx chart does not hold its references. Each <c>cx:f</c> names a hidden workbook-scoped name,
/// <c>_xlchart.v1.1</c>, and the name holds the reference, <c>Data!$B$1:$B$3</c>. On a rename Excel
/// renames the sheet in the names and leaves the chart part as it was. On a delete it removes every
/// such name, and takes the references out of the chart part: each dimension loses its <c>cx:f</c> and
/// holds an empty <c>cx:lvl</c> for each level it had, and the series loses its <c>cx:tx</c>.
/// </para>
/// <para>
/// Two things are left out of the comparison. On the delete Excel gives the pivot cache's source an
/// <c>r:id</c>, a relationship to the path of the file it was saved from, and XLibur does not, so the
/// source is compared by its sheet and range. And the pivot table's own list of its conditional
/// formats, in the pivot table's <c>x14</c> extension, is not written back by XLibur's pivot table
/// writer on any save, edit or none, which is a defect of its own.
/// </para>
/// </remarks>
public class SheetLifecycleChartExFixtureTests
{
    private const string Folder = @"Other\SheetLifecycle\";
    private const string Before = "chartex-pivotcf-before.xlsx";

    [Test]
    public async Task A_rename_matches_Excel()
    {
        var saved = Read(EditAndSave(wb => wb.Worksheet("Data").Name = "Renamed"));
        var excel = Read(Resource("chartex-pivotcf-rename-after.xlsx"));

        await AssertSameText(saved, excel);
        await Assert.That(saved.ChartNames).Contains("_xlchart.v1.1 = Renamed!$B$1:$B$3 (hidden)");
        await Assert.That(saved.PivotConditionalFormats)
            .Contains("pivot=1 sqref=G2:G4: expression {FE99ACD7-A252-4AFF-8D90-563FB220535F} = Renamed!$A$2>0");
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task A_delete_matches_Excel(bool throughCollection)
    {
        var saved = Read(EditAndSave(wb => Delete(wb, throughCollection)));
        var excel = Read(Resource("chartex-pivotcf-delete-after.xlsx"));

        await AssertSameText(saved, excel);
        await Assert.That(saved.ChartNames).IsEmpty();
        await Assert.That(saved.PivotConditionalFormats)
            .Contains("pivot=1 sqref=G2:G4: expression {FE99ACD7-A252-4AFF-8D90-563FB220535F} = #REF!>0");
    }

    /// <summary>
    /// The chart and the conditional format a delete left survive a reload and a second save as they
    /// are, and the reloaded chart's series have no references and no name.
    /// </summary>
    [Test]
    public async Task A_delete_round_trips()
    {
        using var saved = EditAndSave(wb => Delete(wb, throughCollection: false));
        using var resaved = new MemoryStream();
        using (var reloaded = new XLWorkbook(saved))
        {
            var series = reloaded.Worksheet("Other").Charts.Single().Series.Single();
            await Assert.That(series.Name).IsEqualTo(string.Empty);
            await Assert.That(series.ValueReferences).IsEqualTo(string.Empty);
            await Assert.That(series.CategoryReferences).IsNull();
            reloaded.SaveAs(resaved);
        }

        await AssertSameText(Read(resaved), Read(saved));
    }

    [Test]
    public async Task A_rename_round_trips()
    {
        using var saved = EditAndSave(wb => wb.Worksheet("Data").Name = "Renamed");
        using var resaved = new MemoryStream();
        using (var reloaded = new XLWorkbook(saved))
        {
            var series = reloaded.Worksheet("Other").Charts.Single().Series.Single();
            await Assert.That(series.ValueReferences).IsEqualTo("_xlchart.v1.2");
            reloaded.SaveAs(resaved);
        }

        await AssertSameText(Read(resaved), Read(saved));
    }

    /// <summary>
    /// A delete takes the references out of the loaded chart in memory too, as a reload of the saved
    /// file reads it.
    /// </summary>
    [Test]
    public async Task A_delete_takes_the_references_out_of_the_loaded_chart()
    {
        using var wb = new XLWorkbook(Resource(Before));
        var series = wb.Worksheet("Other").Charts.Single().Series.Single();
        await Assert.That(series.Name).IsEqualTo("S");
        await Assert.That(series.CategoryReferences).IsEqualTo("_xlchart.v1.1");

        wb.Worksheet("Data").Delete();

        await Assert.That(series.Name).IsEqualTo(string.Empty);
        await Assert.That(series.ValueReferences).IsEqualTo(string.Empty);
        await Assert.That(series.CategoryReferences).IsNull();
        await Assert.That(wb.DefinedNames.Select(n => n.Name)).IsEmpty();
    }

    /// <summary>
    /// <see cref="XLWorkbook.Save()"/> reopens the package the previous save wrote, whose chart part a
    /// delete has already patched. A second save leaves it as the first one did.
    /// </summary>
    [Test]
    public async Task A_delete_matches_Excel_after_two_saves()
    {
        using var package = new MemoryStream();
        using (var source = Resource(Before))
            source.CopyTo(package);

        package.Position = 0;
        using (var wb = new XLWorkbook(package))
        {
            wb.Worksheet("Data").Delete();
            wb.Save();
            wb.Save();
        }

        await AssertSameText(Read(package), Read(Resource("chartex-pivotcf-delete-after.xlsx")));
    }

    /// <summary>
    /// <see cref="XLWorkbook.SaveAs(Stream)"/> makes the stream it wrote the one the next save starts
    /// from, so a second <c>SaveAs</c> also patches a part the first one patched.
    /// </summary>
    [Test]
    public async Task A_delete_matches_Excel_after_two_SaveAs()
    {
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(Before)))
        {
            wb.Worksheet("Data").Delete();
            wb.SaveAs(first);
            wb.SaveAs(second);
        }

        var excel = Read(Resource("chartex-pivotcf-delete-after.xlsx"));
        await AssertSameText(Read(first), excel);
        await AssertSameText(Read(second), excel);
    }

    private static async Task AssertSameText(Holders saved, Holders excel)
    {
        await Assert.That(Lines(saved.ChartNames)).IsEqualTo(Lines(excel.ChartNames));
        await Assert.That(Lines(saved.ChartData)).IsEqualTo(Lines(excel.ChartData));
        await Assert.That(Lines(saved.PivotConditionalFormats)).IsEqualTo(Lines(excel.PivotConditionalFormats));
        await Assert.That(saved.PivotSource).IsEqualTo(excel.PivotSource);
    }

    private static string Lines(IEnumerable<string> items) => string.Join(Environment.NewLine, items);

    private static MemoryStream EditAndSave(Action<XLWorkbook> edit)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook(Resource(Before)))
        {
            edit(wb);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return ms;
    }

    private static void Delete(XLWorkbook wb, bool throughCollection)
    {
        if (throughCollection)
            wb.Worksheets.Delete("Data");
        else
            wb.Worksheet("Data").Delete();
    }

    private static Stream Resource(string fileName)
        => TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Folder + fileName));

    /// <summary>
    /// The text of every holder these fixtures exercise, read from a saved package. The chart, the
    /// conditional format and the pivot table are all on <c>Other</c>.
    /// </summary>
    private static Holders Read(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbookPart = document.WorkbookPart!;
        var sheets = workbookPart.Workbook!.Sheets!.Elements<S.Sheet>().ToList();

        var chartNames = (workbookPart.Workbook.DefinedNames?.Elements<S.DefinedName>() ?? [])
            .Where(n => n.Name!.Value!.StartsWith("_xlchart.", StringComparison.OrdinalIgnoreCase))
            .Select(n => $"{n.Name!.Value} = {n.Text}{(n.Hidden?.Value == true ? " (hidden)" : string.Empty)}")
            .Order(StringComparer.Ordinal)
            .ToList();

        var other = (WorksheetPart)workbookPart.GetPartById(sheets.Single(s => s.Name == "Other").Id!.Value!);
        var chartSpaces = other.DrawingsPart?.Parts.Select(p => p.OpenXmlPart).OfType<ExtendedChartPart>()
            .Select(p => p.RootElement!)
            .ToList() ?? [];
        var chartData = chartSpaces.SelectMany(DescribeChartData).ToList();

        var pivotConditionalFormats = other.Worksheet!.Descendants<X14.ConditionalFormatting>()
            .Select(cf =>
            {
                var pivot = cf.GetAttributes().FirstOrDefault(a => a.LocalName == "pivot").Value ?? "0";
                var sqref = cf.Elements<OfficeExcel.ReferenceSequence>().FirstOrDefault()?.Text;
                var rules = cf.Elements<X14.ConditionalFormattingRule>().Select(r =>
                    $"{r.Type?.InnerText} {r.Id?.Value} = {string.Join(" | ", r.Descendants<OfficeExcel.Formula>().Select(f => f.Text))}");
                return $"pivot={pivot} sqref={sqref}: {string.Join("; ", rules)}";
            })
            .Order(StringComparer.Ordinal)
            .ToList();

        var cacheSource = workbookPart.PivotTableCacheDefinitionParts.SingleOrDefault()?
            .PivotCacheDefinition?.CacheSource?.WorksheetSource;
        var pivotSource = cacheSource is null ? null : $"{cacheSource.Sheet?.Value}!{cacheSource.Reference?.Value}";

        return new Holders(chartNames, chartData, pivotConditionalFormats, pivotSource);
    }

    /// <summary>
    /// One line for each data dimension of a ChartEx chart space, and one for each series' name: every
    /// child element, with the text of a <c>cx:f</c> and its direction, and the point count of a
    /// <c>cx:lvl</c>.
    /// </summary>
    private static IEnumerable<string> DescribeChartData(OpenXmlElement chartSpace)
    {
        foreach (var data in chartSpace.Descendants().Where(e => e.LocalName == "data" && e.Parent?.LocalName == "chartData"))
        {
            var id = Attribute(data, "id");
            foreach (var dimension in data.Elements())
                yield return $"data {id} {dimension.LocalName} {Attribute(dimension, "type")}: {Describe(dimension)}";
        }

        foreach (var series in chartSpace.Descendants().Where(e => e.LocalName == "series"))
        {
            var tx = series.Elements().FirstOrDefault(e => e.LocalName == "tx");
            yield return $"series {Attribute(series, "layoutId")}: " +
                         (tx is null ? "no tx" : $"tx {Describe(tx.Elements().Single())}");
        }

        static string Describe(OpenXmlElement parent) => string.Join(" ", parent.Elements().Select(e => e.LocalName switch
        {
            "f" => $"f[{Attribute(e, "dir")}]={e.InnerText}",
            "lvl" => $"lvl[{Attribute(e, "ptCount")}]",
            "v" => $"v={e.InnerText}",
            _ => e.LocalName,
        }));

        static string? Attribute(OpenXmlElement element, string name)
            => element.GetAttributes().FirstOrDefault(a => a.LocalName == name).Value;
    }

    private sealed record Holders(
        List<string> ChartNames,
        List<string> ChartData,
        List<string> PivotConditionalFormats,
        string? PivotSource);
}
