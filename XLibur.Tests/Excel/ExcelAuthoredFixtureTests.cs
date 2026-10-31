using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel;

/// <summary>
/// Files authored in Excel itself (over COM, with Excel 365): inputs XLibur's own writer cannot
/// produce, so a reader written against that writer cannot be tested with them (specs 41 and 48).
/// </summary>
/// <remarks>
/// Provenance: <c>Resource/Other/ExcelAuthored</c>. <c>ConditionalFormats.xlsx</c> holds three
/// rules over the values -5, -2, 0, 3, 6: a solid data bar whose negative fill is the same as the
/// positive (A), a bordered data bar with an axis in the middle and a different negative border (B)
/// and a 3Flags icon set (C). The pivot files pivot a 24-row sheet "Data" with columns Date (every
/// 37 days from 2025-01-15), Amount, Cat and Mixed (alternating dates and numbers):
/// <c>PivotGroupedByMonths</c> groups Date by months and years, <c>PivotWithErrorValues</c> pivots
/// Cat, which holds #N/A and #DIV/0!, and <c>PivotMixedDatesAndNumbers</c> pivots Mixed.
/// <c>SparklineThemeAxisColor.xlsx</c> has sheet "S" with the values in A1:C3 and one sparkline
/// group in E1:E3 whose horizontal axis is shown and coloured Accent1 darker by 25 percent.
/// </remarks>
public class ExcelAuthoredFixtureTests
{
    private static MemoryStream Open(string name)
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\ExcelAuthored\" + name));
        var copy = new MemoryStream();
        source.CopyTo(copy);
        copy.Position = 0;
        return copy;
    }

    private static MemoryStream Resave(string name)
    {
        using var source = Open(name);
        using var wb = new XLWorkbook(source);
        var saved = new MemoryStream();
        wb.SaveAs(saved);
        saved.Position = 0;
        return saved;
    }

    private static XLConditionalFormat[] DataBarsAndIconSet(XLWorkbook wb)
        => wb.Worksheet("CF").ConditionalFormats.Cast<XLConditionalFormat>()
            .OrderBy(cf => cf.Ranges.First().RangeAddress.FirstAddress.ColumnNumber).ToArray();

    #region Conditional formats

    [Test]
    public async Task DataBarWithNegativeFillSameAsPositiveLoadsAndKeepsItsAttributes()
    {
        using var source = Open("ConditionalFormats.xlsx");
        using var wb = new XLWorkbook(source);

        var bar = DataBarsAndIconSet(wb)[0];
        await Assert.That(bar.ConditionalFormatType).IsEqualTo(XLConditionalFormatType.DataBar);
        await Assert.That(bar.NegativeBarColorSameAsPositive).IsTrue();
        await Assert.That(bar.Gradient).IsFalse();
        await Assert.That(bar.BarAxisPosition).IsEqualTo(XLDataBarAxisPosition.None);
        await Assert.That(bar.BarMinLength).IsEqualTo(0u);
        await Assert.That(bar.BarMaxLength).IsEqualTo(100u);
    }

    [Test]
    public async Task BorderedDataBarKeepsItsBorderColoursAndFlags()
    {
        using var source = Open("ConditionalFormats.xlsx");
        using var wb = new XLWorkbook(source);

        var bar = DataBarsAndIconSet(wb)[1];
        await Assert.That(bar.BarBorder).IsTrue();
        await Assert.That(bar.BarBorderColor).IsEqualTo(XLColor.FromArgb(0xFF, 0x00, 0xFF, 0x00));
        await Assert.That(bar.NegativeBarBorderColor).IsEqualTo(XLColor.FromArgb(0xFF, 0x00, 0x00, 0xFF));
        await Assert.That(bar.NegativeBarBorderColorSameAsPositive).IsFalse();
        await Assert.That(bar.BarAxisPosition).IsEqualTo(XLDataBarAxisPosition.Middle);
    }

    [Test]
    public async Task IconSetLoads()
    {
        using var source = Open("ConditionalFormats.xlsx");
        using var wb = new XLWorkbook(source);

        var icons = DataBarsAndIconSet(wb)[2];
        await Assert.That(icons.IconSetStyle).IsEqualTo(XLIconSetStyle.ThreeFlags);
        await Assert.That(icons.ContentTypes.Count).IsEqualTo(3);
    }

    [Test]
    public async Task ConditionalFormatsAreEqualAfterALoadSaveLoad()
    {
        using var source = Open("ConditionalFormats.xlsx");
        using var first = new XLWorkbook(source);
        using var saved = Resave("ConditionalFormats.xlsx");
        using var second = new XLWorkbook(saved);

        var before = DataBarsAndIconSet(first);
        var after = DataBarsAndIconSet(second);
        await Assert.That(after.Length).IsEqualTo(before.Length);
        for (var i = 0; i < before.Length; i++)
        {
            // The ranges belong to a workbook each, so they are compared by address.
            await Assert.That(XLConditionalFormat.NoRangeComparer.Equals(before[i], after[i])).IsTrue();
            await Assert.That(after[i].Ranges.Select(r => r.RangeAddress.ToStringRelative()))
                .IsEquivalentTo(before[i].Ranges.Select(r => r.RangeAddress.ToStringRelative()));
        }
    }

    #endregion

    #region Sparklines

    [Test]
    public async Task SparklineGroupWithAThemeColouredAxisLoadsAndSurvivesASave()
    {
        using var source = Open("SparklineThemeAxisColor.xlsx");
        using var wb = new XLWorkbook(source);

        var group = wb.Worksheet("S").SparklineGroups.Single();
        await Assert.That(group.Count()).IsEqualTo(3);
        await Assert.That(group.HorizontalAxis.IsVisible).IsTrue();
        await Assert.That(group.HorizontalAxis.Color.ColorType).IsEqualTo(XLColorType.Theme);
        await Assert.That(group.HorizontalAxis.Color.ThemeColor).IsEqualTo(XLThemeColor.Accent1);

        using var saved = Resave("SparklineThemeAxisColor.xlsx");
        await Assert.That(saved.Sheet1Xml()).Contains("<x14:colorAxis theme=\"4\" tint=\"-0.249977111117893\"");
        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var again = reloaded.Worksheet("S").SparklineGroups.Single();
        await Assert.That(again.Count()).IsEqualTo(3);
        await Assert.That(again.HorizontalAxis.Color).IsEqualTo(group.HorizontalAxis.Color);
    }

    #endregion

    #region Pivot cache

    [Test]
    public async Task PivotOverErrorValuesKeepsTheErrorsAsErrorItems()
    {
        using var saved = Resave("PivotWithErrorValues.xlsx");

        var definition = saved.ReadPartUnder("xl/pivotCache/pivotCacheDefinition");
        await Assert.That(definition).Contains("<x:e v=\"#N/A\"");
        await Assert.That(definition).Contains("<x:e v=\"#DIV/0!\"");

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Pivot").PivotTables.Count()).IsEqualTo(1);
    }

    [Test]
    public async Task RecordsPartStatesTheNumberOfRecordsExcelWrote()
    {
        foreach (var name in new[] { "PivotWithErrorValues.xlsx", "PivotMixedDatesAndNumbers.xlsx", "PivotGroupedByMonths.xlsx" })
        {
            using var saved = Resave(name);
            await Assert.That(saved.ReadPartUnder("xl/pivotCache/pivotCacheRecords")).Contains(" count=\"24\"");
        }
    }

    /// <summary>
    /// For a field that mixes dates and numbers Excel writes the dates' range as <c>minDate</c> and
    /// <c>maxDate</c> with the numbers read as serial dates and no <c>minValue</c>, <c>maxValue</c> or
    /// <c>containsNumber</c> (checked in the fixture): the writer folding the numbers into the date
    /// range is what Excel does, so there is no numeric range to keep.
    /// </summary>
    [Test]
    public async Task FieldMixingDatesAndNumbersIsWrittenAsExcelWritesIt()
    {
        using var saved = Resave("PivotMixedDatesAndNumbers.xlsx");

        var definition = saved.ReadPartUnder("xl/pivotCache/pivotCacheDefinition");
        var mixed = Regex.Match(definition, "<x:cacheField name=\"Mixed\".*?</x:cacheField>", RegexOptions.Singleline).Value;
        await Assert.That(mixed).Contains("containsDate=\"1\"");
        await Assert.That(mixed).Contains("containsMixedTypes=\"1\"");
        await Assert.That(mixed).Contains("minDate=");
        await Assert.That(mixed).Contains("maxDate=");
        await Assert.That(mixed).DoesNotContain("containsNumber");
        await Assert.That(mixed).DoesNotContain("minValue");
        await Assert.That(mixed).DoesNotContain("maxValue");
        await Assert.That(mixed).Contains("<x:n v=\"4.5\"");
        await Assert.That(mixed).Contains("<x:d v=\"2025-01-15T00:00:00\"");

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Pivot").PivotTables.Count()).IsEqualTo(1);
    }

    [Test]
    [Arguments("PivotGroupedByMonths.xlsx")]
    [Arguments("PivotWithErrorValues.xlsx")]
    [Arguments("PivotMixedDatesAndNumbers.xlsx")]
    public async Task EveryRecordHasOneColumnPerDatabaseFieldDeclared(string name)
    {
        using var saved = Resave(name);
        var definition = saved.ReadPartUnder("xl/pivotCache/pivotCacheDefinition");
        var records = saved.ReadPartUnder("xl/pivotCache/pivotCacheRecords");

        var fieldCount = Regex.Matches(definition, "<x:cacheField ").Count;
        var nonDatabase = Regex.Matches(definition, "<x:cacheField [^>]*databaseField=\"0\"").Count;
        var firstRecord = Regex.Match(records, "<r>(.*?)</r>", RegexOptions.Singleline).Groups[1].Value;
        var columns = Regex.Matches(firstRecord, "<[a-z] ").Count;
        await Assert.That(columns).IsEqualTo(fieldCount - nonDatabase);
    }

    /// <summary>
    /// The set of non-database fields is filled while loading, so it may hold an index outside the
    /// fields or the same index twice. Neither may disturb which fields have a record column.
    /// </summary>
    [Test]
    public async Task DatabaseFieldIndexesAgreeWithTheirCountForAMalformedSet()
    {
        using var source = Open("PivotWithErrorValues.xlsx");
        using var wb = new XLWorkbook(source);
        var cache = (XLPivotCache)wb.PivotCaches.Single();
        var fieldCount = cache.FieldCount;

        cache.SetNonDatabaseField(fieldCount + 5);
        cache.SetNonDatabaseField(1);
        cache.SetNonDatabaseField(1);

        var indexes = cache.GetDatabaseFieldIndexes();
        await Assert.That(indexes).IsEquivalentTo(Enumerable.Range(0, fieldCount).Where(i => i != 1).ToArray());
        await Assert.That(cache.DatabaseFieldCount).IsEqualTo(indexes.Length);
    }

    #endregion
}
