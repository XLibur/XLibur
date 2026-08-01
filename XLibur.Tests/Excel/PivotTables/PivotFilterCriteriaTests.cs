using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// A pivot filter's <c>autoFilter</c> says what the filter actually keeps. It is modelled rather
/// than preserved as a string, so every one of the six <c>filterColumn</c> children — and the
/// parts of the element XLibur cannot act on at all — needs pinning against a round trip.
/// </summary>
/// <remarks>
/// Attributes sitting at their schema default are written as absent, so the assertions here read
/// the effective value rather than requiring the attribute to be present.
/// </remarks>
public class PivotFilterCriteriaTests
{
    /// <summary>
    /// <c>filters</c> — the tick-box list, plus the <c>blank</c> attribute the worksheet model
    /// has no room for.
    /// </summary>
    [Test]
    public async Task ValuesFilter_SurvivesARoundTrip()
    {
        var filters = new Filters { Blank = true };
        filters.Append(new Filter { Val = "Cookies" });
        filters.Append(new Filter { Val = "Cake" });

        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(filters);

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn).Elements<Filters>().Single();

        await Assert.That(written.Blank!.Value).IsTrue();
        await Assert.That(written.Elements<Filter>().Select(f => f.Val!.Value).ToList())
            .IsEquivalentTo(new[] { "Cookies", "Cake" });
    }

    /// <summary>
    /// The other form of <c>filters</c>: date groups rather than values. The schema makes the two
    /// a choice, so a column uses one or the other.
    /// </summary>
    [Test]
    public async Task DateGroupFilter_SurvivesARoundTrip()
    {
        var filters = new Filters { CalendarType = CalendarValues.Hijri };
        filters.Append(new DateGroupItem
        {
            DateTimeGrouping = DateTimeGroupingValues.Day,
            Year = 2024,
            Month = 3,
            Day = 17,
        });

        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(filters);

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn).Elements<Filters>().Single();

        await Assert.That(written.CalendarType!.InnerText).IsEqualTo("hijri");

        var dateGroup = written.Elements<DateGroupItem>().Single();
        await Assert.That(dateGroup.DateTimeGrouping!.InnerText).IsEqualTo("day");
        await Assert.That(dateGroup.Year!.Value).IsEqualTo((ushort)2024);
        await Assert.That(dateGroup.Month!.Value).IsEqualTo((ushort)3);
        await Assert.That(dateGroup.Day!.Value).IsEqualTo((ushort)17);

        // The parts finer than the grouping stay absent: Excel reads a present part as one the
        // filter matches on, so writing hour="0" would narrow the filter.
        await Assert.That(dateGroup.Hour).IsNull();
    }

    /// <summary>
    /// <c>top10</c> taking the bottom by percent, so none of the three booleans sit at their
    /// default and all have to be written.
    /// </summary>
    [Test]
    public async Task Top10Filter_SurvivesARoundTrip()
    {
        var filterColumn = new FilterColumn { ColumnId = 1U };
        filterColumn.Append(new Top10 { Top = false, Percent = true, Val = 25D, FilterValue = 640D });

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn).Elements<Top10>().Single();

        await Assert.That(written.Top!.Value).IsFalse();
        await Assert.That(written.Percent!.Value).IsTrue();
        await Assert.That(written.Val!.Value).IsEqualTo(25D);
        await Assert.That(written.FilterValue!.Value).IsEqualTo(640D);
    }

    /// <summary>
    /// <c>customFilters</c> — two comparisons joined by AND, which is the only shape that uses
    /// both the connector and the operator attributes.
    /// </summary>
    [Test]
    public async Task CustomFilters_SurviveARoundTrip()
    {
        var customFilters = new CustomFilters { And = true };
        customFilters.Append(new CustomFilter { Operator = FilterOperatorValues.GreaterThanOrEqual, Val = "100" });
        customFilters.Append(new CustomFilter { Operator = FilterOperatorValues.LessThan, Val = "900" });

        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(customFilters);

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn).Elements<CustomFilters>().Single();

        await Assert.That(written.And!.Value).IsTrue();

        var criteria = written.Elements<CustomFilter>().ToList();
        await Assert.That(criteria.Count).IsEqualTo(2);
        await Assert.That(criteria[0].Operator!.InnerText).IsEqualTo("greaterThanOrEqual");
        await Assert.That(criteria[0].Val!.Value).IsEqualTo("100");
        await Assert.That(criteria[1].Operator!.InnerText).IsEqualTo("lessThan");
        await Assert.That(criteria[1].Val!.Value).IsEqualTo("900");
    }

    /// <summary>
    /// <c>dynamicFilter</c> with a type outside the two averages XLibur can evaluate. Forcing
    /// those onto <c>XLFilterDynamicType</c> would rewrite "last quarter" as "above average", so
    /// the token is carried rather than mapped.
    /// </summary>
    [Test]
    public async Task DynamicFilter_KeepsATypeXLiburCannotEvaluate()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(new DynamicFilter
        {
            Type = DynamicFilterValues.LastQuarter,
            Val = 45000D,
            MaxVal = 45090D,
        });

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn).Elements<DynamicFilter>().Single();

        await Assert.That(written.Type!.InnerText).IsEqualTo("lastQuarter");
        await Assert.That(written.Val!.Value).IsEqualTo(45000D);
        await Assert.That(written.MaxVal!.Value).IsEqualTo(45090D);
    }

    /// <summary>
    /// <c>colorFilter</c> matching a font colour rather than a fill, so <c>cellColor</c> is off
    /// its default.
    /// </summary>
    [Test]
    public async Task ColorFilter_SurvivesARoundTrip()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(new ColorFilter { FormatId = 3U, CellColor = false });

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn).Elements<ColorFilter>().Single();

        await Assert.That(written.FormatId!.Value).IsEqualTo(3U);
        await Assert.That(written.CellColor!.Value).IsFalse();
    }

    /// <summary>
    /// <c>iconFilter</c> — the child no part of XLibur can evaluate, and the one that survived
    /// before only because the whole sub-tree was opaque.
    /// </summary>
    [Test]
    public async Task IconFilter_SurvivesARoundTrip()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(new IconFilter { IconSet = IconSetValues.ThreeTrafficLights1, IconId = 2U });

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn).Elements<IconFilter>().Single();

        await Assert.That(written.IconSet!.InnerText).IsEqualTo("3TrafficLights1");
        await Assert.That(written.IconId!.Value).IsEqualTo(2U);
    }

    /// <summary>
    /// The button attributes are not modelled anywhere else in XLibur, so a model that only
    /// covered the criteria would drop them.
    /// </summary>
    [Test]
    public async Task FilterColumnButtonAttributes_SurviveARoundTrip()
    {
        var filterColumn = new FilterColumn { ColumnId = 4U, HiddenButton = true, ShowButton = false };
        filterColumn.Append(new Top10 { Val = 5D });

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn);

        await Assert.That(written.ColumnId!.Value).IsEqualTo(4U);
        await Assert.That(written.HiddenButton!.Value).IsTrue();
        await Assert.That(written.ShowButton!.Value).IsFalse();
    }

    /// <summary>
    /// Nothing reads <c>extLst</c>, which is exactly why it has to be carried: it is where a
    /// newer Excel puts whatever this build has never heard of.
    /// </summary>
    [Test]
    public async Task FilterColumnExtensionList_SurvivesARoundTrip()
    {
        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(new Top10 { Val = 5D });
        filterColumn.Append(BuildExtensionList("{FEDCBA98-7654-3210-FEDC-BA9876543210}"));

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn);
        var extension = written.GetFirstChild<ExtensionList>()!.Elements<Extension>().Single();

        await Assert.That(extension.Uri!.Value).IsEqualTo("{FEDCBA98-7654-3210-FEDC-BA9876543210}");
    }

    /// <summary>
    /// <c>autoFilter</c> carries <c>ref</c>, <c>sortState</c> and its own <c>extLst</c> alongside
    /// the columns. Sorting is modelled far more narrowly elsewhere in XLibur than
    /// <c>CT_SortState</c> allows, so it is preserved rather than modelled — but preserved means
    /// it still has to come back.
    /// </summary>
    [Test]
    public async Task AutoFilterReferenceSortStateAndExtensions_SurviveARoundTrip()
    {
        var sortState = new SortState { Reference = "A2:A3", CaseSensitive = true };
        sortState.Append(new SortCondition { Reference = "A2:A3", Descending = true });

        var filterColumn = new FilterColumn { ColumnId = 0U };
        filterColumn.Append(new Top10 { Val = 5D });

        var autoFilter = new AutoFilter { Reference = "A1:A3" };
        autoFilter.Append(filterColumn);
        autoFilter.Append(sortState);
        autoFilter.Append(BuildExtensionList("{01234567-89AB-CDEF-0123-456789ABCDEF}"));

        var written = PivotFilterWorkbook.RoundTripAutoFilter(autoFilter);

        await Assert.That(written.Reference!.Value).IsEqualTo("A1:A3");

        var writtenSort = written.GetFirstChild<SortState>()!;
        await Assert.That(writtenSort.Reference!.Value).IsEqualTo("A2:A3");
        await Assert.That(writtenSort.CaseSensitive!.Value).IsTrue();
        await Assert.That(writtenSort.Elements<SortCondition>().Single().Descending!.Value).IsTrue();

        var extension = written.GetFirstChild<ExtensionList>()!.Elements<Extension>().Single();
        await Assert.That(extension.Uri!.Value).IsEqualTo("{01234567-89AB-CDEF-0123-456789ABCDEF}");
    }

    /// <summary>
    /// A column with no criteria at all is odd but legal, and dropping it would lose the button
    /// attributes it exists to carry.
    /// </summary>
    [Test]
    public async Task FilterColumnWithoutCriteria_SurvivesARoundTrip()
    {
        var filterColumn = new FilterColumn { ColumnId = 2U, ShowButton = false };

        var written = PivotFilterWorkbook.RoundTripFilterColumn(filterColumn);

        await Assert.That(written.ColumnId!.Value).IsEqualTo(2U);
        await Assert.That(written.ShowButton!.Value).IsFalse();
        await Assert.That(written.ChildElements.Count).IsEqualTo(0);
    }

    /// <summary>
    /// Several columns, so the collection is not assumed to hold one.
    /// </summary>
    [Test]
    public async Task SeveralFilterColumns_SurviveARoundTripInOrder()
    {
        var first = new FilterColumn { ColumnId = 0U };
        first.Append(new Top10 { Val = 5D });

        var second = new FilterColumn { ColumnId = 2U };
        second.Append(new IconFilter { IconSet = IconSetValues.FiveArrows, IconId = 4U });

        var autoFilter = new AutoFilter();
        autoFilter.Append(first);
        autoFilter.Append(second);

        var written = PivotFilterWorkbook.RoundTripAutoFilter(autoFilter).Elements<FilterColumn>().ToList();

        await Assert.That(written.Count).IsEqualTo(2);
        await Assert.That(written[0].ColumnId!.Value).IsEqualTo(0U);
        await Assert.That(written[0].Elements<Top10>().Single().Val!.Value).IsEqualTo(5D);
        await Assert.That(written[1].ColumnId!.Value).IsEqualTo(2U);
        await Assert.That(written[1].Elements<IconFilter>().Single().IconSet!.InnerText).IsEqualTo("5Arrows");
    }

    /// <summary>
    /// A workbook carrying one filter of every kind validates against the schema after a round
    /// trip. Excel repairs a file it cannot parse rather than reporting where it broke, so this
    /// is the check that catches a criteria writer emitting the right values in the wrong shape.
    /// </summary>
    [Test]
    public async Task EveryFilterKindTogether_ProducesASchemaValidPackage()
    {
        var kinds = new OpenXmlElement[]
        {
            BuildValuesFilter(),
            BuildDateGroupFilter(),
            new Top10 { Top = false, Percent = true, Val = 25D, FilterValue = 640D },
            BuildCustomFilters(),
            new DynamicFilter { Type = DynamicFilterValues.LastQuarter, Val = 45000D },
            new ColorFilter { FormatId = 0U, CellColor = false },
            new IconFilter { IconSet = IconSetValues.ThreeTrafficLights1, IconId = 2U },
        };

        var filters = kinds.Select((kind, index) =>
        {
            var filterColumn = new FilterColumn { ColumnId = 0U };
            filterColumn.Append(kind);

            var autoFilter = new AutoFilter { Reference = "A1:A3" };
            autoFilter.Append(filterColumn);

            var pivotFilter = new PivotFilter
            {
                Field = 0U,
                Id = (uint)index + 1,
                Type = PivotFilterValues.CaptionEqual,
            };
            pivotFilter.Append(autoFilter);
            return pivotFilter;
        }).ToArray();

        using var input = new MemoryStream();
        PivotFilterWorkbook.Create(input, filters);
        input.Position = 0;

        using var wb = new XLWorkbook(input);
        using var output = new MemoryStream();
        wb.SaveAs(output);

        // All of them came back, and the package the SDK validator sees is well-formed.
        await Assert.That(PivotFilterWorkbook.ReadFilters(output).Count).IsEqualTo(kinds.Length);

        output.Position = 0;
        using var doc = SpreadsheetDocument.Open(output, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2010)
            .Validate(doc)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToList();

        await Assert.That(errors).IsEmpty();
    }

    private static Filters BuildValuesFilter()
    {
        var filters = new Filters { Blank = true };
        filters.Append(new Filter { Val = "Cookies" });
        return filters;
    }

    private static Filters BuildDateGroupFilter()
    {
        var filters = new Filters();
        filters.Append(new DateGroupItem
        {
            DateTimeGrouping = DateTimeGroupingValues.Month,
            Year = 2024,
            Month = 3,
        });

        return filters;
    }

    private static CustomFilters BuildCustomFilters()
    {
        var customFilters = new CustomFilters { And = true };
        customFilters.Append(new CustomFilter { Operator = FilterOperatorValues.GreaterThanOrEqual, Val = "100" });
        customFilters.Append(new CustomFilter { Operator = FilterOperatorValues.LessThan, Val = "900" });
        return customFilters;
    }

    private static ExtensionList BuildExtensionList(string uri)
    {
        var extension = new Extension { Uri = uri };
        extension.AppendChild(new OpenXmlUnknownElement("x14", "filter", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"));

        var extensionList = new ExtensionList();
        extensionList.Append(extension);
        return extensionList;
    }
}
