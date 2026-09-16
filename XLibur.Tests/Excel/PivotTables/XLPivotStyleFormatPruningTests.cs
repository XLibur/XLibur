using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// #577: a pivot table's <c>formats</c> can hold a reference on the 'data' field, and such a
/// reference names values by their <em>position</em> in <c>dataFields</c>. Taking a value out must
/// leave no format naming it, and must renumber the formats naming the values after it.
/// </summary>
/// <remarks>
/// <para>
/// The fixture is <c>Other\Lion\PivotTables\PivotWithStyles.xlsx</c>, written by Excel
/// (<c>docProps/app.xml</c> says so). Its pivot table has three values and twelve formats, five of
/// which carry a reference on the 'data' field: three naming one position each (1, 0 and 2) and two
/// naming all three (0, 1, 2).
/// </para>
/// <para>
/// That the items are positions in <c>dataFields</c> and not pivot field indices is established by
/// two other Excel-written fixtures rather than by this one, whose three values happen to sit on
/// pivot fields 0, 1 and 2. In <c>TryToLoad\TemplateWithTableSourcePivotTables.xlsx</c> the values of
/// <c>pivotTable6</c> are on pivot fields 16 and 18 while its format reference names items 0 and 1,
/// and <c>pivotTable3</c> has two of its four values on the one pivot field 6.
/// </para>
/// </remarks>
internal class XLPivotStyleFormatPruningTests
{
    private const string FixturePath = @"Other\Lion\PivotTables\PivotWithStyles.xlsx";

    /// <summary>
    /// What a reference's <c>field</c> holds when it is on the 'data' field rather than a pivot
    /// field: -2 unsigned, written as 4294967294.
    /// </summary>
    private const uint DataFieldReferenceIndex = unchecked((uint)-2);

    /// <summary>
    /// The data-field references the fixture holds as it is loaded, one entry per reference.
    /// </summary>
    private static readonly string[] AsLoaded = ["1", "0", "2", "0,1,2", "0,1,2"];

    [Test]
    [Property("Description", "#577: a plain load and save must leave every format exactly as it was")]
    public async Task Loading_and_saving_keeps_every_style_format_naming_a_value()
    {
        using var saved = SaveFixture(_ => { });

        var xml = PivotTableXml(saved);

        await Assert.That(FormatCount(xml)).IsEqualTo(12);
        await Assert.That(DataFieldReferences(xml)).IsEquivalentTo(AsLoaded)
            .Because("pruning belongs to the removal paths, not to the load path");
    }

    [Test]
    [Property("Description", "#577: removing the first of several values shifted the rest onto the wrong value")]
    public async Task Removing_a_value_renumbers_the_formats_naming_the_values_after_it()
    {
        // Position 0 goes, so positions 1 and 2 become 0 and 1. The format that named only
        // position 0 is left naming nothing and goes with it; the other four follow their values.
        using var saved = SaveFixture(pt => pt.Values.Remove(FirstValueName(pt)));

        var xml = PivotTableXml(saved);

        await Assert.That(DataFieldReferences(xml)).IsEquivalentTo(["0", "1", "0,1", "0,1"])
            .Because("a surviving reference must follow its own value down a position");
        await Assert.That(FormatCount(xml)).IsEqualTo(11)
            .Because("only the one format that named the removed value alone is dropped");
    }

    [Test]
    [Property("Description", "#577: removing the last value must renumber nothing")]
    public async Task Removing_the_last_value_leaves_the_earlier_positions_alone()
    {
        using var saved = SaveFixture(pt => pt.Values.Remove(pt.Values.Last().CustomName));

        var xml = PivotTableXml(saved);

        await Assert.That(DataFieldReferences(xml)).IsEquivalentTo(["1", "0", "0,1", "0,1"])
            .Because("nothing sits after the removed position, so only position 2 itself goes");
        await Assert.That(FormatCount(xml)).IsEqualTo(11);
    }

    [Test]
    [Property("Description", "#577: after a clear, no format may name a data field the file does not have")]
    public async Task Clearing_the_values_drops_every_format_naming_a_value()
    {
        using var saved = SaveFixture(pt => pt.Values.Clear());

        var xml = PivotTableXml(saved);

        await Assert.That(xml).DoesNotContain("<dataFields");
        await Assert.That(DataFieldReferences(xml)).IsEmpty();
        await Assert.That(FormatCount(xml)).IsEqualTo(7)
            .Because("the seven formats that never named a value must be left alone");
    }

    [Test]
    [Property("Description", "#577: a value removed from a table whose formats survive must still reload")]
    public async Task A_renumbered_format_reloads_as_it_was_saved()
    {
        using var saved = SaveFixture(pt => pt.Values.Remove(FirstValueName(pt)));

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var pt = (XLPivotTable)reloaded.Worksheets.SelectMany(ws => ws.PivotTables).First();

        await Assert.That(pt.Values.Count()).IsEqualTo(2);
        await Assert.That(pt.Formats.Count).IsEqualTo(11);

        // Every position a surviving format names must be one the table actually has.
        var named = pt.Formats
            .SelectMany(format => format.PivotArea.References)
            .Where(reference => reference.Field == DataFieldReferenceIndex)
            .SelectMany(reference => reference.FieldItems)
            .ToList();

        await Assert.That(named).IsNotEmpty();
        await Assert.That(named.All(position => position < pt.Values.Count())).IsTrue();
    }

    /// <summary>
    /// A cache refresh empties the values and puts the surviving ones back. It must not be read as
    /// a removal, or a refresh would cost the table the formatting of values it still has.
    /// </summary>
    [Test]
    [Property("Description", "#577: pruning is for user-initiated removal, not for a cache refresh")]
    public async Task Refreshing_the_cache_keeps_the_formats_naming_a_value()
    {
        using var saved = SaveFixture(pt => pt.PivotCache.Refresh());

        var xml = PivotTableXml(saved);

        await Assert.That(DataFieldReferences(xml)).IsEquivalentTo(AsLoaded);
    }

    /// <summary>
    /// A refresh keeps only the values whose source column is still in the cache. The rest are
    /// dropped for good, so they are a removal and their formats must be pruned as one.
    /// </summary>
    [Test]
    [Property("Description", "#577: a refresh that drops a value must not leave its formats behind")]
    public async Task Refreshing_onto_a_narrower_source_prunes_the_formats_of_the_dropped_values()
    {
        // The source is A1:C5, columns F1, F2 and F3, with a value on each. Narrowing it to A:B
        // leaves F1 and F2, so the value at position 2 is gone for good.
        //
        // Narrowing all the way to column A leaves only one value, which is #583: the 'Values'
        // field still sits on the column axis at this point, and the loop used to read
        // oldNames[-2] for it once includeDataField turned false. See
        // Refreshing_onto_a_single_column_source_takes_the_values_field_off_the_axis below.
        using var saved = SaveFixture(pt =>
        {
            var source = pt.PivotCache.SourceRange!;
            pt.PivotCache.SetSourceRange(source.Worksheet.Range("A1:B5"));
            pt.PivotCache.Refresh();
        });

        var xml = PivotTableXml(saved);

        await Assert.That(DataFieldCount(xml)).IsEqualTo(2)
            .Because("only the values whose source columns survived are put back");
        await Assert.That(DataFieldReferences(xml)).IsEquivalentTo(["1", "0", "0,1", "0,1"])
            .Because("the format naming only the dropped value goes, and the two naming all three "
                     + "values must be left naming the two that are left");
    }

    /// <summary>
    /// #583: narrowing the source down to one surviving value leaves the 'Values' field on the
    /// column axis with nothing to distinguish, since <c>GetKeptNames</c> used to fall through to
    /// <c>oldNames[-2]</c> and throw <see cref="ArgumentOutOfRangeException"/> once the field was
    /// no longer kept. A single value needs no 'Values' field, so refreshing must take it off the
    /// axis instead, the way removing values one at a time already does (#572).
    /// </summary>
    [Test]
    [Property("Description", "#583: a refresh leaving one value must not throw, and must take Values off the axis")]
    public async Task Refreshing_onto_a_single_column_source_takes_the_values_field_off_the_axis()
    {
        using var saved = SaveFixture(pt =>
        {
            var source = pt.PivotCache.SourceRange!;
            pt.PivotCache.SetSourceRange(source.Worksheet.Range("A1:A5"));
            pt.PivotCache.Refresh();
        });

        var xml = PivotTableXml(saved);

        await Assert.That(DataFieldCount(xml)).IsEqualTo(1)
            .Because("only F1's value survives the narrower source");
        await Assert.That(xml).DoesNotContain("<colFields")
            .Because("one value needs no 'Values' field to tell it apart from another");
    }

    /// <summary>
    /// #583: the same fall-through is reached with zero surviving values, not just one, because
    /// includeDataField is false in both cases.
    /// </summary>
    [Test]
    [Property("Description", "#583: a refresh leaving no values must not throw, and must take Values off the axis")]
    public async Task Refreshing_onto_a_source_with_no_matching_columns_drops_every_value()
    {
        using var saved = SaveFixture(pt =>
        {
            var otherSheet = pt.Worksheet.Workbook.AddWorksheet("Unrelated source");
            otherSheet.Cell("A1").Value = "Other";
            otherSheet.Cell("A2").Value = 1;
            otherSheet.Cell("A3").Value = 2;

            pt.PivotCache.SetSourceRange(otherSheet.Range("A1:A3"));
            pt.PivotCache.Refresh();
        });

        var xml = PivotTableXml(saved);

        await Assert.That(xml).DoesNotContain("<dataFields")
            .Because("none of F1, F2 or F3 survive in the new source, so no value is left");
        await Assert.That(xml).DoesNotContain("<colFields")
            .Because("with no values left, the 'Values' field has nothing to tell apart");
    }

    /// <summary>
    /// The loader deliberately does not check a format's position against the data fields, so a
    /// file can arrive naming a value it does not have. Renumbering such a position would walk it
    /// down onto a value it was never written for, so it is dropped instead.
    /// </summary>
    [Test]
    [Property("Description", "#577: a position the file never had must not be renumbered onto a real value")]
    public async Task A_position_past_the_values_is_dropped_rather_than_renumbered()
    {
        using var saved = SaveFixture(pt =>
        {
            // Name a position the table does not have, as a hand-made file can.
            var area = pt.Formats.First(f => f.PivotArea.References
                .Any(r => r.Field == DataFieldReferenceIndex)).PivotArea;
            area.References.First(r => r.Field == DataFieldReferenceIndex).FieldItems.Add(9);

            pt.Values.Remove(FirstValueName(pt));
        });

        var xml = PivotTableXml(saved);

        await Assert.That(DataFieldReferences(xml).Any(reference => reference.Contains('9'))).IsFalse()
            .Because("position 9 names no value, before or after the removal");
        await Assert.That(DataFieldReferences(xml)).IsEquivalentTo(["0", "1", "0,1", "0,1"])
            .Because("the out-of-range position goes and the rest are renumbered as usual");
    }

    private static string FirstValueName(XLPivotTable pivotTable)
    {
        return pivotTable.Values.First().CustomName;
    }

    /// <summary>
    /// Load the fixture, apply <paramref name="change"/> to its pivot table and save.
    /// </summary>
    private static MemoryStream SaveFixture(Action<XLPivotTable> change)
    {
        var saved = new MemoryStream();
        using (var wb = new XLWorkbook(TestHelper.GetStreamFromResource(
                   TestHelper.GetResourcePath(FixturePath))))
        {
            var pt = (XLPivotTable)wb.Worksheets.SelectMany(ws => ws.PivotTables).First();

            // Guard the premise: the assertions below say nothing unless the fixture really holds
            // three values and the five data-field references this file is chosen for.
            if (pt.Values.Count() != 3)
                throw new InvalidOperationException($"The fixture must hold three values, not {pt.Values.Count()}.");

            change(pt);
            wb.SaveAs(saved);
        }

        return saved;
    }

    /// <summary>
    /// The <c>&lt;x v&gt;</c> items of every <c>&lt;reference field="4294967294"&gt;</c> inside
    /// <c>&lt;formats&gt;</c>, one comma-joined string per reference. Read off the saved XML rather
    /// than the object model, because the point is what the file says.
    /// </summary>
    private static List<string> DataFieldReferences(string pivotTableXml)
    {
        var formats = Regex.Match(pivotTableXml, "<formats.*?</formats>", RegexOptions.Singleline);
        if (!formats.Success)
            return [];

        var references = Regex.Matches(
            formats.Value,
            "<reference[^>]*field=\"4294967294\"[^>]*>(?<items>.*?)</reference>",
            RegexOptions.Singleline);

        return references
            .Select(reference => string.Join(",", Regex.Matches(reference.Groups["items"].Value, "<x v=\"(?<v>\\d+)\"")
                .Select(item => item.Groups["v"].Value)))
            .ToList();
    }

    /// <summary>
    /// How many <c>dataField</c> elements the saved table has.
    /// </summary>
    private static int DataFieldCount(string pivotTableXml)
    {
        return Regex.Matches(pivotTableXml, "<dataField[ /]").Count;
    }

    private static int FormatCount(string pivotTableXml)
    {
        var formats = Regex.Match(pivotTableXml, "<formats[^>]*count=\"(?<count>\\d+)\"");
        return formats.Success ? int.Parse(formats.Groups["count"].Value) : 0;
    }

    private static string PivotTableXml(Stream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var entry = archive.Entries
            .First(e => e.FullName.StartsWith("xl/pivotTables/pivotTable", StringComparison.Ordinal))
            .Open();
        using var reader = new StreamReader(entry);
        return reader.ReadToEnd();
    }
}
