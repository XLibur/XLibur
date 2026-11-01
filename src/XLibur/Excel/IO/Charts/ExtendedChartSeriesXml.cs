using System;
using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Excel.Coordinates;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace XLibur.Excel.IO.Charts;

/// <summary>
/// Writes what a sheet rename or delete did to the series of a loaded ChartEx chart into its part. The
/// rest of the part is kept as it was loaded (see <see cref="ChartPatcher"/>).
/// </summary>
/// <remarks>
/// <para>
/// A ChartEx chart keeps each series' data in a <c>cx:data</c> of <c>cx:chartData</c>, which the series
/// names by id: a <c>cx:strDim</c> for the categories and a <c>cx:numDim</c> for the values, each with
/// its reference in a <c>cx:f</c>. The series' name is in <c>cx:tx/cx:txData</c>, as a <c>cx:f</c> and
/// the cached <c>cx:v</c>. <see cref="ChartReader"/> reads the same elements in the same order, so the
/// n-th <c>cx:series</c> is the n-th series of the model.
/// </para>
/// <para>
/// Excel writes a hidden name in each <c>cx:f</c>, not a reference. A rename renames the sheet in the
/// name, and the part does not change. A delete removes the name (see
/// <see cref="XLChartSeries.DropChartDataNames"/>), and this takes the reference out of the part as
/// Excel did in the <c>chartex-pivotcf-*</c> fixture. A chart XLibur created holds its references in
/// its <c>cx:f</c> elements, and a reference a rename or delete rewrote is written where it stands, as
/// <see cref="ChartWriter"/> writes it into a chart it creates.
/// </para>
/// </remarks>
internal static class ExtendedChartSeriesXml
{
    internal static void Apply(Cx.ChartSpace chartSpace, XLChartSeriesCollection series)
    {
        var chartData = chartSpace.Descendants<Cx.ChartData>().FirstOrDefault();
        var elements = chartSpace.Descendants<Cx.Series>().ToList();
        var count = Math.Min(series.Count, elements.Count);
        for (var i = 0; i < count; i++)
        {
            var xlSeries = series.Items[i];
            var assigned = xlSeries.AssignedFormat;
            if (assigned == XLChartSeriesFormat.None)
                continue;

            PatchName(elements[i], xlSeries, assigned);

            var data = FindData(elements[i], chartData);
            if (data is null)
                continue;

            PatchDimension(data.Elements<Cx.StringDimension>().FirstOrDefault(), xlSeries.CategoryReferences,
                assigned.HasFlag(XLChartSeriesFormat.CategoryReferencesRewritten), xlSeries.DroppedCategoryArea);
            PatchDimension(data.Elements<Cx.NumericDimension>().FirstOrDefault(), xlSeries.ValueReferences,
                assigned.HasFlag(XLChartSeriesFormat.ValueReferencesRewritten), xlSeries.DroppedValueArea);
        }
    }

    /// <summary>
    /// Takes out the series' name when a delete took its reference, as Excel did: the fixture's series
    /// kept no <c>cx:tx</c>, neither the reference nor the cached name. A reference a rename or delete
    /// rewrote is written where it stands.
    /// </summary>
    private static void PatchName(Cx.Series element, XLChartSeries xlSeries, XLChartSeriesFormat assigned)
    {
        var tx = element.Elements<Cx.Text>().FirstOrDefault();
        if (tx is null)
            return;

        if (assigned.HasFlag(XLChartSeriesFormat.NameReferenceDropped))
        {
            tx.Remove();
            return;
        }

        if (assigned.HasFlag(XLChartSeriesFormat.NameReference)
            && xlSeries.NameReference is { } reference
            && tx.Descendants<Cx.Formula>().FirstOrDefault() is { } formula)
            formula.Text = reference;
    }

    private static Cx.Data? FindData(Cx.Series element, Cx.ChartData? chartData)
    {
        var dataId = element.Elements<Cx.DataId>().FirstOrDefault()?.Val?.Value;
        if (dataId is null || chartData is null)
            return null;

        return chartData.Elements<Cx.Data>().FirstOrDefault(d => d.Id?.Value == dataId);
    }

    /// <summary>
    /// Takes the reference out of <paramref name="dimension"/> when a delete dropped it, and otherwise
    /// writes a reference a rename or delete rewrote.
    /// </summary>
    private static void PatchDimension(OpenXmlCompositeElement? dimension, string? reference, bool rewritten,
        Area? dropped)
    {
        if (dimension is null)
            return;

        var formula = dimension.Elements<Cx.Formula>().FirstOrDefault();
        if (dropped is { } area)
            DropReference(dimension, formula, area);
        else if (rewritten)
            WriteReference(formula, reference);
    }

    /// <summary>
    /// Takes a dropped reference out of <paramref name="dimension"/>: its <c>cx:f</c> and its levels
    /// go, and an empty level is left for each level the reference had.
    /// </summary>
    /// <remarks>
    /// The fixture shows it: <c>Data!$B$1:$B$3</c>, read by row (<c>dir="row"</c>), left three empty
    /// levels, and <c>Data!$B$4</c> one. Read by column, the default, a level is a column; the fixture
    /// shows no such dimension.
    /// </remarks>
    private static void DropReference(OpenXmlCompositeElement dimension, Cx.Formula? formula, Area area)
    {
        // Without its cx:f the dimension was patched by an earlier save. Save() and SaveAs() start
        // from the package the last save wrote, and a dropped reference stays dropped in the model,
        // so the patch leaves its own work as it is: the direction it counted levels by is gone.
        if (formula is null)
            return;

        var levels = formula.Dir?.Value == Cx.FormulaDirection.Row ? area.Height : area.Width;
        formula.Remove();
        RemoveLevels(dimension);
        AddEmptyLevels(dimension, levels);
    }

    private static void RemoveLevels(OpenXmlCompositeElement dimension)
    {
        foreach (var level in dimension.ChildElements.Where(IsLevel).ToList())
            level.Remove();
    }

    private static bool IsLevel(OpenXmlElement element) => element is Cx.StringLevel or Cx.NumericLevel;

    /// <summary>
    /// Adds <paramref name="count"/> empty levels to <paramref name="dimension"/>, before its extension
    /// list, which ends it.
    /// </summary>
    private static void AddEmptyLevels(OpenXmlCompositeElement dimension, int count)
    {
        var extensions = dimension.ChildElements.FirstOrDefault(e => e.LocalName == "extLst");
        for (var i = 0; i < count; i++)
        {
            var level = EmptyLevel(dimension);
            if (extensions is null)
                dimension.AppendChild(level);
            else
                dimension.InsertBefore(level, extensions);
        }
    }

    private static OpenXmlElement EmptyLevel(OpenXmlCompositeElement dimension)
        => dimension is Cx.StringDimension
            ? new Cx.StringLevel { PtCount = 0U }
            : new Cx.NumericLevel { PtCount = 0U };

    /// <summary>Writes a reference a rename or delete rewrote where it stands.</summary>
    private static void WriteReference(Cx.Formula? formula, string? reference)
    {
        if (formula is not null && reference is not null)
            formula.Text = reference;
    }
}
