namespace XLibur.Excel;

/// <summary>
/// A description of formatting that should be applied to a <see cref="XLPivotTable"/>.
/// </summary>
internal sealed class XLPivotFormat
{
    internal XLPivotFormat(XLPivotArea pivotArea)
    {
        PivotArea = pivotArea;
    }

    /// <summary>
    /// Pivot area that should be formatted.
    /// </summary>
    internal XLPivotArea PivotArea { get; }

    /// <summary>
    /// Should the formatting (determined by <see cref="DxfStyleValue"/>) be applied or not?
    /// </summary>
    internal XLPivotFormatAction Action { get; init; } = XLPivotFormatAction.Formatting;

    /// <summary>
    /// Differential formatting to apply to the <see cref="PivotArea"/>. It can be empty, e.g. if
    /// <see cref="Action"/> is blank. Empty dxf is represented by <see cref="XLStyle.Default"/>,
    /// until we get better dxf representation.
    /// </summary>
    internal XLStyleValue DxfStyleValue { get; set; } = XLStyle.Default.Value;

    /// <summary>
    /// The dxf the format was loaded with, and the style it decoded to.
    /// </summary>
    internal (XLStyleValue Style, string Xml)? LoadedDxf { get; init; }

    /// <summary>
    /// The loaded dxf while <see cref="DxfStyleValue"/> still holds what it decoded to. The save
    /// writes it as loaded, because the style model cannot hold all of a dxf: the inside borders
    /// (#956), or a property the dxf clears (#949).
    /// </summary>
    internal string? UnchangedLoadedDxfXml =>
        LoadedDxf is { } loaded && loaded.Style.Equals(DxfStyleValue) ? loaded.Xml : null;
}
