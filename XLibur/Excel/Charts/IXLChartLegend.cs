namespace XLibur.Excel;

/// <summary>
/// Where the legend sits relative to the plot area.
/// </summary>
public enum XLLegendPosition
{
    /// <summary>To the right of the plot area. This is the default.</summary>
    Right,

    /// <summary>Below the plot area.</summary>
    Bottom,

    /// <summary>To the left of the plot area.</summary>
    Left,

    /// <summary>Above the plot area.</summary>
    Top,

    /// <summary>In the top right corner.</summary>
    TopRight
}

/// <summary>
/// The legend of a chart — the key that names each series.
/// </summary>
/// <remarks>
/// Charts XLibur creates have no legend until <see cref="Visible"/> is set, and a chart read from a
/// file keeps the legend it came with unless one of these properties is assigned.
/// </remarks>
public interface IXLChartLegend
{
    /// <summary>
    /// Gets or sets whether the legend is shown. Defaults to <c>false</c>.
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// Gets or sets where the legend sits. Defaults to <see cref="XLLegendPosition.Right"/>.
    /// Ignored while <see cref="Visible"/> is <c>false</c>.
    /// </summary>
    XLLegendPosition Position { get; set; }

    /// <summary>
    /// Gets or sets whether the legend is drawn on top of the plot area rather than beside it,
    /// which gives the plot more room. Defaults to <c>false</c>.
    /// </summary>
    bool Overlay { get; set; }
}
