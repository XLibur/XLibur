namespace XLibur.Excel;

/// <summary>
/// The direction an axis runs in.
/// </summary>
public enum XLAxisOrientation
{
    /// <summary>Smallest value first — left to right, or bottom to top. This is the default.</summary>
    MinMax,

    /// <summary>Reversed: largest value first.</summary>
    MaxMin
}

/// <summary>
/// One axis of a chart.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="IXLChart.CategoryAxis"/> is the horizontal axis and <see cref="IXLChart.ValueAxis"/>
/// the vertical one. On a scatter or bubble chart both are value axes, so the "category" axis takes
/// numbers there too.
/// </para>
/// <para>
/// Not every property reaches every axis: <see cref="MajorUnit"/>, <see cref="MinorUnit"/> and
/// <see cref="LogScale"/> exist only on a value axis in the file format, and are skipped on a
/// category axis. Pie and doughnut charts have no axes at all and ignore everything here.
/// </para>
/// </remarks>
public interface IXLChartAxis
{
    /// <summary>
    /// Gets or sets the axis title. <c>null</c> — the default — means no title.
    /// </summary>
    string? Title { get; set; }

    /// <summary>
    /// Gets or sets the number format the axis labels are drawn with, e.g. <c>"#,##0"</c>.
    /// <c>null</c> — the default — takes the format from the source cells.
    /// </summary>
    string? NumberFormat { get; set; }

    /// <summary>
    /// Gets or sets the smallest value shown on the axis. <c>null</c> — the default — lets Excel
    /// choose it from the data. Value axes only.
    /// </summary>
    double? Min { get; set; }

    /// <summary>
    /// Gets or sets the largest value shown on the axis. <c>null</c> — the default — lets Excel
    /// choose it from the data. Value axes only.
    /// </summary>
    double? Max { get; set; }

    /// <summary>
    /// Gets or sets the interval between major tick marks and labels. <c>null</c> — the default —
    /// lets Excel choose. Value axes only.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">The value is zero or negative.</exception>
    double? MajorUnit { get; set; }

    /// <summary>
    /// Gets or sets the interval between minor tick marks. <c>null</c> — the default — lets Excel
    /// choose. Value axes only.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">The value is zero or negative.</exception>
    double? MinorUnit { get; set; }

    /// <summary>
    /// Gets or sets whether the axis is drawn. Defaults to <c>true</c>. A hidden axis still scales
    /// the plot — it just is not painted.
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// Gets or sets whether major gridlines are drawn across the plot area from this axis.
    /// Defaults to <c>false</c>.
    /// </summary>
    bool MajorGridlines { get; set; }

    /// <summary>
    /// Gets or sets the direction the axis runs in. Defaults to
    /// <see cref="XLAxisOrientation.MinMax"/>.
    /// </summary>
    XLAxisOrientation Orientation { get; set; }

    /// <summary>
    /// Gets or sets whether the axis is scaled logarithmically. Defaults to <c>false</c>.
    /// Value axes only.
    /// </summary>
    bool LogScale { get; set; }

    /// <summary>
    /// Gets or sets the base of the logarithmic scale, used when <see cref="LogScale"/> is
    /// <c>true</c>. Defaults to 10.
    /// </summary>
    /// <exception cref="System.ArgumentOutOfRangeException">
    /// The value is outside the range Excel accepts (2 to 1000).
    /// </exception>
    double LogBase { get; set; }
}
