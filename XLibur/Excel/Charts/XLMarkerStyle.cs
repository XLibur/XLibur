namespace XLibur.Excel;

/// <summary>
/// The marker symbol drawn at each data point of a line, scatter or radar series.
/// </summary>
public enum XLMarkerStyle
{
    /// <summary>
    /// No explicit symbol: the marker element is omitted and Excel picks the default for the
    /// chart type (markers for <c>LineWithMarkers*</c> and scatter types, none otherwise).
    /// This is the default.
    /// </summary>
    Auto,

    /// <summary>No marker is drawn.</summary>
    None,

    /// <summary>A filled circle.</summary>
    Circle,

    /// <summary>A short horizontal dash.</summary>
    Dash,

    /// <summary>A filled diamond.</summary>
    Diamond,

    /// <summary>A small filled dot.</summary>
    Dot,

    /// <summary>A plus sign.</summary>
    Plus,

    /// <summary>A filled square.</summary>
    Square,

    /// <summary>A filled star.</summary>
    Star,

    /// <summary>A filled triangle.</summary>
    Triangle,

    /// <summary>An X (cross) symbol.</summary>
    X
}
