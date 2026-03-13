using System;
using System.Drawing;
using XLibur.Graphics;

namespace XLibur.Excel;

/// <summary>
/// A class that defines various aspects of a newly created workbook.
/// </summary>
public class LoadOptions
{
    /// <summary>
    /// A graphics engine that will be used for workbooks without explicitly set engine.
    /// </summary>
    public static IXLGraphicEngine? DefaultGraphicEngine { internal get; set; }

    /// <summary>
    /// Should all formulas in a workbook be recalculated during a load? Default value is <c>false</c>.
    /// </summary>
    public bool RecalculateAllFormulas { get; set; } = false;

    /// <summary>
    /// Graphic engine used by the workbook.
    /// </summary>
    public IXLGraphicEngine? GraphicEngine { get; set; }

    /// <summary>
    /// DPI for the workbook. Default is 96.
    /// </summary>
    /// <remarks>Used in various places, e.g., determining a physical size of an image without a DPI or to determine a size of a text in a cell.</remarks>
    public Point Dpi
    {
        get;
        set => field = value is { X: > 0, Y: > 0 } ? value : throw new ArgumentException("DPI must be positive");
    } = new(96, 96);
}
