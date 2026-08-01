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
    /// A font engine that will be used for workbooks without explicitly set font engine.
    /// </summary>
    public static IXLFontEngine? DefaultFontEngine { get; set; }

    /// <summary>
    /// Should all formulas in a workbook be recalculated during a load? Default value is <c>false</c>.
    /// </summary>
    public bool RecalculateAllFormulas { get; set; } = false;

    /// <summary>
    /// Password used to decrypt a password-protected workbook. Default value is <c>null</c>, which
    /// is correct for the unencrypted files that make up nearly all input.
    /// </summary>
    /// <remarks>
    /// Only consulted when the file turns out to be encrypted, so setting it for a file that isn't
    /// encrypted is harmless and setting it wrongly for one that is throws
    /// <see cref="Exceptions.XLInvalidPasswordException"/>. This is the password that encrypts the
    /// whole file, unrelated to workbook and sheet protection, which control edit permissions
    /// inside a file that anyone can open.
    /// <para>
    /// When the file was encrypted, the password is retained for the lifetime of the workbook so
    /// that <c>Save</c> can write it back encrypted. It is an ordinary string, as passwords are
    /// throughout this API, so it lives until the GC reclaims it; keep the workbook short-lived if
    /// that matters to your threat model.
    /// </para>
    /// </remarks>
    public string? Password { get; set; }

    /// <summary>
    /// Graphic engine used by the workbook.
    /// </summary>
    public IXLGraphicEngine? GraphicEngine { get; set; }

    /// <summary>
    /// Font engine used by the workbook for text measurement and font metrics.
    /// </summary>
    /// <remarks>
    /// Resolution order: <see cref="FontEngine"/> if set, then <see cref="DefaultFontEngine"/>,
    /// then the workbook's <see cref="GraphicEngine"/> if it implements <see cref="IXLFontEngine"/>,
    /// otherwise a <see cref="GraphicEngineFontAdapter"/> wrapping the graphic engine.
    /// If no font engine or graphic engine is available, an <see cref="System.InvalidOperationException"/> is thrown.
    /// </remarks>
    public IXLFontEngine? FontEngine { get; set; }

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
