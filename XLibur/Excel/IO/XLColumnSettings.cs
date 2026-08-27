using System;
using XLibur.Extensions;

namespace XLibur.Excel.IO;

/// <summary>The <c>&lt;col&gt;</c> attributes both write paths emit, with defaults applied.</summary>
/// <remarks>
/// <para>
/// This type owns the per-<c>&lt;col&gt;</c> attribute decision; the two writers own only the
/// emission. Before spec 29 only the width rule was shared (the streaming path called
/// <see cref="ColumnWriter.GetColumnWidth"/> directly); <c>customWidth</c>, <c>style</c>,
/// <c>hidden</c>, <c>outlineLevel</c> and <c>collapsed</c> were decided twice.
/// </para>
/// <para>
/// It does not own <em>which</em> columns get written. The DOM path expands every column in
/// <c>[min, max]</c>, back-fills the columns either side with the worksheet style and collapses
/// equal neighbours into runs; the streaming path writes one <c>&lt;col&gt;</c> per registered
/// range and does no filling or collapsing. Those are different products and stay where they are.
/// </para>
/// </remarks>
internal readonly struct XLColumnSettings
{
    /// <summary><c>min</c>, 1-based and inclusive.</summary>
    internal required uint Min { get; init; }

    /// <summary><c>max</c>, 1-based and inclusive.</summary>
    internal required uint Max { get; init; }

    /// <summary><c>style</c>, or <c>null</c> when the column carries no style of its own.</summary>
    internal required uint? StyleId { get; init; }

    /// <summary>Already through <see cref="ColumnWriter.GetColumnWidth"/> and <c>SaveRound</c>.</summary>
    internal required double? Width { get; init; }

    /// <summary><c>hidden</c>.</summary>
    internal required bool Hidden { get; init; }

    /// <summary><c>collapsed</c>.</summary>
    internal required bool Collapsed { get; init; }

    /// <summary><c>outlineLevel</c>. 0 means the attribute is omitted.</summary>
    internal required byte OutlineLevel { get; init; }

    /// <summary><c>customWidth</c> accompanies a width and is omitted without one.</summary>
    internal bool CustomWidth => Width is not null;

    /// <param name="min">First column the settings apply to, 1-based.</param>
    /// <param name="max">Last column the settings apply to, 1-based and inclusive.</param>
    /// <param name="styleId">The column's own style id, or <c>null</c> for none.</param>
    /// <param name="rawWidth">
    /// The column width as the model holds it, <em>before</em> <see cref="ColumnWriter.GetColumnWidth"/>
    /// and <c>SaveRound</c>, or <c>null</c> to omit the width and <c>customWidth</c> with it. Passing
    /// an already-resolved width here rounds it twice.
    /// </param>
    /// <param name="hidden">Whether the column is hidden.</param>
    /// <param name="collapsed">Whether the outline group the column belongs to is collapsed.</param>
    /// <param name="outlineLevel">Outline level, 0 for none.</param>
    internal static XLColumnSettings Resolve(
        uint min, uint max, uint? styleId, double? rawWidth,
        bool hidden, bool collapsed, int outlineLevel)
        => new()
        {
            Min = min,
            Max = max,
            StyleId = styleId,
            Width = rawWidth is { } w ? ColumnWriter.GetColumnWidth(w).SaveRound() : null,
            Hidden = hidden,
            Collapsed = collapsed,
            OutlineLevel = outlineLevel > 0 ? (byte)Math.Min(outlineLevel, byte.MaxValue) : (byte)0,
        };
}
