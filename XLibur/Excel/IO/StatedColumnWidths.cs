using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The column widths one worksheet part states, as the file has them: the sheet's default from
/// <c>&lt;sheetFormatPr&gt;</c>, and each <c>&lt;col&gt;</c>. Recorded while the sheet loads, for
/// <see cref="AnchorGrid"/> to lay the columns out as Excel draws them (#782).
/// </summary>
/// <remarks>
/// The model cannot give these back. A sheet default from <c>defaultColWidth</c> or
/// <c>baseColWidth</c> is held in characters, converted with the font of the default cell format,
/// while a default from a <c>&lt;col&gt;</c> range that runs to the last column, and the width of
/// every column the file sized, is held as the file's width less its padding, snapped to whole
/// pixels only when it is a width Excel wrote (<see cref="XLHelper.FileWidthToNoC"/>). A column the file sized to its sheet's default holds
/// the same width as one the file did not size at all. Each form converts to pixels differently.
/// </remarks>
internal sealed class StatedColumnWidths
{
    private readonly List<(int Min, int Max, double? Width, bool Hidden)> _columns = [];

    /// <summary><c>defaultColWidth</c>: a width in the unit of <c>&lt;col width&gt;</c>.</summary>
    internal double? DefaultColumnWidth { get; private set; }

    /// <summary><c>baseColWidth</c>, in digits without padding. 8 when the file does not state it.</summary>
    internal double BaseColumnWidth { get; private set; } = 8;

    /// <summary>Each <c>&lt;col&gt;</c>, in the order of the file.</summary>
    internal IReadOnlyList<(int Min, int Max, double? Width, bool Hidden)> Columns => _columns;

    internal void Read(SheetFormatProperties sheetFormat)
    {
        DefaultColumnWidth = sheetFormat.DefaultColumnWidth?.Value;
        if (sheetFormat.BaseColumnWidth?.Value is { } baseWidth)
            BaseColumnWidth = baseWidth;
    }

    internal void Read(Columns columns)
    {
        foreach (var column in columns.Elements<Column>())
        {
            if (column.Min?.Value is not { } min || column.Max?.Value is not { } max)
                continue;

            _columns.Add(((int)min, (int)max, column.Width?.Value, column.Hidden?.Value ?? false));
        }
    }
}
