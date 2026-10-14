using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.IO;

/// <summary>
/// The <c>sqref</c> values of the elements a writer rebuilds from the model, such as data
/// validations and conditional formats, as the loaded part spelled them (#722).
/// </summary>
/// <remarks>
/// The same cells can be spelled more than one way: Excel writes a single cell as <c>B5</c>, and an
/// earlier XLibur wrote it as <c>B5:B5</c>. A rebuilt element whose ranges name the cells a loaded one
/// named, in the same order, is given the loaded spelling, so a part that did not change can be kept
/// (#702).
/// </remarks>
internal sealed class SqrefSpellings
{
    private readonly Dictionary<string, string> _byCells = new(StringComparer.Ordinal);

    /// <param name="loaded">The <c>sqref</c> of each element the part had.</param>
    internal SqrefSpellings(IEnumerable<string?> loaded)
    {
        foreach (var sqref in loaded)
        {
            if (sqref is not null && CellsOf(sqref) is { } cells)
                _byCells.TryAdd(cells, sqref);
        }
    }

    /// <summary>
    /// The loaded spelling of <paramref name="written"/>, or <paramref name="written"/> when the part
    /// named no such cells.
    /// </summary>
    internal string Spell(string written) =>
        _byCells.Count > 0 && CellsOf(written) is { } cells && _byCells.TryGetValue(cells, out var loaded)
            ? loaded
            : written;

    /// <summary>
    /// The areas of <paramref name="sqref"/> in a form that does not depend on spelling, or null when
    /// one of them is not a cell or a range of cells.
    /// </summary>
    private static string? CellsOf(string sqref)
    {
        var cells = new StringBuilder();
        foreach (var token in sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (!Area.TryParse(token, out var area))
                return null;

            cells.Append(CultureInfo.InvariantCulture,
                $"{area.FirstPoint.Row},{area.FirstPoint.Column},{area.LastPoint.Row},{area.LastPoint.Column};");
        }

        return cells.ToString();
    }
}
