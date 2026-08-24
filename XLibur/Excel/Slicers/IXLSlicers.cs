using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel;

/// <summary>
/// The slicers drawn on a worksheet.
/// </summary>
/// <remarks>
/// The worksheet owns its slicers: this is the collection a slicer is added to and removed from.
/// What a slicer filters is a separate relationship held by its cache, exposed as a view on
/// <see cref="IXLPivotTable.Slicers"/>.
/// </remarks>
public interface IXLSlicers : IEnumerable<IXLSlicer>
{
    /// <summary>
    /// The number of slicers on the worksheet.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// The slicer with the given <see cref="IXLSlicer.Name"/>.
    /// </summary>
    /// <param name="name">The slicer's internal name, which is not its caption.</param>
    /// <exception cref="System.Collections.Generic.KeyNotFoundException">
    /// The worksheet has no slicer with that name.
    /// </exception>
    IXLSlicer Slicer(string name);

    /// <summary>
    /// Finds the slicer with the given <see cref="IXLSlicer.Name"/>.
    /// </summary>
    /// <param name="name">The slicer's internal name, which is not its caption.</param>
    /// <param name="slicer">The slicer, when one was found.</param>
    /// <returns>Whether a slicer with that name is on the worksheet.</returns>
    bool TryGetSlicer(string name, [NotNullWhen(true)] out IXLSlicer? slicer);
}
