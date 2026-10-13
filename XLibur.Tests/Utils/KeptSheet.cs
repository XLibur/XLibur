using System;
using System.IO;
using XLibur.Excel;
using SaveOptions = XLibur.Excel.SaveOptions;

namespace XLibur.Tests.Utils;

/// <summary>
/// Marks a worksheet part so a test can tell what the first save of a loaded workbook did with it
/// (#702): kept the part whole (tier 1), kept only its cells (tier 2), or wrote it from the model.
/// </summary>
/// <remarks>
/// <see cref="Mark"/> spells the part's XML declaration as no save writes it, and the number 1.5 as
/// <c>1.50</c>, which XLibur writes as <c>1.5</c>. A part that still has the declaration was kept
/// whole; one that still has <c>1.50</c> had at least its cells kept. A rewritten zip entry keeps
/// its compression method, so the entry itself shows nothing.
/// </remarks>
internal static class KeptSheet
{
    /// <summary>The XML declaration as no save writes it.</summary>
    private const string Declaration = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";

    /// <summary>The number 1.5 as XLibur never writes it.</summary>
    private const string Number = "<x:v>1.50</x:v>";

    /// <summary>
    /// Respells the declaration of <paramref name="partName"/>, and every <c>1.5</c> in its cells as
    /// <c>1.50</c>. Returns the same stream, rewound.
    /// </summary>
    internal static MemoryStream Mark(this MemoryStream package, string partName) =>
        package.RewritePart(partName, xml =>
        {
            var afterDeclaration = xml.IndexOf("?>", StringComparison.Ordinal) + 2;
            return Declaration + xml[afterDeclaration..].Replace("<x:v>1.5</x:v>", Number, StringComparison.Ordinal);
        });

    /// <summary>Is the part as the file had it, declaration and all?</summary>
    internal static bool IsKeptWhole(this Stream package, string partName) =>
        package.ReadPart(partName).StartsWith(Declaration, StringComparison.Ordinal);

    /// <summary>Does the part still have the cells as the file had them?</summary>
    internal static bool HasKeptCells(this Stream package, string partName) =>
        package.ReadPart(partName).Contains(Number, StringComparison.Ordinal);

    /// <summary>Was the part written from the model, cells and all?</summary>
    internal static bool IsWrittenFromModel(this Stream package, string partName) =>
        !package.IsKeptWhole(partName) && !package.HasKeptCells(partName);

    /// <summary>
    /// Loads <paramref name="source"/>, applies <paramref name="edit"/> and saves it to a new stream,
    /// rewound.
    /// </summary>
    internal static MemoryStream LoadAndSave(this MemoryStream source, Action<XLWorkbook>? edit = null,
        SaveOptions? options = null)
    {
        source.Position = 0;
        using var wb = new XLWorkbook(source);
        edit?.Invoke(wb);

        var saved = new MemoryStream();
        wb.SaveAs(saved, options ?? new SaveOptions());
        saved.Position = 0;
        return saved;
    }
}
