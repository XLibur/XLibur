using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The dxfs the package had before the save rebuilt them from the model, for a <c>dxfId</c> that a
/// writer carries over from a loaded part rather than derives from the model (#904).
/// </summary>
internal sealed class LoadedDifferentialFormats
{
    private readonly Stylesheet _stylesheet;
    private readonly DifferentialFormat[] _loaded;
    private readonly Dictionary<uint, uint> _kept = new();

    internal LoadedDifferentialFormats(Stylesheet stylesheet)
    {
        _stylesheet = stylesheet;
        _loaded = stylesheet.DifferentialFormats?.Elements<DifferentialFormat>().ToArray() ?? [];
    }

    internal uint? WrittenIndexOf(uint loadedId)
    {
        if (_kept.TryGetValue(loadedId, out var keptId))
            return keptId;

        if (loadedId >= _loaded.Length)
            return null;

        var differentialFormats = _stylesheet.DifferentialFormats ??= new DifferentialFormats();
        differentialFormats.AppendChild((DifferentialFormat)_loaded[loadedId].CloneNode(true));
        keptId = (uint)differentialFormats.Elements<DifferentialFormat>().Count() - 1;
        differentialFormats.Count = keptId + 1;
        _kept.Add(loadedId, keptId);
        return keptId;
    }
}
