using System.IO.Packaging;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Excel.Exceptions;
using XLibur.Fonts.PinataLabs;

namespace XLibur.Operations.Workbooks;

internal static class WorkbookLoader
{
    /// <summary>
    /// The options every workbook this tool opens or creates is given. XLibur needs a font engine to
    /// load a workbook; PinataLabs is managed code with embedded fallback fonts, so the tool needs no
    /// native library and measures the same on every machine. It is passed per workbook rather than
    /// registered process-wide, so a host that embeds the operations keeps its own default.
    /// </summary>
    internal static LoadOptions Options(bool recalculate = false) => new()
    {
        FontEngine = PinataLabsFontEngine.Instance.Value,
        RecalculateAllFormulas = recalculate,
    };

    internal static XLWorkbook Create() => new(Options());

    internal static XLWorkbook Load(byte[] bytes, string displayPath, Limits limits, bool recalculate)
    {
        ZipGuard.Check(bytes, displayPath, limits);
        try
        {
            return new XLWorkbook(new MemoryStream(bytes, writable: false), Options(recalculate));
        }
        catch (Exception e) when (e is NotImplementedException or NotSupportedException)
        {
            throw OperationException.Unsupported("unsupported_workbook", $"'{displayPath}' uses a feature XLibur does not load: {e.Message}", e);
        }
        catch (Exception e) when (e is XLiburException or XmlException or InvalidDataException or FileFormatException
                                      or OpenXmlPackageException or FormatException)
        {
            throw OperationException.BadInput("unreadable_workbook", $"'{displayPath}' cannot be loaded: {e.Message}", e);
        }
    }
}
