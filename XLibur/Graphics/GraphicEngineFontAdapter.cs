using System;
using XLibur.Excel;

namespace XLibur.Graphics;

/// <summary>
/// Adapts a legacy <see cref="IXLGraphicEngine"/> that does not implement <see cref="IXLFontEngine"/>
/// by delegating the font measurement methods to the graphic engine's identically-signatured methods.
/// </summary>
internal sealed class GraphicEngineFontAdapter(IXLGraphicEngine engine) : IXLFontEngine
{
    public double GetTextHeight(IXLFontBase font, double dpiY)
        => engine.GetTextHeight(font, dpiY);

    public double GetTextWidth(string text, IXLFontBase font, double dpiX)
        => engine.GetTextWidth(text, font, dpiX);

    public double GetMaxDigitWidth(IXLFontBase font, double dpiX)
        => engine.GetMaxDigitWidth(font, dpiX);

    public double GetDescent(IXLFontBase font, double dpiY)
        => engine.GetDescent(font, dpiY);

    public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
        => engine.GetGlyphBox(graphemeCluster, font, dpi);
}
