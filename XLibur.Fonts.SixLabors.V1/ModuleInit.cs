using System.Runtime.CompilerServices;
using XLibur.Graphics;

namespace XLibur.Fonts.SixLabors.V1;

internal static class ModuleInit
{
    [ModuleInitializer]
    internal static void Register()
    {
        FontEngineProvider.FromFallbackFont ??= fallbackFont => new DefaultFontEngine(fallbackFont);
        FontEngineProvider.FromStreams ??= (stream, useSystemFonts, fontStreams)
            => new DefaultFontEngine(stream, useSystemFonts, fontStreams);
    }
}
