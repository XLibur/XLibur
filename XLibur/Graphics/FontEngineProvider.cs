using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Loader;

namespace XLibur.Graphics;

/// <summary>
/// Registration point for font engine factories. The default font engine package
/// (e.g., XLibur.Fonts.SixLabors.V1) registers its factories here via a module initializer
/// so that <see cref="DefaultGraphicEngine"/> can create font engines without a direct assembly reference.
/// </summary>
public static class FontEngineProvider
{
    private const string V1AssemblyName = "XLibur.Fonts.SixLabors.V1";
    private static readonly object _lock = new();
    private static bool _discoveryAttempted;

    /// <summary>
    /// Factory that creates an <see cref="IXLFontEngine"/> from a fallback font name.
    /// </summary>
    public static Func<string, IXLFontEngine>? FromFallbackFont { get; set; }

    /// <summary>
    /// Factory that creates an <see cref="IXLFontEngine"/> from font streams.
    /// </summary>
    /// <remarks>Parameters: fallbackFontStream, useSystemFonts, additional fontStreams.</remarks>
    public static Func<Stream, bool, Stream[], IXLFontEngine>? FromStreams { get; set; }

    internal static IXLFontEngine CreateFromFallbackFont(string fallbackFont)
    {
        EnsureProviderLoaded();
        return FromFallbackFont?.Invoke(fallbackFont)
               ?? throw new InvalidOperationException(
                   "No font engine provider is registered. Install the XLibur.Fonts.SixLabors.V1 NuGet package, " +
                   "or set FontEngineProvider.FromFallbackFont / LoadOptions.FontEngine manually.");
    }

    internal static IXLFontEngine CreateFromStreams(Stream fallbackFontStream, bool useSystemFonts, Stream[] fontStreams)
    {
        EnsureProviderLoaded();
        return FromStreams?.Invoke(fallbackFontStream, useSystemFonts, fontStreams)
               ?? throw new InvalidOperationException(
                   "No font engine provider is registered. Install the XLibur.Fonts.SixLabors.V1 NuGet package, " +
                   "or set FontEngineProvider.FromStreams / LoadOptions.FontEngine manually.");
    }

    /// <summary>
    /// Attempt to load the default font engine assembly if no provider has been registered yet.
    /// Loading the assembly triggers its module initializer which registers the factories.
    /// </summary>
    private static void EnsureProviderLoaded()
    {
        if (FromFallbackFont is not null)
            return;

        lock (_lock)
        {
            if (_discoveryAttempted)
                return;

            _discoveryAttempted = true;

            // Check if V1 assembly is already loaded (e.g., by the host application)
            if (AppDomain.CurrentDomain.GetAssemblies()
                .Any(a => a.GetName().Name == V1AssemblyName))
                return;

            // Try to load V1 from known probing paths
            var probePaths = new[]
            {
                Path.GetDirectoryName(typeof(FontEngineProvider).Assembly.Location),
                AppContext.BaseDirectory
            };

            foreach (var dir in probePaths)
            {
                if (string.IsNullOrEmpty(dir))
                    continue;

                var dllPath = Path.Combine(dir, V1AssemblyName + ".dll");
                if (!File.Exists(dllPath))
                    continue;

                try
                {
                    var asm = AssemblyLoadContext.Default.LoadFromAssemblyPath(dllPath);
                    // Explicitly run the module constructor (.cctor) which is where
                    // [ModuleInitializer] methods are compiled to
                    foreach (var module in asm.GetModules())
                        RuntimeHelpers.RunModuleConstructor(module.ModuleHandle);

                    if (FromFallbackFont is not null)
                        return;
                }
                catch
                {
                    // Continue to next probe path
                }
            }
        }
    }
}
