using System;
using System.IO;
using System.Reflection;

namespace XLibur.Graphics;

/// <summary>
/// Locates the default font engine at first use when no engine was registered explicitly.
/// </summary>
/// <remarks>
/// <para>
/// XLibur core has no compile-time dependency on any font library. To make a plain
/// <c>new XLWorkbook()</c> work with zero setup, this probe reflectively loads the default font
/// engine package (<c>XLibur.Fonts.SkiaSharp</c>) — the only assembly that is guaranteed to be
/// resolvable through the app's normal assembly-load path when the package is referenced.
/// </para>
/// <para>
/// A module initializer in the font package cannot be relied upon here: it only runs once the CLR
/// actually loads that assembly, which never happens for zero-config usage because the consumer only
/// touches core types. Probing from core — the always-loaded assembly — is what makes the default work.
/// </para>
/// <para>
/// The result is cached. The probe is consulted only after an explicit
/// <see cref="Excel.LoadOptions.FontEngine"/> / <see cref="Excel.LoadOptions.DefaultFontEngine"/> and
/// a workbook graphic engine, so any explicit registration always wins.
/// </para>
/// </remarks>
internal static class DefaultFontEngineProbe
{
    private const string AssemblyName = "XLibur.Fonts.SkiaSharp";
    private const string BootstrapTypeName = "XLibur.Fonts.SkiaSharp.SkiaSharpFontBootstrap";
    private const string FactoryMethodName = "CreateDefault";

    private static readonly object Gate = new();
    private static bool _probed;
    private static IXLFontEngine? _cached;

    /// <summary>
    /// Try to resolve the default font engine by reflectively loading the default font package.
    /// Returns <c>null</c> if the package is not present, in which case the caller surfaces a
    /// helpful error explaining how to register an engine.
    /// </summary>
    public static IXLFontEngine? TryResolveDefault()
    {
        if (_probed)
            return _cached;

        lock (Gate)
        {
            if (_probed)
                return _cached;

            _cached = Probe();
            _probed = true;
            return _cached;
        }
    }

    private static IXLFontEngine? Probe()
    {
        try
        {
            var assembly = Assembly.Load(new AssemblyName(AssemblyName));
            var bootstrapType = assembly.GetType(BootstrapTypeName, throwOnError: false);
            var factory = bootstrapType?.GetMethod(
                FactoryMethodName,
                BindingFlags.Public | BindingFlags.Static,
                binder: null,
                types: Type.EmptyTypes,
                modifiers: null);

            return factory?.Invoke(null, null) as IXLFontEngine;
        }
        catch (Exception e) when (e is FileNotFoundException
                                      or FileLoadException
                                      or BadImageFormatException
                                      or TargetInvocationException
                                      or MemberAccessException)
        {
            // The default font package is absent or could not be initialized. The caller reports a
            // clear, actionable error rather than surfacing a reflection exception here.
            return null;
        }
    }
}
