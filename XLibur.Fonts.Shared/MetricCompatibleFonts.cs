using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Fonts;

/// <summary>
/// The fonts that measure like a font the machine may not have, tried before the engine's
/// fallback font.
/// </summary>
/// <remarks>
/// <para>
/// Linked into each font engine as source, so the engines share one list without a public API.
/// </para>
/// <para>
/// Excel's default font in Microsoft 365 is Aptos Narrow. Windows does not install it: Office
/// downloads it into its own font cache, which other programs do not see, so an engine resolved it
/// to its fallback font, Microsoft Sans Serif. Those digits are about 10% wider, and every column
/// width came out wrong: a default column was 72 px at 96 DPI instead of Excel's 64 px (#793). The
/// digits of Aptos Narrow are 1038/2048 of an em wide, the same as those of Calibri and its
/// metric-compatible clone Carlito, so either one gives the digit width Excel uses.
/// </para>
/// </remarks>
internal static class MetricCompatibleFonts
{
    private static readonly Dictionary<string, string[]> Substitutes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Aptos Narrow"] = ["Calibri", "Carlito", "CarlitoBare"],
    };

    /// <summary>
    /// The families to try, in order, when <paramref name="familyName"/> is not available. Empty
    /// when no family measures like it.
    /// </summary>
    internal static IReadOnlyList<string> For(string familyName) =>
        Substitutes.TryGetValue(familyName, out var substitutes) ? substitutes : [];

    /// <summary>
    /// The first of the families <see cref="For"/> gives that <paramref name="tryGetFamily"/>, the
    /// engine's own lookup, finds.
    /// </summary>
    internal static bool TryGetSubstitute<TFamily>(string familyName, TryGetFamily<TFamily> tryGetFamily,
        [MaybeNullWhen(false)] out TFamily family)
    {
        foreach (var substitute in For(familyName))
        {
            if (tryGetFamily(substitute, out family))
                return true;
        }

        family = default;
        return false;
    }

    /// <summary>An engine's lookup of a font family by its name.</summary>
    internal delegate bool TryGetFamily<TFamily>(string familyName, [MaybeNullWhen(false)] out TFamily family);
}
