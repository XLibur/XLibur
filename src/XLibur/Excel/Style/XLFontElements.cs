using System;

namespace XLibur.Excel;

/// <summary>
/// The children a loaded <c>&lt;font&gt;</c> states. A font that leaves a child out takes its
/// default, so two fonts can decode to one <see cref="XLFontValue"/> and still differ in what they
/// say: one states <c>&lt;family val="2"/&gt;</c>, the other leaves the family to whatever the
/// reader assumes. A style records the elements of the font it was loaded with, so the save writes
/// it to that font and not to the first one that decodes the same (#978).
/// </summary>
/// <remarks>
/// <see cref="None"/> is a font made through the API, or a loaded font no other loaded font shares a
/// value with, and matches any font of its value. A font whose elements are recorded always has
/// <see cref="Loaded"/>, even one that states no child at all.
/// </remarks>
[Flags]
internal enum XLFontElements
{
    None = 0,
    Loaded = 1 << 0,
    Bold = 1 << 1,
    Italic = 1 << 2,
    Underline = 1 << 3,
    Strikethrough = 1 << 4,
    VerticalAlignment = 1 << 5,
    Shadow = 1 << 6,
    FontSize = 1 << 7,
    FontColor = 1 << 8,
    FontName = 1 << 9,
    FontFamilyNumbering = 1 << 10,
    FontCharSet = 1 << 11,
    FontScheme = 1 << 12,
}
