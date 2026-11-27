using System;

namespace XLibur.Excel.RichText;

/// <summary>
/// Which of a run's <c>&lt;rPr&gt;</c> properties the source actually stated, for those whose value
/// cannot be told apart from the one the run would take anyway, or from the default. A run whose
/// <c>&lt;rPr&gt;</c> omits <c>&lt;vertAlign&gt;</c> or <c>&lt;family&gt;</c> inherits the cell font's value for it, so writing
/// that value back out would turn an inherited default into formatting the source never asked for -
/// the same trap as an explicit black for an automatic colour (issue #219).
/// <para>
/// A run created through the public API states all of its formatting, and so does a loaded run once
/// its font is edited, which is why <see cref="All"/> is the default rather than
/// <see cref="None"/>.
/// </para>
/// </summary>
[Flags]
internal enum XLStatedRunProperties
{
    None = 0,

    /// <summary>The run stated <c>&lt;vertAlign&gt;</c>.</summary>
    VerticalAlignment = 1 << 0,

    /// <summary>The run stated <c>&lt;family&gt;</c>.</summary>
    FontFamilyNumbering = 1 << 1,

    /// <summary>
    /// The run stated <c>&lt;charset&gt;</c>. A charset other than the default is written anyway;
    /// this only keeps an explicit <c>&lt;charset val="1"/&gt;</c>, which Excel writes on every note
    /// run (#957). It is not part of <see cref="All"/>, so a run built or edited through the API
    /// still omits the default charset.
    /// </summary>
    CharSet = 1 << 2,

    All = VerticalAlignment | FontFamilyNumbering
}
