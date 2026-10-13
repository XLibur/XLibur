using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;

namespace XLibur.Excel.IO;

/// <summary>
/// Leaves out of a worksheet what only restates a schema default, as Excel does (#709).
/// </summary>
/// <remarks>
/// <para>
/// An attribute whose value is its schema default is not written, and an element left with no
/// attribute and no child is not written either. Excel writes a new sheet with neither, and a sheet
/// that spells them out cannot be told from one that does not.
/// </para>
/// <para>
/// What a loaded file spelled out is kept. The writers update the element the load read, so an
/// attribute that already holds the value being written stays as the file had it, default or not,
/// and an element the file had empty stays empty. A sheet XLibur or ClosedXML wrote, which spells
/// out its defaults, then saves as it was loaded and can keep its part (#702).
/// </para>
/// </remarks>
internal static class SchemaDefault
{
    /// <summary>The value to set a boolean attribute to.</summary>
    /// <param name="loaded">The attribute as the element holds it now.</param>
    /// <param name="value">The value the model holds.</param>
    /// <param name="schemaDefault">The value the attribute means when it is absent.</param>
    /// <returns><paramref name="loaded"/> when it already holds <paramref name="value"/>, null when
    /// <paramref name="value"/> is the default, else a new attribute value.</returns>
    internal static BooleanValue? Bool(BooleanValue? loaded, bool value, bool schemaDefault)
    {
        if (loaded is { HasValue: true } && loaded.Value == value)
            return loaded;

        return value == schemaDefault ? null : new BooleanValue(value);
    }

    /// <inheritdoc cref="Bool"/>
    internal static UInt32Value? UInt(UInt32Value? loaded, uint value, uint schemaDefault)
    {
        if (loaded is { HasValue: true } && loaded.Value == value)
            return loaded;

        return value == schemaDefault ? null : new UInt32Value(value);
    }

    /// <inheritdoc cref="Bool"/>
    internal static EnumValue<T>? Enum<T>(EnumValue<T>? loaded, T value, T schemaDefault)
        where T : struct, IEnumValue, IEnumValueFactory<T>
    {
        if (loaded is { HasValue: true } && loaded.Value.Equals(value))
            return loaded;

        return value.Equals(schemaDefault) ? null : new EnumValue<T>(value);
    }

    /// <summary>The value to set a number attribute to.</summary>
    /// <param name="loaded">The attribute as the element holds it now.</param>
    /// <param name="value">The value the model holds.</param>
    /// <returns><paramref name="loaded"/> when it holds <paramref name="value"/>, else a new attribute
    /// value. A file can spell a number with more digits than it takes to tell it apart, such as
    /// <c>0.78749999999999998</c> for 0.7875, and that spelling stays while the number is the same
    /// (#722).</returns>
    internal static DoubleValue Double(DoubleValue? loaded, double value) =>
        loaded is { HasValue: true } && loaded.Value.Equals(value) ? loaded : new DoubleValue(value);

    /// <summary>
    /// The value to set a measurement such as a width or a height to, which a save writes to six
    /// decimals.
    /// </summary>
    /// <param name="loaded">The attribute as the element holds it now.</param>
    /// <param name="value">The value to write, already rounded to six decimals.</param>
    /// <returns><paramref name="loaded"/> when it rounds to <paramref name="value"/>, else a new
    /// attribute value. Excel writes a width with more digits than a save keeps, such as
    /// <c>14.85546875</c>, and the file's spelling stays while the width is the same (#722).</returns>
    internal static DoubleValue Measure(DoubleValue? loaded, double value) =>
        SameMeasure(loaded, value) ? loaded! : new DoubleValue(value);

    /// <summary>
    /// Does <paramref name="loaded"/> round to <paramref name="value"/>, the precision a save writes
    /// a measurement with?
    /// </summary>
    internal static bool SameMeasure(DoubleValue? loaded, double value) =>
        loaded is { HasValue: true } && Math.Abs(Math.Round(loaded.Value, 6) - value) < XLHelper.Epsilon;

    /// <summary>Does the element have neither an attribute nor a child?</summary>
    internal static bool IsEmpty(OpenXmlElement element) => !element.HasAttributes && !element.HasChildren;

    /// <summary>
    /// Puts a top-level element in its place in the worksheet, or takes it out when it is empty and
    /// the file did not have it empty.
    /// </summary>
    /// <param name="worksheet">The worksheet the element belongs in.</param>
    /// <param name="cm">Where each worksheet element is.</param>
    /// <param name="slot">The element's place in the worksheet.</param>
    /// <param name="element">The element, already in the worksheet when the load read it.</param>
    /// <param name="loadedEmpty">Did the file have the element, with nothing in it?</param>
    internal static void Place(Worksheet worksheet, XLWorksheetContentManager cm, XLWorksheetContents slot,
        OpenXmlElement element, bool loadedEmpty)
    {
        if (IsEmpty(element) && !loadedEmpty)
        {
            if (element.Parent is not null)
                element.Remove();

            cm.SetElement(slot, null);
            return;
        }

        if (element.Parent is null)
            worksheet.InsertAfter(element, cm.GetPreviousElementFor(slot));

        cm.SetElement(slot, element);
    }
}
