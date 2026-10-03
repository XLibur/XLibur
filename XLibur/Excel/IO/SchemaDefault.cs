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
