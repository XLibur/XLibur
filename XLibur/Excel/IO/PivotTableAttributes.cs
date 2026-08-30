using System;
using System.Collections.Generic;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Extensions;

namespace XLibur.Excel.IO;

/// <summary>
/// One row of the pivot table definition attribute table: the OOXML attribute name, how its value
/// is read from a loaded file, written to a saved file, and copied from one pivot table to another.
/// </summary>
/// <remarks>
/// This is the single description backing three of the previous five hand-written enumerations of
/// the pivot table's ~60 settings (reader, writer, copy — the fourth, the Excel-defaults
/// initialiser, is now just the properties' own field initializers, and the fifth was the
/// hand-written round-trip test). A setting can no longer be present in the writer and missing from
/// copy, because both consume this same list. <see cref="XLPivotTable.DataPosition"/> and the
/// location element's row/column page counts are deliberately not rows here: both are derived from
/// other state and computed fresh on write rather than round-tripped.
/// </remarks>
internal sealed class PivotTableAttribute
{
    internal required string Name { get; init; }
    internal required Action<XmlWriter, XLPivotTable> Write { get; init; }
    internal required Action<PivotTableDefinition, XLPivotTable> Read { get; init; }
    internal required Action<XLPivotTable, XLPivotTable> Copy { get; init; }

    /// <summary>
    /// Test-only accessors: a boxed reader of the current value, and a setter that assigns a
    /// value distinct from this attribute's default. Used by the property-based round-trip test to
    /// iterate the table instead of hand-listing ~60 assertions.
    /// </summary>
    internal required Func<XLPivotTable, object?> GetValue { get; init; }

    internal required Action<XLPivotTable> SetNonDefault { get; init; }
}

internal static class PivotTableAttributes
{
    /// <summary>
    /// Every scalar attribute of the <c>pivotTableDefinition</c> root element that maps directly to
    /// an <see cref="XLPivotTable"/> property, in schema order. Excludes <c>name</c>/<c>cacheId</c>
    /// (identity, not a setting) and <c>dataPosition</c> (derived, see <see cref="PivotTableAttribute"/>
    /// remarks). The <c>pivotTableStyleInfo</c> child element's five flags are a separate, smaller
    /// group handled next to where that element is written/read/copied, because the whole element is
    /// conditionally omitted rather than each attribute independently defaulted.
    /// </summary>
    internal static readonly IReadOnlyList<PivotTableAttribute> All =
    [
        Bool("dataOnRows", pt => pt.DataOnRows, (pt, v) => pt.DataOnRows = v, false, s => s.DataOnRows),
        OptionalUInt("autoFormatId", pt => pt.AutoFormatId, (pt, v) => pt.AutoFormatId = v, s => s.AutoFormatId, 15u),
        Bool("applyNumberFormats", pt => pt.ApplyNumberFormats, (pt, v) => pt.ApplyNumberFormats = v, false, s => s.ApplyNumberFormats, alwaysWrite: true),
        Bool("applyBorderFormats", pt => pt.ApplyBorderFormats, (pt, v) => pt.ApplyBorderFormats = v, false, s => s.ApplyBorderFormats, alwaysWrite: true),
        Bool("applyFontFormats", pt => pt.ApplyFontFormats, (pt, v) => pt.ApplyFontFormats = v, false, s => s.ApplyFontFormats, alwaysWrite: true),
        Bool("applyPatternFormats", pt => pt.ApplyPatternFormats, (pt, v) => pt.ApplyPatternFormats = v, false, s => s.ApplyPatternFormats, alwaysWrite: true),
        Bool("applyAlignmentFormats", pt => pt.ApplyAlignmentFormats, (pt, v) => pt.ApplyAlignmentFormats = v, false, s => s.ApplyAlignmentFormats, alwaysWrite: true),
        Bool("applyWidthHeightFormats", pt => pt.ApplyWidthHeightFormats, (pt, v) => pt.ApplyWidthHeightFormats = v, false, s => s.ApplyWidthHeightFormats, alwaysWrite: true),
        RequiredString("dataCaption", pt => pt.DataCaption, (pt, v) => pt.DataCaption = v, s => s.DataCaption, "Test values"),
        OptionalString("grandTotalCaption", pt => pt.GrandTotalCaption, (pt, v) => pt.GrandTotalCaption = v, s => s.GrandTotalCaption),
        OptionalString("errorCaption", pt => pt.ErrorValueReplacement, (pt, v) => pt.ErrorValueReplacement = v, s => s.ErrorCaption),
        Bool("showError", pt => pt.ShowError, (pt, v) => pt.ShowError = v, false, s => s.ShowError),
        OptionalString("missingCaption", pt => pt.MissingCaption, (pt, v) => pt.MissingCaption = v ?? string.Empty, s => s.MissingCaption),
        Bool("showMissing", pt => pt.ShowMissing, (pt, v) => pt.ShowMissing = v, true, s => s.ShowMissing),
        OptionalString("pageStyle", pt => pt.PageStyle, (pt, v) => pt.PageStyle = v, s => s.PageStyle),
        OptionalString("pivotTableStyle", pt => pt.PivotTableStyleName, (pt, v) => pt.PivotTableStyleName = v, s => s.PivotTableStyleName),
        OptionalString("vacatedStyle", pt => pt.VacatedStyle, (pt, v) => pt.VacatedStyle = v, s => s.VacatedStyle),
        OptionalString("tag", pt => pt.Tag, (pt, v) => pt.Tag = v, s => s.Tag),
        Byte("updatedVersion", pt => pt.UpdatedVersion, (pt, v) => pt.UpdatedVersion = v, 0, s => s.UpdatedVersion),
        Byte("minRefreshableVersion", pt => pt.MinRefreshableVersion, (pt, v) => pt.MinRefreshableVersion = v, 0, s => s.MinRefreshableVersion),
        Bool("asteriskTotals", pt => pt.AsteriskTotals, (pt, v) => pt.AsteriskTotals = v, false, s => s.AsteriskTotals),
        Bool("showItems", pt => pt.DisplayItemLabels, (pt, v) => pt.DisplayItemLabels = v, true, s => s.ShowItems),
        Bool("editData", pt => pt.EditData, (pt, v) => pt.EditData = v, false, s => s.EditData),
        Bool("disableFieldList", pt => pt.DisableFieldList, (pt, v) => pt.DisableFieldList = v, false, s => s.DisableFieldList),
        Bool("showCalcMbrs", pt => pt.ShowCalculatedMembers, (pt, v) => pt.ShowCalculatedMembers = v, true, s => s.ShowCalculatedMembers),
        Bool("visualTotals", pt => pt.VisualTotals, (pt, v) => pt.VisualTotals = v, true, s => s.VisualTotals),
        Bool("showMultipleLabel", pt => pt.ShowMultipleLabel, (pt, v) => pt.ShowMultipleLabel = v, true, s => s.ShowMultipleLabel),
        Bool("showDataDropDown", pt => pt.ShowDataDropDown, (pt, v) => pt.ShowDataDropDown = v, true, s => s.ShowDataDropDown),
        Bool("showDrill", pt => pt.ShowExpandCollapseButtons, (pt, v) => pt.ShowExpandCollapseButtons = v, true, s => s.ShowDrill),
        Bool("printDrill", pt => pt.PrintExpandCollapsedButtons, (pt, v) => pt.PrintExpandCollapsedButtons = v, false, s => s.PrintDrill),
        Bool("showMemberPropertyTips", pt => pt.ShowPropertiesInTooltips, (pt, v) => pt.ShowPropertiesInTooltips = v, true, s => s.ShowMemberPropertyTips),
        Bool("showDataTips", pt => pt.ShowContextualTooltips, (pt, v) => pt.ShowContextualTooltips = v, true, s => s.ShowDataTips),
        Bool("enableWizard", pt => pt.EnableEditingMechanism, (pt, v) => pt.EnableEditingMechanism = v, true, s => s.EnableWizard),
        Bool("enableDrill", pt => pt.EnableShowDetails, (pt, v) => pt.EnableShowDetails = v, true, s => s.EnableDrill),
        Bool("enableFieldProperties", pt => pt.EnableFieldProperties, (pt, v) => pt.EnableFieldProperties = v, true, s => s.EnableFieldProperties),
        Bool("preserveFormatting", pt => pt.PreserveCellFormatting, (pt, v) => pt.PreserveCellFormatting = v, true, s => s.PreserveFormatting),
        Bool("useAutoFormatting", pt => pt.AutofitColumns, (pt, v) => pt.AutofitColumns = v, false, s => s.UseAutoFormatting),
        IntAsUInt("pageWrap", pt => pt.FilterFieldsPageWrap, (pt, v) => pt.FilterFieldsPageWrap = v, 0, s => s.PageWrap),
        new PivotTableAttribute
        {
            Name = "pageOverThenDown",
            Write = (xml, pt) => xml.WriteAttributeDefault("pageOverThenDown", pt.FilterAreaOrder == XLFilterAreaOrder.OverThenDown, false),
            Read = (src, pt) => pt.FilterAreaOrder = (src.PageOverThenDown?.Value ?? false)
                ? XLFilterAreaOrder.OverThenDown
                : XLFilterAreaOrder.DownThenOver,
            Copy = (s, t) => t.FilterAreaOrder = s.FilterAreaOrder,
            GetValue = pt => pt.FilterAreaOrder == XLFilterAreaOrder.OverThenDown,
            SetNonDefault = pt => pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown,
        },
        Bool("subtotalHiddenItems", pt => pt.FilteredItemsInSubtotals, (pt, v) => pt.FilteredItemsInSubtotals = v, false, s => s.SubtotalHiddenItems),
        Bool("rowGrandTotals", pt => pt.ShowGrandTotalsRows, (pt, v) => pt.ShowGrandTotalsRows = v, true, s => s.RowGrandTotals),
        Bool("colGrandTotals", pt => pt.ShowGrandTotalsColumns, (pt, v) => pt.ShowGrandTotalsColumns = v, true, s => s.ColumnGrandTotals),
        Bool("fieldPrintTitles", pt => pt.PrintTitles, (pt, v) => pt.PrintTitles = v, false, s => s.FieldPrintTitles),
        Bool("itemPrintTitles", pt => pt.RepeatRowLabels, (pt, v) => pt.RepeatRowLabels = v, false, s => s.ItemPrintTitles),
        Bool("mergeItem", pt => pt.MergeAndCenterWithLabels, (pt, v) => pt.MergeAndCenterWithLabels = v, false, s => s.MergeItem),
        Bool("showDropZones", pt => pt.ShowDropZones, (pt, v) => pt.ShowDropZones = v, true, s => s.ShowDropZones),
        Byte("createdVersion", pt => pt.PivotCacheCreatedVersion, (pt, v) => pt.PivotCacheCreatedVersion = v, 0, s => s.CreatedVersion),
        IntAsUInt("indent", pt => pt.RowLabelIndent, (pt, v) => pt.RowLabelIndent = v, 1, s => s.Indent),
        Bool("showEmptyRow", pt => pt.ShowEmptyItemsOnRows, (pt, v) => pt.ShowEmptyItemsOnRows = v, false, s => s.ShowEmptyRow),
        Bool("showEmptyCol", pt => pt.ShowEmptyItemsOnColumns, (pt, v) => pt.ShowEmptyItemsOnColumns = v, false, s => s.ShowEmptyColumn),
        Bool("showHeaders", pt => pt.DisplayCaptionsAndDropdowns, (pt, v) => pt.DisplayCaptionsAndDropdowns = v, true, s => s.ShowHeaders),
        Bool("compact", pt => pt.Compact, (pt, v) => pt.Compact = v, true, s => s.Compact),
        Bool("outline", pt => pt.Outline, (pt, v) => pt.Outline = v, false, s => s.Outline),
        Bool("outlineData", pt => pt.OutlineData, (pt, v) => pt.OutlineData = v, false, s => s.OutlineData),
        Bool("compactData", pt => pt.CompactData, (pt, v) => pt.CompactData = v, true, s => s.CompactData),
        Bool("published", pt => pt.Published, (pt, v) => pt.Published = v, false, s => s.Published),
        Bool("gridDropZones", pt => pt.ClassicPivotTableLayout, (pt, v) => pt.ClassicPivotTableLayout = v, false, s => s.GridDropZones),
        Bool("immersive", pt => pt.StopImmersiveUi, (pt, v) => pt.StopImmersiveUi = v, true, s => s.StopImmersiveUi),
        Bool("multipleFieldFilters", pt => pt.AllowMultipleFilters, (pt, v) => pt.AllowMultipleFilters = v, true, s => s.MultipleFieldFilters),
        UInt("chartFormat", pt => pt.ChartFormat, (pt, v) => pt.ChartFormat = v, 0, s => s.ChartFormat),
        OptionalString("rowHeaderCaption", pt => pt.RowHeaderCaption, (pt, v) => pt.RowHeaderCaption = v, s => s.RowHeaderCaption),
        OptionalString("colHeaderCaption", pt => pt.ColumnHeaderCaption, (pt, v) => pt.ColumnHeaderCaption = v, s => s.ColumnHeaderCaption),
        Bool("fieldListSortAscending", pt => pt.SortFieldsAtoZ, (pt, v) => pt.SortFieldsAtoZ = v, false, s => s.FieldListSortAscending),
        Bool("mdxSubqueries", pt => pt.MdxSubQueries, (pt, v) => pt.MdxSubQueries = v, false, s => s.MdxSubqueries),
        Bool("customListSort", pt => pt.UseCustomListsForSorting, (pt, v) => pt.UseCustomListsForSorting = v, true, s => s.CustomListSort),
    ];

    private static PivotTableAttribute Bool(
        string name,
        Func<XLPivotTable, bool> get, Action<XLPivotTable, bool> set,
        bool defaultValue,
        Func<PivotTableDefinition, BooleanValue?> source,
        bool alwaysWrite = false)
        => new()
        {
            Name = name,
            Write = (xml, pt) =>
            {
                if (alwaysWrite)
                    xml.WriteAttribute(name, get(pt));
                else
                    xml.WriteAttributeDefault(name, get(pt), defaultValue);
            },
            Read = (src, pt) => set(pt, source(src)?.Value ?? defaultValue),
            Copy = (s, t) => set(t, get(s)),
            GetValue = pt => get(pt),
            SetNonDefault = pt => set(pt, !defaultValue),
        };

    private static PivotTableAttribute OptionalString(
        string name,
        Func<XLPivotTable, string?> get, Action<XLPivotTable, string?> set,
        Func<PivotTableDefinition, StringValue?> source)
        => new()
        {
            Name = name,
            Write = (xml, pt) => xml.WriteAttributeOptional(name, get(pt)),
            Read = (src, pt) => set(pt, source(src)?.Value),
            Copy = (s, t) => set(t, get(s)),
            GetValue = pt => get(pt),
            SetNonDefault = pt => set(pt, $"Test {name}"),
        };

    private static PivotTableAttribute RequiredString(
        string name,
        Func<XLPivotTable, string> get, Action<XLPivotTable, string> set,
        Func<PivotTableDefinition, StringValue?> source,
        string nonDefaultValue)
        => new()
        {
            Name = name,
            Write = (xml, pt) => xml.WriteAttribute(name, get(pt)),
            Read = (src, pt) => set(pt, source(src)?.Value ?? throw PartStructureException.MissingAttribute()),
            Copy = (s, t) => set(t, get(s)),
            GetValue = pt => get(pt),
            SetNonDefault = pt => set(pt, nonDefaultValue),
        };

    private static PivotTableAttribute OptionalUInt(
        string name,
        Func<XLPivotTable, uint?> get, Action<XLPivotTable, uint?> set,
        Func<PivotTableDefinition, UInt32Value?> source,
        uint nonDefaultValue)
        => new()
        {
            Name = name,
            Write = (xml, pt) => xml.WriteAttributeOptional(name, get(pt)),
            Read = (src, pt) => set(pt, source(src)?.Value),
            Copy = (s, t) => set(t, get(s)),
            GetValue = pt => get(pt),
            SetNonDefault = pt => set(pt, nonDefaultValue),
        };

    private static PivotTableAttribute UInt(
        string name,
        Func<XLPivotTable, uint> get, Action<XLPivotTable, uint> set,
        uint defaultValue,
        Func<PivotTableDefinition, UInt32Value?> source)
        => new()
        {
            Name = name,
            Write = (xml, pt) => xml.WriteAttributeDefault(name, get(pt), defaultValue),
            Read = (src, pt) => set(pt, source(src)?.Value ?? defaultValue),
            Copy = (s, t) => set(t, get(s)),
            GetValue = pt => get(pt),
            SetNonDefault = pt => set(pt, defaultValue + 7),
        };

    private static PivotTableAttribute Byte(
        string name,
        Func<XLPivotTable, byte> get, Action<XLPivotTable, byte> set,
        byte defaultValue,
        Func<PivotTableDefinition, ByteValue?> source)
        => new()
        {
            Name = name,
            Write = (xml, pt) => xml.WriteAttributeDefault(name, (int)get(pt), defaultValue),
            Read = (src, pt) => set(pt, source(src)?.Value ?? defaultValue),
            Copy = (s, t) => set(t, get(s)),
            GetValue = pt => get(pt),
            SetNonDefault = pt => set(pt, (byte)(defaultValue + 9)),
        };

    /// <summary>
    /// A pivot table property stored as <c>int</c> (for range-checked setters), written on the
    /// wire as an unsigned schema value (<c>xsd:unsignedInt</c>).
    /// </summary>
    private static PivotTableAttribute IntAsUInt(
        string name,
        Func<XLPivotTable, int> get, Action<XLPivotTable, int> set,
        int defaultValue,
        Func<PivotTableDefinition, UInt32Value?> source)
        => new()
        {
            Name = name,
            Write = (xml, pt) => xml.WriteAttributeDefault(name, checked((uint)get(pt)), checked((uint)defaultValue)),
            Read = (src, pt) => set(pt, checked((int)(source(src)?.Value ?? (uint)defaultValue))),
            Copy = (s, t) => set(t, get(s)),
            GetValue = pt => get(pt),
            SetNonDefault = pt => set(pt, defaultValue + 7),
        };
}
