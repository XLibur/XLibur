using System;
using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// One entry of the sheet-view property table: a name, the OOXML attribute it maps to (or the
/// element it lives in, for the one property that is not on <c>sheetView</c> itself), the value
/// the attribute takes when unset, how that default is polarised against the omit-when-default
/// rule, a way to push the property to a fixed non-default value through <see cref="IXLWorksheet"/>,
/// and a way to read it back the same way.
/// </summary>
/// <remarks>
/// This is the list spec 38 promotes to data: copying and default-seeding are driven from the
/// matching constructors on <see cref="XLSheetView"/>, and this table is what lets a test iterate
/// every view property instead of restating them by hand — see the property-enumerating tests in
/// <c>XLSheetViewTests</c>.
/// </remarks>
internal readonly record struct XLViewProperty(
    string Name,
    string OoxmlAttribute,
    string Default,
    string Polarity,
    Action<IXLWorksheet> SetNonDefault,
    Func<IXLWorksheet, object> Get,
    // Whether copying a sheet is expected to carry this property onto the copy: true for every
    // property that describes how the sheet *looks*, false for TabSelected alone, which describes
    // which tab the user is on and must not be duplicated onto a second sheet.
    bool SurvivesCopy = true);

internal static class XLViewProperties
{
    /// <summary>
    /// Every sheet-view property spec 38 covers, in an order that is safe to apply top to bottom:
    /// <see cref="IXLSheetView.View"/> before <see cref="IXLSheetView.ZoomScale"/> (whose setter
    /// mirrors onto whichever of the three named scales matches the current view), and
    /// <see cref="IXLSheetView.ZoomScale"/> before the three named scales (whose values must win
    /// over that mirroring). Panes and splits are out of scope for this spec and are not listed
    /// here; their coverage is <c>WritePathAgreementTests</c> and the existing
    /// <c>CopyWorksheetSheetViews</c> test.
    /// </summary>
    internal static readonly IReadOnlyList<XLViewProperty> All =
    [
        new(
            "ShowFormulas", "showFormulas", "false",
            "written (true) when non-default; omitted when default",
            ws => ws.ShowFormulas = true, ws => ws.ShowFormulas),
        new(
            "ShowGridLines", "showGridLines", "true",
            "omitted when default; written (false) when non-default",
            ws => ws.ShowGridLines = false, ws => ws.ShowGridLines),
        new(
            "ShowOutlineSymbols", "showOutlineSymbols", "true",
            "omitted when default; written (false) when non-default",
            ws => ws.ShowOutlineSymbols = false, ws => ws.ShowOutlineSymbols),
        new(
            "ShowRowColHeaders", "showRowColHeaders", "true",
            "omitted when default; written (false) when non-default",
            ws => ws.ShowRowColHeaders = false, ws => ws.ShowRowColHeaders),
        new(
            "ShowRuler", "showRuler", "true",
            "omitted when default; written (false) when non-default",
            ws => ws.ShowRuler = false, ws => ws.ShowRuler),
        new(
            "ShowWhiteSpace", "showWhiteSpace", "true",
            "omitted when default; written (false) when non-default",
            ws => ws.ShowWhiteSpace = false, ws => ws.ShowWhiteSpace),
        new(
            "ShowZeros", "showZeros", "true",
            "omitted when default; written (false) when non-default",
            ws => ws.ShowZeros = false, ws => ws.ShowZeros),
        new(
            "RightToLeft", "rightToLeft", "false",
            "written (true) when non-default; omitted when default",
            ws => ws.RightToLeft = true, ws => ws.RightToLeft),
        new(
            "TabSelected", "tabSelected", "false",
            "written (true) when non-default; omitted when default",
            ws => ws.TabSelected = true, ws => ws.TabSelected,
            SurvivesCopy: false),
        new(
            "TabColor", "sheetPr/tabColor (not on sheetView)", "no colour (XLColor.Automatic)",
            "omitted when the colour has no value; written when it does",
            ws => ws.TabColor = XLColor.FromArgb(200, 30, 60), ws => ws.TabColor),
        new(
            "View", "view", "normal",
            "omitted when Normal; written otherwise",
            ws => ws.SheetView.SetView(XLSheetViewOptions.PageLayout), ws => ws.SheetView.View),
        new(
            "ZoomScale", "zoomScale", "100",
            "omitted when 100; written (clamped 10-400) otherwise",
            ws => ws.SheetView.ZoomScale = 120, ws => ws.SheetView.ZoomScale),
        new(
            "ZoomScaleNormal", "zoomScaleNormal", "100 (XLibur sentinel; OOXML default is 0/automatic)",
            "omitted when 100; written otherwise",
            ws => ws.SheetView.ZoomScaleNormal = 130, ws => ws.SheetView.ZoomScaleNormal),
        new(
            "ZoomScalePageLayoutView", "zoomScalePageLayoutView",
            "100 (XLibur sentinel; OOXML default is 0/automatic)",
            "omitted when 100; written otherwise",
            ws => ws.SheetView.ZoomScalePageLayoutView = 140, ws => ws.SheetView.ZoomScalePageLayoutView),
        new(
            "ZoomScaleSheetLayoutView", "zoomScaleSheetLayoutView",
            "100 (XLibur sentinel; OOXML default is 0/automatic)",
            "omitted when 100; written otherwise",
            ws => ws.SheetView.ZoomScaleSheetLayoutView = 150, ws => ws.SheetView.ZoomScaleSheetLayoutView),
    ];
}
