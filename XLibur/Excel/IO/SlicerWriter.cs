using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Excel.IO.DrawingML;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the worksheet half of a slicer: the slicers part holding its definition and the
/// <c>extLst</c> reference that makes the worksheet point at it.
/// </summary>
/// <remarks>
/// <para>
/// A surviving slicer part is worthless if the sheet stops referencing it — the same trap the form
/// controls in <c>docs/round-trip-fidelity.md</c> illustrate, and one that bites here because the
/// worksheet part is rebuilt from the model on every save while the slicer part is not.
/// </para>
/// <para>
/// Excel uses a different extension URI for each kind of slicer, and writes one slicers part per
/// slicer rather than one per sheet, so a sheet carrying a table slicer and two pivot slicers has
/// three parts listed under two extensions. That layout is reproduced rather than flattened — see
/// <see cref="NewSlicersPart"/> for what happens when it is not.
/// </para>
/// </remarks>
internal static class SlicerWriter
{
    /// <summary>The worksheet extension holding the list of pivot slicers on the sheet.</summary>
    private const string PivotSlicerExtensionUri = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}";

    /// <summary>The worksheet extension holding the list of table slicers on the sheet.</summary>
    private const string TableSlicerExtensionUri = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}";

    internal static void WriteSlicers(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        WorksheetPart worksheetPart,
        SaveContext context)
    {
        var slicers = xlWorksheet.SlicersInternal;

        RemoveDeletedSlicers(worksheet, cm, worksheetPart, slicers);

        foreach (var slicer in slicers.Items)
        {
            if (!slicer.IsNew)
            {
                // A slicer that already exists in the package is never regenerated — that is what
                // keeps the parts of its XML XLibur does not model intact. Only the properties the
                // caller actually changed are patched into the existing part.
                SlicerPatcher.PatchSlicer(worksheetPart, slicer);
                continue;
            }

            WriteNewSlicer(worksheet, cm, worksheetPart, slicer, context);
            slicer.IsNew = false;
        }
    }

    private static void WriteNewSlicer(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLSlicer xlSlicer,
        SaveContext context)
    {
        var part = NewSlicersPart(worksheetPart, context, out var relId);
        xlSlicer.PartRelId = relId;

        var slicersRoot = part.Slicers = NewSlicersRoot();

        var slicer = new X14.Slicer
        {
            Name = xlSlicer.Name,
            Cache = xlSlicer.Cache.Name,
            Caption = xlSlicer.Caption,
        };

        // Attributes at their schema default are left off, which is what Excel writes and what
        // keeps a generated part comparable with a hand-made one.
        if (!xlSlicer.ShowCaption)
            slicer.ShowCaption = false;

        if (xlSlicer.Style is { } style)
            slicer.Style = style;

        if (xlSlicer.ColumnCount != 1)
            slicer.ColumnCount = xlSlicer.ColumnCount;

        // rowHeight is required rather than optional, so it is always written — falling back to
        // Excel's own default for a slicer that has none.
        var rowHeightPt = xlSlicer.RowHeightPt ?? XLSlicer.DefaultRowHeightPt;
        slicer.RowHeight = (uint)System.Math.Round(
            rowHeightPt * DrawingUnits.EmuPerPoint, System.MidpointRounding.AwayFromZero);

        slicersRoot.AppendChild(slicer);

        EnsureSlicerListReference(worksheet, cm, xlSlicer.SourceKind, relId);
        WriteAnchor(worksheet, cm, worksheetPart, xlSlicer, context);
    }

    /// <summary>
    /// Draws the slicer: the graphic frame in the sheet's drawing part, and the sheet's reference to
    /// that part.
    /// </summary>
    /// <remarks>
    /// The sixth of the six pieces a created slicer needs. Without it the workbook opens and the
    /// slicer is simply not there, because <c>xl/slicers/slicerN.xml</c> says what a slicer filters
    /// but never where it sits.
    /// </remarks>
    private static void WriteAnchor(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLSlicer xlSlicer,
        SaveContext context)
    {
        var drawingsPart = DrawingPartScaffold.EnsureDrawingsPart(worksheetPart, context);
        var worksheetDrawing = drawingsPart.WorksheetDrawing!;
        DrawingPartScaffold.EnsureNamespaces(worksheetDrawing);

        SlicerAnchorXml.Append(worksheetDrawing, xlSlicer);

        DrawingPartScaffold.EnsureDrawingElement(worksheet, cm, worksheetPart, drawingsPart);
    }

    // ── The slicers part ────────────────────────────────────────────────

    /// <summary>
    /// A slicers part of its own for a newly created slicer.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>A created slicer never goes into a part that is already in the package.</b> An earlier
    /// version appended the new definition to the sheet's existing <c>xl/slicers/slicerN.xml</c>,
    /// which meant opening a part Excel had written and handing the SDK the job of serialising it
    /// again. Every automated gate passed — nothing was dropped from the XML, the validator was
    /// clean, both slicers reloaded — and Excel then silently stopped drawing the slicer that was
    /// already there. That is the exact failure the first of the four mechanisms in
    /// <c>docs/round-trip-fidelity.md</c> exists to prevent: a slicer part survives a round trip
    /// only because nothing ever opens it.
    /// </para>
    /// <para>
    /// One part per slicer also matches every slicers part Excel itself writes, each of which holds
    /// exactly one <c>x14:slicer</c>, and it is the shape the manual acceptance check confirmed
    /// working. The sheet's <c>slicerList</c> is a list precisely so it can name several parts.
    /// </para>
    /// </remarks>
    private static SlicersPart NewSlicersPart(
        WorksheetPart worksheetPart, SaveContext context, out string relId)
    {
        relId = context.RelIdGenerator.GetNext(RelType.Workbook);
        return worksheetPart.AddNewPart<SlicersPart>(relId);
    }

    private static X14.Slicers NewSlicersRoot()
    {
        var slicers = new X14.Slicers();
        slicers.AddNamespaceDeclaration("x", Main2006SsNs);
        return slicers;
    }

    // ── The worksheet reference ─────────────────────────────────────────

    /// <summary>
    /// Makes the worksheet's <c>extLst</c> point at the slicers part, under the URI Excel uses for
    /// this kind of slicer.
    /// </summary>
    private static void EnsureSlicerListReference(
        Worksheet worksheet, XLWorksheetContentManager cm, XLSlicerSourceKind kind, string relId)
    {
        var uri = ExtensionUri(kind);
        var extension = FindExtension(worksheet, uri);

        if (extension is null)
        {
            if (!worksheet.Elements<WorksheetExtensionList>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.WorksheetExtensionList);
                worksheet.InsertAfter(new WorksheetExtensionList(), previousElement);
            }

            var extensionList = worksheet.Elements<WorksheetExtensionList>().First();
            cm.SetElement(XLWorksheetContents.WorksheetExtensionList, extensionList);

            extension = new WorksheetExtension { Uri = uri };
            extension.AddNamespaceDeclaration("x14", X14Main2009SsNs);
            extension.AppendChild(new X14.SlicerList());
            extensionList.AppendChild(extension);
        }

        var slicerList = extension.GetFirstChild<X14.SlicerList>();
        if (slicerList is null)
        {
            slicerList = new X14.SlicerList();
            extension.AppendChild(slicerList);
        }

        if (!slicerList.Elements<X14.SlicerRef>().Any(r => r.Id?.Value == relId))
            slicerList.AppendChild(new X14.SlicerRef { Id = relId });
    }

    // ── Removal ─────────────────────────────────────────────────────────

    /// <summary>
    /// Unpicks the worksheet half of every slicer removed since the workbook was loaded.
    /// </summary>
    /// <remarks>
    /// The definition goes from the slicers part, and the part itself goes once it holds nothing —
    /// leaving an empty <c>&lt;slicers/&gt;</c> behind is a schema violation, not merely untidy.
    /// The extension is dropped once its list is empty, and the extension list once it is. The
    /// anchored frame goes from the sheet's drawing as well, since that is what Excel actually
    /// draws a slicer through. The cache half is unpicked by <see cref="SlicerCacheWriter"/>.
    /// </remarks>
    private static void RemoveDeletedSlicers(
        Worksheet worksheet, XLWorksheetContentManager cm, WorksheetPart worksheetPart, XLSlicers slicers)
    {
        if (slicers.Removed.Count == 0)
            return;

        var touchedParts = new HashSet<SlicersPart>();

        foreach (var removed in slicers.Removed)
        {
            // The frame lives in the drawing rather than in the slicers part, and it names the
            // slicer it draws, so it has to go too — otherwise the sheet still asks Excel to draw
            // something the package no longer defines. Done before the guard below, because a
            // slicer whose part has already gone still has a frame to take out.
            SlicerAnchorXml.Remove(worksheetPart.DrawingsPart, removed);

            if (removed.PartRelId is not { } relId
                || !worksheetPart.Parts.Any(p => p.RelationshipId == relId)
                || worksheetPart.GetPartById(relId) is not SlicersPart part)
            {
                continue;
            }

            var slicer = part.Slicers?
                .Elements<X14.Slicer>()
                .FirstOrDefault(s => s.Name?.Value == removed.Name);

            slicer?.Remove();
            touchedParts.Add(part);
        }

        foreach (var part in touchedParts)
        {
            if (part.Slicers?.Elements<X14.Slicer>().Any() == true)
                continue;

            var relId = worksheetPart.GetIdOfPart(part);
            RemoveSlicerListReference(worksheet, cm, relId);
            worksheetPart.DeletePart(part);
        }
    }

    private static void RemoveSlicerListReference(Worksheet worksheet, XLWorksheetContentManager cm, string relId)
    {
        var extensionList = worksheet.Elements<WorksheetExtensionList>().FirstOrDefault();
        if (extensionList is null)
            return;

        foreach (var extension in extensionList.Elements<WorksheetExtension>().ToList())
        {
            var slicerList = extension.GetFirstChild<X14.SlicerList>();
            if (slicerList is null)
                continue;

            foreach (var reference in slicerList.Elements<X14.SlicerRef>().Where(r => r.Id?.Value == relId).ToList())
                reference.Remove();

            if (!slicerList.Elements<X14.SlicerRef>().Any())
                extension.Remove();
        }

        if (!extensionList.HasChildren)
        {
            worksheet.RemoveChild(extensionList);
            cm.SetElement(XLWorksheetContents.WorksheetExtensionList, null);
        }
    }

    private static WorksheetExtension? FindExtension(Worksheet worksheet, string uri) =>
        worksheet.Elements<WorksheetExtensionList>()
            .FirstOrDefault()?
            .Elements<WorksheetExtension>()
            .FirstOrDefault(e => string.Equals(e.Uri?.Value, uri, System.StringComparison.OrdinalIgnoreCase));

    private static string ExtensionUri(XLSlicerSourceKind kind) =>
        kind == XLSlicerSourceKind.Table ? TableSlicerExtensionUri : PivotSlicerExtensionUri;
}
