using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
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
/// Excel uses a different extension URI for each kind of slicer and puts each kind in its own
/// slicers part, so a sheet carrying both a table slicer and a pivot slicer has two of each. That
/// split is reproduced rather than flattened.
/// </para>
/// </remarks>
internal static class SlicerWriter
{
    /// <summary>The worksheet extension holding the list of pivot slicers on the sheet.</summary>
    private const string PivotSlicerExtensionUri = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}";

    /// <summary>The worksheet extension holding the list of table slicers on the sheet.</summary>
    private const string TableSlicerExtensionUri = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}";

    /// <remarks>
    /// Declared here rather than shared with <c>ChartSeriesFormatXml</c>, which holds the same
    /// constant: that file belongs to spec 16's DrawingML extraction and is not ours to touch.
    /// </remarks>
    internal const double EmuPerPoint = 12700;

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
                SlicerPatcher.PatchSlicer(worksheetPart, slicer, EmuPerPoint);
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
        var part = EnsureSlicersPart(worksheet, worksheetPart, xlSlicer.SourceKind, context, out var relId);
        xlSlicer.PartRelId = relId;

        var slicersRoot = part.Slicers ??= NewSlicersRoot();

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
        slicer.RowHeight = (uint)System.Math.Round(rowHeightPt * EmuPerPoint, System.MidpointRounding.AwayFromZero);

        slicersRoot.AppendChild(slicer);

        EnsureSlicerListReference(worksheet, cm, xlSlicer.SourceKind, relId);
    }

    // ── The slicers part ────────────────────────────────────────────────

    /// <summary>
    /// The slicers part for one kind of slicer on this worksheet, creating it if the sheet has none
    /// yet.
    /// </summary>
    /// <remarks>
    /// A part is reused only when the worksheet's <c>extLst</c> already points at it under the URI
    /// for this kind. Sharing one part between a table slicer and a pivot slicer would leave one of
    /// the two reachable only through the wrong extension, which Excel reads as a broken workbook.
    /// </remarks>
    private static SlicersPart EnsureSlicersPart(
        Worksheet worksheet,
        WorksheetPart worksheetPart,
        XLSlicerSourceKind kind,
        SaveContext context,
        out string relId)
    {
        var extension = FindExtension(worksheet, ExtensionUri(kind));
        if (extension is not null)
        {
            foreach (var reference in extension.Descendants<X14.SlicerRef>())
            {
                if (reference.Id?.Value is { } listed
                    && worksheetPart.Parts.Any(p => p.RelationshipId == listed)
                    && worksheetPart.GetPartById(listed) is SlicersPart existing)
                {
                    relId = listed;
                    return existing;
                }
            }
        }

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
    /// cache half is unpicked by <see cref="SlicerCacheWriter"/>.
    /// </remarks>
    private static void RemoveDeletedSlicers(
        Worksheet worksheet, XLWorksheetContentManager cm, WorksheetPart worksheetPart, XLSlicers slicers)
    {
        if (slicers.Removed.Count == 0)
            return;

        var touchedParts = new HashSet<SlicersPart>();

        foreach (var removed in slicers.Removed)
        {
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
