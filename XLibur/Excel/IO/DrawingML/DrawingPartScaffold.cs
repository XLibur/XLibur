using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// The part and the sheet reference a drawing needs before anything can be anchored into it.
/// </summary>
/// <remarks>
/// <para>
/// Every kind of drawing needs the same two things and neither has anything to do with what is being
/// drawn: a <c>DrawingsPart</c> for the sheet, and a <c>&lt;drawing r:id&gt;</c> element in the
/// worksheet pointing at it. A part with no reference is an orphan; a reference with no part makes
/// Excel offer to repair the file.
/// </para>
/// <para>
/// Both were private to <see cref="ChartWriter"/> and duplicated again inside
/// <c>PictureWriter</c>. This is the shared copy; the chart writer now calls it rather than keeping
/// a third. Folding the picture writer's own copy in belongs with the rest of spec 16's
/// consolidation.
/// </para>
/// </remarks>
internal static class DrawingPartScaffold
{
    /// <summary>
    /// The sheet's drawing part, created along with an empty <c>xdr:wsDr</c> root if it has none.
    /// </summary>
    /// <remarks>
    /// Materialises the drawing DOM, so a caller with nothing to add should not call it: touching
    /// the part is what makes the SDK write it back on save.
    /// </remarks>
    internal static DrawingsPart EnsureDrawingsPart(WorksheetPart worksheetPart, SaveContext context)
    {
        var drawingsPart = worksheetPart.DrawingsPart
                           ?? worksheetPart.AddNewPart<DrawingsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
        return drawingsPart;
    }

    /// <summary>
    /// Makes the worksheet point at its drawing part, if it does not already.
    /// </summary>
    /// <remarks>
    /// The schema fixes the order of a worksheet's children, so the element goes after the last one
    /// that comes before <c>drawing</c>. A sheet with no tables has no <c>&lt;tableParts&gt;</c> to
    /// put it in front of (#709).
    /// </remarks>
    internal static void EnsureDrawingElement(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        DrawingsPart drawingsPart)
    {
        if (worksheet.OfType<Drawing>().Any())
            return;

        var drawingRef = new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) };
        drawingRef.AddNamespaceDeclaration("r", RelationshipsNs);

        worksheet.InsertAfter(drawingRef, ElementBeforeDrawing(worksheet, cm));
        cm.SetElement(XLWorksheetContents.Drawing, drawingRef);
    }

    /// <summary>The element a new <c>&lt;drawing&gt;</c> goes after, or null to put it first.</summary>
    /// <remarks>
    /// <c>&lt;smartTags&gt;</c> comes right before <c>&lt;drawing&gt;</c> in the schema. The SDK has
    /// no class for it, so it loads as an unknown element the content manager does not track, and it
    /// is looked for here by name.
    /// </remarks>
    internal static OpenXmlElement? ElementBeforeDrawing(Worksheet worksheet, XLWorksheetContentManager cm) =>
        worksheet.ChildElements.LastOrDefault(e => e.LocalName == "smartTags" && e.NamespaceUri == Main2006SsNs)
        ?? cm.GetPreviousElementFor(XLWorksheetContents.Drawing);

    /// <summary>
    /// Declares the two namespaces every anchored drawing uses, when the root does not already.
    /// </summary>
    internal static void EnsureNamespaces(Xdr.WorksheetDrawing worksheetDrawing)
    {
        if (!worksheetDrawing.NamespaceDeclarations.Any(nd => nd.Value.Equals(DrawingMain2006Ns)))
            worksheetDrawing.AddNamespaceDeclaration("a", DrawingMain2006Ns);

        if (!worksheetDrawing.NamespaceDeclarations.Any(nd => nd.Value.Equals(RelationshipsNs)))
            worksheetDrawing.AddNamespaceDeclaration("r", RelationshipsNs);
    }
}
