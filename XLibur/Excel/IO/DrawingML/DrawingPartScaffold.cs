using System.Linq;
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
    /// The element goes before <c>&lt;tableParts&gt;</c> when there is one, because the schema fixes
    /// the order of a worksheet's children and <c>drawing</c> comes first.
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

        var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
        if (tableParts is not null)
            worksheet.InsertBefore(drawingRef, tableParts);
        else
            worksheet.AppendChild(drawingRef);

        cm.SetElement(XLWorksheetContents.Drawing, worksheet.Elements<Drawing>().First());
    }

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
