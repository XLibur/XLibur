using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// The list of relationship ids a worksheet keeps under one <c>extLst</c> URI — the sheet's half of
/// a slicer or a timeline.
/// </summary>
/// <remarks>
/// A surviving control part is worthless if the sheet stops referencing it, and the worksheet part
/// is rebuilt from the model on every save while the control's part is not. Both callers need the
/// same three things: create the extension list if the sheet has none, register it with the content
/// manager so it lands in schema order, and prune an emptied registry — which is a schema violation
/// rather than merely untidy.
/// </remarks>
internal static class SheetExtensionRefs
{
    /// <summary>
    /// The reference list under the given extension URI, creating the extension list, the extension
    /// and the list itself if any is missing.
    /// </summary>
    internal static TList EnsureList<TList>(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        string extensionUri,
        string namespacePrefix,
        string namespaceUri)
        where TList : OpenXmlCompositeElement, new()
    {
        var extension = FindExtension(worksheet, extensionUri);

        if (extension is null)
        {
            if (!worksheet.Elements<WorksheetExtensionList>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.WorksheetExtensionList);
                worksheet.InsertAfter(new WorksheetExtensionList(), previousElement);
            }

            var extensionList = worksheet.Elements<WorksheetExtensionList>().First();
            cm.SetElement(XLWorksheetContents.WorksheetExtensionList, extensionList);

            extension = new WorksheetExtension { Uri = extensionUri };
            extension.AddNamespaceDeclaration(namespacePrefix, namespaceUri);
            extension.AppendChild(new TList());
            extensionList.AppendChild(extension);
        }

        var list = extension.GetFirstChild<TList>();
        if (list is null)
        {
            list = new TList();
            extension.AppendChild(list);
        }

        return list;
    }

    /// <summary>
    /// Drops every reference the predicate matches from every list of that type on the sheet, then
    /// prunes the extension and the extension list if either is left empty.
    /// </summary>
    /// <remarks>
    /// Every extension is scanned rather than one named URI, because a sheet may carry more than one
    /// list of the same type under different URIs — a sheet with both a pivot slicer and a table
    /// slicer has exactly that — and the caller knows only the relationship id.
    /// </remarks>
    internal static void RemoveRefs<TList>(
        Worksheet worksheet, XLWorksheetContentManager cm, Predicate<OpenXmlElement> matches)
        where TList : OpenXmlCompositeElement
    {
        var extensionList = worksheet.Elements<WorksheetExtensionList>().FirstOrDefault();
        if (extensionList is null)
            return;

        foreach (var extension in extensionList.Elements<WorksheetExtension>().ToList())
        {
            var list = extension.GetFirstChild<TList>();
            if (list is null)
                continue;

            foreach (var reference in list.ChildElements.ToList())
            {
                if (matches(reference))
                    reference.Remove();
            }

            if (!list.HasChildren)
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
            .FirstOrDefault(e => string.Equals(e.Uri?.Value, uri, StringComparison.OrdinalIgnoreCase));
}
