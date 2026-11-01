using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// The registry of cache relationship ids a workbook keeps under one <c>extLst</c> URI — the
/// workbook's half of a slicer cache or a timeline cache. The workbook-side mirror of
/// <see cref="SheetExtensionRefs"/>.
/// </summary>
/// <remarks>
/// A cache part the workbook does not register is an orphan Excel offers to repair. Both callers
/// need the same things: create the extension list and the extension if the workbook has neither,
/// and prune an emptied registry — which is a schema violation rather than merely untidy — along
/// with an emptied extension list. What differs is only the URI, the registry element and how a
/// fresh extension is declared, so those are passed in.
/// </remarks>
internal static class WorkbookExtensionRefs
{
    /// <summary>
    /// The registry under the given extension URI, creating the extension list and the extension
    /// (through <paramref name="createExtension"/>) if either is missing.
    /// </summary>
    /// <returns>
    /// The registry, or null when an existing extension under that URI holds none — a malformed
    /// extension that is left alone rather than repaired.
    /// </returns>
    internal static OpenXmlCompositeElement? EnsureRegistry(
        Workbook workbook,
        string extensionUri,
        Func<WorkbookExtension> createExtension,
        Func<WorkbookExtension, OpenXmlCompositeElement?> registryOf)
    {
        var extensionList = workbook.GetFirstChild<WorkbookExtensionList>();
        if (extensionList is null)
        {
            extensionList = new WorkbookExtensionList();
            workbook.AppendChild(extensionList);
        }

        var extension = FindExtension(extensionList, extensionUri);
        if (extension is null)
        {
            extension = createExtension();
            extensionList.AppendChild(extension);
        }

        return registryOf(extension);
    }

    /// <summary>
    /// Drops every <typeparamref name="TRef"/> the predicate matches from the registry under the
    /// given URI, then prunes the extension and the extension list if either is left empty.
    /// </summary>
    internal static void RemoveRefs<TRef>(
        Workbook? workbook,
        string extensionUri,
        Func<WorkbookExtension, OpenXmlCompositeElement?> registryOf,
        Predicate<TRef> matches)
        where TRef : OpenXmlElement
    {
        var extensionList = workbook?.GetFirstChild<WorkbookExtensionList>();
        var extension = extensionList is null ? null : FindExtension(extensionList, extensionUri);
        var registry = extension is null ? null : registryOf(extension);
        if (registry is null)
            return;

        foreach (var reference in registry.Elements<TRef>().Where(r => matches(r)).ToList())
            reference.Remove();

        // An empty registry is a schema violation rather than merely untidy, so the extension goes
        // once its last reference does.
        if (!registry.Elements<TRef>().Any())
            extension!.Remove();

        if (extensionList is { HasChildren: false })
            extensionList.Remove();
    }

    private static WorkbookExtension? FindExtension(WorkbookExtensionList extensionList, string uri) =>
        extensionList.Elements<WorkbookExtension>()
            .FirstOrDefault(e => string.Equals(e.Uri?.Value, uri, StringComparison.OrdinalIgnoreCase));
}
