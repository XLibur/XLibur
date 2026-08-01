using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.Exceptions;
using XLibur.Report.Tags;

namespace XLibur.Report.Rewriting;

/// <summary>
/// One column a <c>&lt;&lt;Pivot&gt;&gt;</c> was told to use, and the heading cell that says which
/// column it is and what the field is called.
/// </summary>
/// <remarks>
/// The heading is a live one-cell range rather than a column number and a string, so that a column
/// deleted or moved between the tag being read and the pivot being built takes its field with it
/// instead of leaving the pivot reading whatever moved into its place.
/// </remarks>
internal sealed record PivotField(PivotFieldTag Tag, IXLRange Heading);

/// <summary>
/// A pivot table a <c>&lt;&lt;Pivot&gt;&gt;</c> asked for, held until every range has been generated.
/// </summary>
internal sealed record PivotRequest(
    IXLRange Source,
    IXLRange Destination,
    string RequestedName,
    IReadOnlyList<PivotField> Fields,
    string SheetName,
    string? CellAddress);

/// <summary>
/// Builds the pivot tables the <c>&lt;&lt;Pivot&gt;&gt;</c> tags asked for.
/// </summary>
/// <remarks>
/// Deferred to the end of a run rather than done by the tag, because a pivot table's position and its
/// cache's source are plain rectangles: created while ranges are still being generated, a pivot would
/// stay behind as the sheet moved under it. Waiting means both are read from live ranges that have
/// finished moving — which is also why <see cref="PivotRewriter"/> leaves these pivots alone.
/// </remarks>
internal static class PivotBuilder
{
    /// <summary>Builds each request, reporting the ones that cannot be built.</summary>
    public static void Build(IReadOnlyList<PivotRequest> requests, TemplateErrors errors)
    {
        foreach (var request in requests)
        {
            Build(request, errors);
        }
    }

    private static void Build(PivotRequest request, TemplateErrors errors)
    {
        if (!request.Source.RangeAddress.IsValid || !request.Destination.RangeAddress.IsValid)
        {
            errors.Add(Error(
                request,
                "<<Pivot>> could not be built: the rows it was to read, or the cell it was to be put "
                + "in, were removed while the report was generated."));
            return;
        }

        var destination = request.Destination.FirstCell();

        IXLPivotTable pivot;
        try
        {
            pivot = destination.Worksheet.PivotTables.Add(
                Name(destination.Worksheet.Workbook, request.RequestedName),
                destination,
                request.Source);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidReferenceException)
        {
            // A name the sheet already holds — two <<Pivot name="…">> asking for the same one, or a
            // template pivot that has it — or a source the cache cannot read. Both are template
            // problems, and the rest of the report should survive either.
            errors.Add(Error(request, $"<<Pivot>> could not be built: {ex.Message}"));
            return;
        }

        foreach (var field in request.Fields)
        {
            AddField(pivot, request, field, errors);
        }

        // What the rewriter guarantees for a pivot the template drew, guaranteed here for one the
        // report built: Excel is the authority on its own layout, and re-reading the source on open is
        // what makes the pivot agree with the data beneath it.
        pivot.PivotCache.SetRefreshDataOnOpen();
    }

    private static void AddField(
        IXLPivotTable pivot,
        PivotRequest request,
        PivotField field,
        TemplateErrors errors)
    {
        var tagName = field.Tag.Token.Name;

        if (!field.Heading.RangeAddress.IsValid)
        {
            errors.Add(Error(
                request,
                $"<<{tagName}>> named a column that the report removed, so the pivot table is without "
                + "that field."));
            return;
        }

        var heading = field.Heading.FirstCell().GetFormattedString();

        if (string.IsNullOrWhiteSpace(heading))
        {
            errors.Add(Error(
                request,
                $"<<{tagName}>> is under a column with no heading, so there is no field name for the "
                + "pivot to use. Put the column's heading in the row above the range."));
            return;
        }

        try
        {
            field.Tag.AddTo(pivot, heading);
        }
        catch (Exception ex) when (ex is ArgumentException or KeyNotFoundException)
        {
            // The same field twice, or one the cache does not hold: a template problem, and one the
            // rest of the report should survive.
            errors.Add(new TemplateError(
                $"<<{tagName}>> could not add '{heading}' to the pivot table: {ex.Message}",
                request.SheetName,
                request.CellAddress,
                ex));
        }
    }

    /// <summary>
    /// The name to give the pivot: the one the tag asked for, or <c>Pivot</c> and the first number not
    /// already taken. Names have to be unique across the workbook, and a template producing several
    /// pivots cannot be made to name them all.
    /// </summary>
    private static string Name(XLWorkbook workbook, string requested)
    {
        if (requested.Length > 0)
        {
            return requested;
        }

        var taken = workbook.Worksheets
            .SelectMany(sheet => sheet.PivotTables)
            .Select(pivot => pivot.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var number = 1;
        while (taken.Contains("Pivot" + number))
        {
            number++;
        }

        return "Pivot" + number;
    }

    private static TemplateError Error(PivotRequest request, string message) =>
        new(message, request.SheetName, request.CellAddress);
}
