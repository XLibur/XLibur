using System;
using System.Collections.Generic;
using System.Text;

namespace XLibur.Excel.IO;

/// <summary>
/// One entry of <c>comments{N}.xml</c> and its matching VML shape.
/// </summary>
/// <param name="Cell">The cell the note is anchored to.</param>
/// <param name="Comment">
/// The note to write. For a threaded cell this is the compatibility fallback rather than anything
/// the user set, since a cell with a thread has no note of its own.
/// </param>
/// <param name="Thread">
/// The comment thread this note is the fallback for, or null for a plain note.
/// </param>
internal readonly record struct CommentWriteEntry(XLCell Cell, XLComment Comment, XLThreadedComment? Thread);

/// <summary>
/// Collects everything that belongs in a sheet's comments and VML parts.
/// </summary>
/// <remarks>
/// Excel pairs every comment thread with a legacy note whose text is "[Threaded comment]"
/// boilerplate followed by the thread's contents, so that a version of Excel too old to understand
/// threads still shows something. XLibur writes the same pairing: the note below is generated from
/// the thread on every save rather than round-tripped, so that it cannot drift out of sync with the
/// thread after an edit.
/// </remarks>
internal static class CommentWriteSource
{
    /// <summary>
    /// The prefix Excel uses to mark a note as a thread's fallback. The rest of the author string is
    /// the thread root's id.
    /// </summary>
    internal const string ThreadAuthorPrefix = "tc=";

    private const string BoilerplateHeader =
        "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, " +
        "any edits to it will get removed if the file is opened in a newer version of Excel. " +
        "Learn more: https://go.microsoft.com/fwlink/?linkid=870924";

    internal static List<CommentWriteEntry> Collect(XLWorksheet worksheet)
    {
        var entries = new List<CommentWriteEntry>();
        foreach (var cell in worksheet.Internals.CellsCollection
            .GetCells(c => c.HasComment || c.HasThreadedComment))
        {
            if (cell.SliceComment is { } note)
                entries.Add(new CommentWriteEntry(cell, note, Thread: null));
            else if (cell.SliceThreadedComment is { } thread)
                entries.Add(new CommentWriteEntry(cell, GetOrCreateFallbackNote(cell, thread), thread));
        }

        return entries;
    }

    /// <summary>
    /// The author string Excel writes for a thread's fallback note, e.g. <c>tc={9E032651-...}</c>.
    /// </summary>
    internal static string ThreadAuthor(XLThreadedComment root) => ThreadAuthorPrefix + FormatId(root.Id);

    /// <summary>
    /// Formats a GUID the way Excel writes ids in threaded comment parts: braced and upper case.
    /// </summary>
    internal static string FormatId(Guid id) => id.ToString("B").ToUpperInvariant();

    /// <summary>
    /// Returns the note that stands in for <paramref name="thread"/> in the compatibility parts,
    /// refreshing its author and text from the thread's current contents.
    /// </summary>
    /// <remarks>
    /// The note object is cached on the thread so that a shape id allocated on one save is reused by
    /// the next, and so that a note loaded from a file keeps the position and size Excel gave it.
    /// </remarks>
    private static XLComment GetOrCreateFallbackNote(XLCell cell, XLThreadedComment thread)
    {
        var note = thread.LegacyNote ??= new XLComment(cell);

        note.Author = ThreadAuthor(thread);
        note.ClearText();
        note.AddText(BuildFallbackText(thread));
        return note;
    }

    private static string BuildFallbackText(XLThreadedComment root)
    {
        var sb = new StringBuilder(BoilerplateHeader);

        sb.Append("\n\nComment:\n    ").Append(root.Text);
        if (root.RepliesInternal is { } replies)
        {
            foreach (var reply in replies)
                sb.Append("\nReply:\n    ").Append(reply.Text);
        }

        return sb.ToString();
    }
}
