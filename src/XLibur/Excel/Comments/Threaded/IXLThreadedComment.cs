using System;
using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// A comment in an Office 365 style comment thread. A cell holds at most one thread, whose root is
/// returned by <see cref="IXLCell.GetThreadedComment"/>; every other comment in the thread is a
/// reply reachable through <see cref="Replies"/>.
/// </summary>
/// <remarks>
/// Threads are flat: Excel supports a root and its direct replies, so <see cref="Parent"/> is the
/// root for every reply and null for the root itself. A cell cannot carry both a threaded comment
/// and a legacy note — see <see cref="IXLCell.CreateThreadedComment"/>.
/// </remarks>
public interface IXLThreadedComment
{
    /// <summary>
    /// The comment's text. Setting it discards any mentions the comment was loaded with, because
    /// their offsets into the text would no longer be meaningful.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// The person who wrote the comment.
    /// </summary>
    IXLPerson Author { get; }

    /// <summary>
    /// When the comment was written, in UTC. Excel stores this without a time zone designator and
    /// interprets it as UTC.
    /// </summary>
    DateTime CreatedUtc { get; }

    /// <summary>
    /// The thread root, or null when this comment <em>is</em> the root.
    /// </summary>
    IXLThreadedComment? Parent { get; }

    /// <summary>
    /// The replies to the thread, oldest first. Always empty for a reply, since threads are flat.
    /// </summary>
    IReadOnlyList<IXLThreadedComment> Replies { get; }

    /// <summary>
    /// The identifier Excel uses to tie a reply to its root and the legacy fallback note to the
    /// thread. Preserved across a load/save round trip.
    /// </summary>
    Guid Id { get; }

    /// <summary>
    /// Whether the thread has been marked resolved. This is a property of the whole thread: reading
    /// it from a reply returns the root's value, and setting it on a reply throws.
    /// </summary>
    /// <exception cref="InvalidOperationException">When set on a reply rather than the thread root.</exception>
    bool Resolved { get; set; }

    /// <summary>
    /// Appends a reply to the thread. Because threads are flat, the reply is added to the thread
    /// root even when this is called on another reply.
    /// </summary>
    /// <param name="author">The reply's author. Must belong to the same workbook.</param>
    /// <param name="text">The reply's text.</param>
    IXLThreadedComment AddReply(IXLPerson author, string text);

    /// <summary>
    /// Removes this comment. Deleting the thread root removes the whole thread, including replies,
    /// from the cell; deleting a reply leaves the rest of the thread intact.
    /// </summary>
    void Delete();
}
