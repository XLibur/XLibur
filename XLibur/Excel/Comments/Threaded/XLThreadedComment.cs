using System;
using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLThreadedComment : IXLThreadedComment
{
    private static readonly IReadOnlyList<IXLThreadedComment> NoReplies = Array.Empty<IXLThreadedComment>();

    private List<XLThreadedComment>? _replies;

    private string _text;

    private bool _resolved;

    internal XLThreadedComment(XLCell cell, Guid id, XLPerson author, string text, DateTime createdUtc)
    {
        Cell = cell;
        Id = id;
        AuthorInternal = author;
        _text = text;
        CreatedUtc = createdUtc;
    }

    public string Text
    {
        get => _text;
        set
        {
            ArgumentNullException.ThrowIfNull(value);

            if (string.Equals(_text, value, StringComparison.Ordinal))
                return;

            _text = value;

            // The offsets a mention carries index into the old text, so they cannot survive an edit.
            MentionsXml = null;
        }
    }

    public IXLPerson Author => AuthorInternal;

    public DateTime CreatedUtc { get; internal set; }

    public IXLThreadedComment? Parent => ParentInternal;

    public IReadOnlyList<IXLThreadedComment> Replies => _replies ?? NoReplies;

    public Guid Id { get; }

    public bool Resolved
    {
        get => ParentInternal?._resolved ?? _resolved;
        set
        {
            if (ParentInternal is not null)
            {
                throw new InvalidOperationException(
                    "Resolved is a property of the whole thread and can only be set on the thread root.");
            }

            _resolved = value;
        }
    }

    /// <summary>
    /// The cell the thread belongs to. Replies share the root's cell.
    /// </summary>
    internal XLCell Cell { get; private set; }

    internal XLPerson AuthorInternal { get; private set; }

    internal XLThreadedComment? ParentInternal { get; private set; }

    /// <summary>
    /// The mutable reply list, or null when the thread has no replies. Prefer <see cref="Replies"/>
    /// for reading; this exists so the loader and the writers can work without allocating.
    /// </summary>
    internal List<XLThreadedComment>? RepliesInternal => _replies;

    /// <summary>
    /// The serialized <c>&lt;xltc:mentions&gt;</c> element the comment was loaded with, so that
    /// mentions survive a round trip even though there is no API to inspect or build them. Null for
    /// comments created through the API and for comments whose <see cref="Text"/> has been changed.
    /// </summary>
    internal string? MentionsXml { get; set; }

    /// <summary>
    /// The legacy fallback note Excel pairs with the thread, kept only for its shape geometry so a
    /// round trip preserves the note's position and size. Null for threads created through the API,
    /// in which case the writer emits a default shape. Only ever set on a thread root.
    /// </summary>
    internal XLComment? LegacyNote { get; set; }

    public IXLThreadedComment AddReply(IXLPerson author, string text)
    {
        ArgumentNullException.ThrowIfNull(author);
        ArgumentNullException.ThrowIfNull(text);

        var root = Root;
        var mapped = root.Cell.Worksheet.Workbook.PersonsInternal.Map(author);
        var reply = new XLThreadedComment(root.Cell, Guid.NewGuid(), mapped, text, UtcNowForFile())
        {
            ParentInternal = root
        };

        (root._replies ??= new List<XLThreadedComment>()).Add(reply);
        return reply;
    }

    public void Delete()
    {
        if (ParentInternal is { } parent)
            parent._replies?.Remove(this);
        else
            Cell.SliceThreadedComment = null;
    }

    internal XLThreadedComment Root => ParentInternal ?? this;

    /// <summary>
    /// The current time at the precision a threaded comment part stores, so that the value held in
    /// memory is the one a later load reads back rather than one rounded on the way out.
    /// </summary>
    internal static DateTime UtcNowForFile() => IO.ThreadedCommentPartWriter.TruncateToFilePrecision(DateTime.UtcNow);

    /// <summary>
    /// Appends a reply that already has an identity, used when loading a file or copying a thread.
    /// </summary>
    internal XLThreadedComment AddLoadedReply(Guid id, XLPerson author, string text, DateTime createdUtc)
    {
        var reply = new XLThreadedComment(Cell, id, author, text, createdUtc)
        {
            ParentInternal = this
        };

        (_replies ??= new List<XLThreadedComment>()).Add(reply);
        return reply;
    }

    /// <summary>
    /// Deep copies the thread onto <paramref name="targetCell"/>, mapping every author into the
    /// target workbook's person list. Ids are preserved within a workbook and regenerated when the
    /// copy crosses into another one, because Excel keys the legacy fallback note off the root id
    /// and two threads may not share it.
    /// </summary>
    internal XLThreadedComment CopyTo(XLCell targetCell)
    {
        var targetPersons = targetCell.Worksheet.Workbook.PersonsInternal;

        // A copy always lands on a different cell, so it is a different thread and needs its own id
        // even inside one workbook — the id is what the fallback note's xr:uid points at.
        var copy = new XLThreadedComment(targetCell, Guid.NewGuid(), targetPersons.Map(AuthorInternal), _text,
            CreatedUtc)
        {
            _resolved = _resolved,
            MentionsXml = MentionsXml
        };

        if (LegacyNote is not null)
        {
            copy.LegacyNote = new XLComment(targetCell, LegacyNote, targetCell.Style.Font, LegacyNote.Style);
        }

        if (_replies is not null)
        {
            copy._replies = new List<XLThreadedComment>(_replies.Count);
            foreach (var reply in _replies)
            {
                copy._replies.Add(new XLThreadedComment(targetCell, Guid.NewGuid(),
                    targetPersons.Map(reply.AuthorInternal), reply._text, reply.CreatedUtc)
                {
                    ParentInternal = copy,
                    MentionsXml = reply.MentionsXml
                });
            }
        }

        return copy;
    }

    /// <summary>
    /// Re-points the thread and its replies at another cell after the thread has been moved.
    /// </summary>
    internal void Rehome(XLCell cell)
    {
        Cell = cell;
        if (_replies is null)
            return;

        foreach (var reply in _replies)
            reply.Cell = cell;
    }
}
