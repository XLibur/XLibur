using System;
using System.Collections.Generic;
using System.Diagnostics;
using XLibur.Excel.RichText;

namespace XLibur.Excel;

/// <summary>
/// A class that holds all texts in a workbook. Each text can be either a simple
/// <c>string</c> or a <see cref="XLImmutableRichText"/>.
/// </summary>
internal sealed class SharedStringTable
{
    /// <summary>
    /// Table of <c>Id</c> to text. Some ids are empty (<c>entry.RefCount = 0</c>) and
    /// are tracked in <see cref="_freeIds"/>.
    /// </summary>
    private readonly List<Entry> _table = new();

    /// <summary>
    /// List of indexes in <see cref="_table"/> that are unused.
    /// </summary>
    private readonly List<int> _freeIds = new();

    /// <summary>
    /// text -&gt; id
    /// </summary>
    private readonly Dictionary<Text, int> _reverseDict = new();

    /// <summary>
    /// Set when a cell was loaded from the file's shared string table without its text recording
    /// the index the cell used. See <see cref="MarkFileIndicesIncomplete"/>.
    /// </summary>
    private bool _fileIndicesIncomplete;

    /// <summary>
    /// Number of texts the table holds reference to.
    /// </summary>
    internal int Count => _table.Count - _freeIds.Count;

    /// <summary>
    /// Release all entries so the memory can be reclaimed by GC.
    /// Called from <see cref="XLWorkbook.Dispose()"/>.
    /// </summary>
    internal void Clear()
    {
        _table.Clear();
        _freeIds.Clear();
        _reverseDict.Clear();
    }

    /// <summary>
    /// Pre-size internal data structures to avoid repeated resizing during bulk inserts.
    /// </summary>
    /// <param name="capacity">Expected number of unique strings.</param>
    internal void EnsureCapacity(int capacity)
    {
        _table.EnsureCapacity(capacity);
        _reverseDict.EnsureCapacity(capacity);
    }

    /// <summary>
    /// Release excess capacity from internal collections after bulk loading is complete.
    /// </summary>
    internal void TrimExcess()
    {
        _table.TrimExcess();
        _reverseDict.TrimExcess();
        _freeIds.TrimExcess();
    }

    /// <summary>
    /// Get a string for specified id. Doesn't matter if it is a plain text or a rich text. In both cases, return text.
    /// </summary>
    internal string this[int id]
    {
        get
        {
            var potentialText = _table[id].Text.Value;
            if (potentialText is string text)
                return text;

            if (potentialText is XLImmutableRichText richText)
                return richText.Text;

            throw new ArgumentException($"Id {id} has no text.");
        }
    }

    /// <summary>
    /// The principle is that every entry is a text, but only some are rich text.
    /// This tries to get a rich text if it is one. If it is just plain text, return null.
    /// </summary>
    internal XLImmutableRichText? GetRichText(int id)
    {
        var text = _table[id].Text.Value;
        if (text is null)
            throw new ArgumentException($"Id {id} has no text.");

        return text as XLImmutableRichText;
    }

    /// <summary>
    /// Get id for a text and increase the number of references to the text by one.
    /// </summary>
    /// <returns>ID of a text in the SST.</returns>
    internal int IncreaseRef(string text, bool inline) => IncreaseTextRef(new Text(text, inline));

    /// <inheritdoc cref="IncreaseRef(string, bool)"/>
    internal int IncreaseRef(XLImmutableRichText text, bool inline) => IncreaseTextRef(new Text(text, inline));

    /// <summary>
    /// Decrease the reference count of a text and free if necessary.
    /// </summary>
    internal void DecreaseRef(int id)
    {
        var entry = _table[id];
        if (entry.Text.Value is null)
            throw new InvalidOperationException("Trying to release a text that doesn't have a reference.");

        if (entry.RefCount > 1)
        {
            _table[id] = new Entry(entry.Text, entry.RefCount - 1, entry.FileIndex);
            return;
        }

        // A freed entry forgets its file index with its text: the id may be reused for another text.
        _table[id] = new Entry(Text.Empty, 0);
        _freeIds.Add(id);
        _reverseDict.Remove(entry.Text);
    }

    /// <summary>
    /// Record that the text with <paramref name="id"/> was loaded from item
    /// <paramref name="fileIndex"/> of the file's shared string table, so a save writes it back at
    /// the same place (see <see cref="GetConsecutiveMap"/>). A text the file lists more than once
    /// keeps the lowest index.
    /// </summary>
    internal void RecordFileIndex(int id, int fileIndex)
    {
        var entry = _table[id];

        // Cells used two indices for one text. Whichever the text keeps, the cells that used the
        // other name an index the save does not write the text at.
        if (entry.FileIndex != Entry.NoFileIndex && entry.FileIndex != fileIndex)
            _fileIndicesIncomplete = true;

        if (entry.FileIndex == Entry.NoFileIndex || fileIndex < entry.FileIndex)
            _table[id] = new Entry(entry.Text, entry.RefCount, fileIndex);
    }

    /// <summary>
    /// Records that a cell was loaded from the file's shared string table, but its text does not
    /// record the index the cell used: the index was out of range, or the text was kept inline.
    /// </summary>
    internal void MarkFileIndicesIncomplete() => _fileIndicesIncomplete = true;

    /// <summary>
    /// Does <paramref name="map"/> (from <see cref="GetConsecutiveMap"/>) write every text at the
    /// index each cell loaded it from? Only then can a save keep cells as the file had them.
    /// </summary>
    /// <remarks>
    /// Not when a cell's index was not recorded (<see cref="MarkFileIndicesIncomplete"/>), nor when
    /// any text moved: that happens when the file listed a text no cell uses ahead of others, or a
    /// cell edit since the load left a file text unused.
    /// </remarks>
    internal bool WritesTextsAtFileIndices(int[] map)
    {
        if (_fileIndicesIncomplete)
            return false;

        for (var i = 0; i < _table.Count; i++)
        {
            var entry = _table[i];
            if (IsShared(entry) && entry.FileIndex != Entry.NoFileIndex && map[i] != entry.FileIndex)
                return false;
        }

        return true;
    }

    /// <summary>
    /// Get a map that takes the actual string id and returns a continuous sequence (i.e., no gaps).
    /// If an id is free (no ref count), the id is mapped to -1.
    /// </summary>
    /// <remarks>
    /// Texts loaded from a file come first, in the order the file listed them, and every other text
    /// follows in id order. So a workbook saved without changing its texts writes each at the index
    /// the file had, and its cells keep the indices they were loaded with. Mapping in id order alone
    /// would write the texts in the order the cells first used them, which is not the order of a
    /// table Excel wrote. The indices shift only where the file listed a text that no cell uses, or
    /// listed one text twice, because neither is written again.
    /// </remarks>
    internal int[] GetConsecutiveMap()
    {
        var map = new int[_table.Count];
        Array.Fill(map, -1);

        var mappedStringId = 0;
        foreach (var id in GetIdsByFileIndex())
        {
            if (id >= 0)
                map[id] = mappedStringId++;
        }

        for (var i = 0; i < _table.Count; i++)
        {
            if (map[i] < 0 && IsShared(_table[i]))
                map[i] = mappedStringId++;
        }

        return map;
    }

    /// <summary>
    /// Slot i holds the id of the first shared text loaded from file item i, or -1.
    /// </summary>
    private int[] GetIdsByFileIndex()
    {
        var maxFileIndex = Entry.NoFileIndex;
        for (var i = 0; i < _table.Count; i++)
        {
            var entry = _table[i];
            if (IsShared(entry) && entry.FileIndex > maxFileIndex)
                maxFileIndex = entry.FileIndex;
        }

        var byFileIndex = maxFileIndex >= 0 ? new int[maxFileIndex + 1] : [];
        Array.Fill(byFileIndex, -1);

        for (var i = 0; i < _table.Count; i++)
        {
            var entry = _table[i];
            if (IsShared(entry) && entry.FileIndex != Entry.NoFileIndex && byFileIndex[entry.FileIndex] < 0)
                byFileIndex[entry.FileIndex] = i;
        }

        return byFileIndex;
    }

    private static bool IsShared(Entry entry) =>
        entry.RefCount > 0 && // Only used entry can be written to sst
        !entry.Text.Inline; // Inline texts shouldn't be written to sst

    private int IncreaseTextRef(Text text)
    {
        if (!_reverseDict.TryGetValue(text, out var id))
        {
            id = AddText(text);
            _reverseDict.Add(text, id);
            return id;
        }

        var entry = _table[id];
        _table[id] = new Entry(entry.Text, entry.RefCount + 1, entry.FileIndex);
        return id;
    }

    private int AddText(Text text)
    {
        if (_freeIds.Count > 0)
        {
            // List only changes size, not underlaying array, if last element is removed.
            var lastIndex = _freeIds.Count - 1;
            var id = _freeIds[lastIndex];
            _freeIds.RemoveAt(lastIndex);
            _table[id] = new Entry(text, 1);
            return id;
        }

        var lastTableIndex = _table.Count;
        _table.Add(new Entry(text, 1));
        return lastTableIndex;
    }

    /// <summary>
    /// A struct to hold a text. It also needs a flag for inline/shared, because they have to be different
    /// in the table. If there was no inline/shared flag, there would be no way to easily determine whether
    /// a text should be written to sst or it should be inlined.
    /// </summary>
    [DebuggerDisplay("{Value} (Shared:{!Inline})")]
    private readonly struct Text : IEquatable<Text>
    {
        internal static readonly Text Empty = new(null, false);

        /// <summary>
        /// Either a <c>string</c>, <c>XLImmutableRichText</c> or null if <c><see cref="Entry.RefCount"/> == 0</c>.
        /// </summary>
        internal readonly object? Value;

        /// <summary>
        /// Must be as flag for inline string, so the default value is false => ShareString is true by default
        /// </summary>
        internal readonly bool Inline;

        internal Text(object? value, bool inline)
        {
            Value = value;
            Inline = inline;
        }

        public override bool Equals(object? obj) => obj is Text other && Equals(other);

        public bool Equals(Text other) => Equals(Value, other.Value) && Inline == other.Inline;

        public override int GetHashCode()
        {
            unchecked
            {
                return ((Value is not null ? Value.GetHashCode() : 0) * 397) ^ Inline.GetHashCode();
            }
        }
    }

    [DebuggerDisplay("{Text.Value}:{RefCount} (Shared:{!Text.Inline})")]
    private readonly struct Entry
    {
        internal readonly Text Text;

        /// <summary>
        /// How many objects (cells, pivot cache entries...) reference the text.
        /// </summary>
        internal readonly int RefCount;

        /// <summary>
        /// The index of the text in the shared string table of the file it was loaded from, or
        /// <see cref="NoFileIndex"/> for a text that was not loaded from one.
        /// </summary>
        internal readonly int FileIndex;

        internal const int NoFileIndex = -1;

        internal Entry(Text text, int refCount, int fileIndex = NoFileIndex)
        {
            Text = text;
            RefCount = refCount;
            FileIndex = fileIndex;
        }
    }
}
