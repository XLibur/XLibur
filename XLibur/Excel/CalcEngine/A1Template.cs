using System;
using System.Buffers;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using XLibur.Parser;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// R1C1 formula text, parsed once, that writes the A1 text of the formula at any cell. The text is
/// the one <see cref="FormulaText.TryConvert"/> gives for that cell, character for character (#542).
/// <see cref="FormulaText.TryCreateA1Template"/> makes it.
/// </summary>
/// <remarks>
/// <para>
/// The conversion to A1 writes each reference again and copies every other part of the text as it
/// is written. So the A1 text of a cell is the parts between the references, which are the same at
/// every cell, and each reference moved to that cell. The template keeps those parts, and one slot
/// for each reference.
/// </para>
/// <para>
/// A file keeps one R1C1 text for all the cells of a shared formula. The loader parses it once, here,
/// and each cell then allocates its A1 text and nothing else.
/// </para>
/// </remarks>
internal sealed class A1Template
{
    /// <summary>The longest A1 text of an area: <c>$XFD$1048576:$XFD$1048576</c>.</summary>
    private const int MaxAreaLength = 25;

    /// <summary>A text up to this length is written on the stack.</summary>
    private const int StackLimit = 256;

    /// <summary>A reference that the move puts off the sheet.</summary>
    private const string RefError = "#REF!";

    /// <summary>A bang reference that the move puts off the sheet.</summary>
    private const string BangRefError = "!#REF!";

    /// <summary>
    /// The text before each slot, and the text after the last one, so one more than
    /// <see cref="_slots"/>.
    /// </summary>
    private readonly string[] _literals;

    private readonly Slot[] _slots;

    /// <summary>The length of the longest text a cell can get.</summary>
    private readonly int _maxLength;

    private A1Template(string[] literals, Slot[] slots)
    {
        Debug.Assert(literals.Length == slots.Length + 1);
        _literals = literals;
        _slots = slots;

        var maxLength = 0;
        foreach (var literal in literals)
            maxLength += literal.Length;

        foreach (var slot in slots)
            maxLength += slot.MaxLength;

        _maxLength = maxLength;
    }

    /// <summary>
    /// The factory that <see cref="FormulaText.TryCreateA1Template"/> hands to the parser. It notes
    /// each part of the text that the conversion to A1 writes again.
    /// </summary>
    internal static IAstFactory<SymbolRange, SymbolRange, Builder> Factory { get; } = new PartsFactory();

    /// <summary>
    /// The A1 text of the formula in the cell at <paramref name="origin"/>.
    /// </summary>
    internal string ToA1(Point origin)
    {
        // A formula with no reference has the same text in every cell.
        if (_slots.Length == 0)
            return _literals[0];

        char[]? rented = null;
        Span<char> buffer = _maxLength <= StackLimit
            ? stackalloc char[StackLimit]
            : rented = ArrayPool<char>.Shared.Rent(_maxLength);

        var literal = _literals[0];
        literal.CopyTo(buffer);
        var written = literal.Length;
        for (var i = 0; i < _slots.Length; ++i)
        {
            written += _slots[i].Write(buffer[written..], origin.Row, origin.Column);
            literal = _literals[i + 1];
            literal.CopyTo(buffer[written..]);
            written += literal.Length;
        }

        var text = new string(buffer[..written]);
        if (rented is not null)
            ArrayPool<char>.Shared.Return(rented);

        return text;
    }

    /// <summary>
    /// One reference of the formula. It writes what the conversion to A1 writes for the reference at a
    /// cell: the prefix and the area moved to the cell, or an error when the move puts an end of the
    /// area off the sheet.
    /// </summary>
    private readonly struct Slot
    {
        private readonly RowCol _first;
        private readonly RowCol _second;

        /// <summary>
        /// A1 writes one end of an area whose two ends are the same cell, and both ends of a row or a
        /// column span, even when they are the same row or column.
        /// </summary>
        private readonly bool _writeBothEnds;

        /// <summary>The sheet prefix as the conversion writes it, for example <c>'My Sheet'!</c>.</summary>
        private readonly string _prefix;

        /// <summary>What the conversion writes for the whole reference when an end is off the sheet.</summary>
        private readonly string _error;

        internal Slot(ReferenceArea area, string prefix, string error)
        {
            _first = area.First;
            _second = area.Second;

            // Moving both ends by the same cell keeps them equal, or different, so the test can be made
            // on the R1C1 ends once.
            _writeBothEnds = area.First != area.Second || area.First.IsRow || area.First.IsColumn;
            _prefix = prefix;
            _error = error;
        }

        internal int MaxLength => _prefix.Length + MaxAreaLength;

        internal int Write(Span<char> destination, int row, int column)
        {
            if (!TryMove(_first, row, column, out var firstRow, out var firstColumn) ||
                !TryMove(_second, row, column, out var secondRow, out var secondColumn))
            {
                _error.CopyTo(destination);
                return _error.Length;
            }

            _prefix.CopyTo(destination);
            var written = _prefix.Length;
            written += WriteEnd(destination[written..], _first, firstRow, firstColumn);
            if (_writeBothEnds)
            {
                destination[written++] = ':';
                written += WriteEnd(destination[written..], _second, secondRow, secondColumn);
            }

            return written;
        }

        /// <summary>
        /// Move an R1C1 end of an area to the cell at <paramref name="row"/> and
        /// <paramref name="column"/>, as the conversion to A1 does.
        /// </summary>
        /// <returns><c>false</c> when a relative axis is moved off the sheet.</returns>
        private static bool TryMove(RowCol end, int row, int column, out int movedRow, out int movedColumn)
        {
            movedRow = Move(end.RowType, end.RowValue, row);
            movedColumn = Move(end.ColumnType, end.ColumnValue, column);
            return (end.RowType != ReferenceAxisType.Relative || movedRow is >= RowCol.MinRow and <= RowCol.MaxRow) &&
                   (end.ColumnType != ReferenceAxisType.Relative || movedColumn is >= RowCol.MinCol and <= RowCol.MaxCol);

            static int Move(ReferenceAxisType type, int value, int anchor) => type switch
            {
                ReferenceAxisType.Relative => value + anchor,
                ReferenceAxisType.Absolute => value,
                ReferenceAxisType.None => 0,
                _ => throw new NotSupportedException()
            };
        }

        /// <summary>Write one end of an area in A1 notation, for example <c>$B3</c>.</summary>
        private static int WriteEnd(Span<char> destination, RowCol end, int row, int column)
        {
            var written = 0;
            if (end.ColumnType != ReferenceAxisType.None)
            {
                if (end.ColumnType == ReferenceAxisType.Absolute)
                    destination[written++] = '$';

                written += WriteColumn(destination[written..], column);
            }

            if (end.RowType != ReferenceAxisType.None)
            {
                if (end.RowType == ReferenceAxisType.Absolute)
                    destination[written++] = '$';

                var formatted = row.TryFormat(destination[written..], out var digits, default, CultureInfo.InvariantCulture);
                Debug.Assert(formatted);
                written += digits;
            }

            return written;
        }

        /// <summary>Write the letters of a column, for example <c>XFD</c> for column 16384.</summary>
        private static int WriteColumn(Span<char> destination, int column)
        {
            var length = column switch
            {
                <= 26 => 1,
                <= 702 => 2,
                _ => 3,
            };

            for (var i = length - 1; i >= 0; --i)
            {
                column--;
                destination[i] = (char)('A' + column % 26);
                column /= 26;
            }

            return length;
        }
    }

    /// <summary>
    /// Collects, while the parser reads the R1C1 text, each part of it that the conversion to A1
    /// writes again, and then makes the template.
    /// </summary>
    /// <param name="text">The text the parser reads. Each range the parser reports indexes it.</param>
    internal sealed class Builder(string text)
    {
        private readonly List<Rewrite> _rewrites = [];

        /// <summary>Sheet prefixes written so far, so a sheet the formula names often is written once.</summary>
        private List<(string Text, string Written)>? _prefixes;

        /// <summary>
        /// The formula calls a cell as a function, as in <c>R1C1(5)</c>. The conversion writes such a
        /// call as <c>#REF!</c>, arguments and all, when the cell is off the sheet, so the call is no
        /// slot of its own. The template does not take it.
        /// </summary>
        private bool _hasCellFunction;

        /// <summary>
        /// Make the template.
        /// </summary>
        /// <param name="original">The R1C1 text as the caller passed it.</param>
        /// <param name="root">The range of the whole formula in the text.</param>
        /// <param name="placeholder">
        /// The character the parser read in place of a colon inside a single-bracket column name, or
        /// <see cref="FormulaText.NoPlaceholder"/> when no colon was hidden. Only a sheet prefix can
        /// reach the template still holding it, because a prefix is taken from the text the parser
        /// read. Every literal is taken from <paramref name="original"/>, the caller's own text, so a
        /// literal holds the real colon already, and a character the formula really had — a fullwidth
        /// colon in a string, say — must be left alone (#557).
        /// </param>
        /// <returns><c>null</c> when the text has a part the template does not take.</returns>
        internal A1Template? Build(string original, SymbolRange root, char placeholder)
        {
            if (_hasCellFunction)
                return null;

            _rewrites.Sort(static (x, y) => x.Range.Start.CompareTo(y.Range.Start));

            var literals = new List<string>(_rewrites.Count + 1);
            var slots = new List<Slot>(_rewrites.Count);
            var literal = new StringBuilder(original.Length);

            // What the conversion keeps of the text around the formula: the whitespace before it, and the
            // whitespace after it, which the parser trims.
            literal.Append(original, 0, root.Start);
            var position = root.Start;
            foreach (var rewrite in _rewrites)
            {
                // The parts are leaves of the formula, so none holds another.
                if (rewrite.Range.Start < position)
                    return null;

                literal.Append(original, position, rewrite.Range.Start - position);
                position = rewrite.Range.End;
                if (rewrite.Replacement is not null)
                {
                    literal.Append(rewrite.Replacement);
                    continue;
                }

                if (!TryGetPrefix(in rewrite, out var prefix))
                    return null;

                literals.Add(literal.ToString());
                literal.Clear();
                slots.Add(MakeSlot(in rewrite, prefix, placeholder));
            }

            literal.Append(original, position, root.End - position);
            var trimmedLength = original.AsSpan().TrimEnd().Length;
            literal.Append(original, trimmedLength, original.Length - trimmedLength);
            literals.Add(literal.ToString());
            return new A1Template(literals.ToArray(), slots.ToArray());
        }

        /// <summary>The text a slot writes before its reference, or <c>false</c> when the template does not take it.</summary>
        private bool TryGetPrefix(in Rewrite rewrite, out string prefix)
        {
            switch (rewrite.Prefix)
            {
                case SlotPrefix.None:
                    prefix = string.Empty;
                    return true;
                case SlotPrefix.Bang:
                    prefix = "!";
                    return true;
                default:
                    return TryWritePrefix(rewrite.Range, out prefix);
            }
        }

        private static Slot MakeSlot(in Rewrite rewrite, string prefix, char placeholder)
        {
            var error = rewrite.Prefix == SlotPrefix.Bang ? BangRefError : RefError;
            return new Slot(rewrite.Area, RestoreColons(prefix, placeholder), error);
        }

        private static string RestoreColons(string value, char placeholder)
            => placeholder == FormulaText.NoPlaceholder ? value : value.Replace(placeholder, ':');

        internal void AddReference(SymbolRange range, ReferenceArea reference, SlotPrefix prefix)
        {
            // Only an R1C1 reference moves. The conversion keeps an A1 one as it is written.
            if (reference.Style == ReferenceStyle.R1C1)
                _rewrites.Add(new Rewrite(range, null, reference, prefix));
        }

        internal void AddError(SymbolRange range, ReadOnlySpan<char> error)
        {
            // The conversion writes a #REF! with more after it, such as #REF!A1, as a plain #REF!, as Excel
            // saves it. It is the same text in every cell.
            if (range.Length != error.Length &&
                text.AsSpan(range.Start, range.Length).StartsWith(RefError, StringComparison.OrdinalIgnoreCase))
            {
                _rewrites.Add(new Rewrite(range, RefError, default, SlotPrefix.None));
            }
        }

        internal void AddCellFunction() => _hasCellFunction = true;

        /// <summary>
        /// Write the sheet prefix of the reference at <paramref name="range"/> as the conversion writes
        /// it. The prefix is the text of the reference up to its last <c>!</c>: the R1C1 area after it
        /// holds none.
        /// </summary>
        private bool TryWritePrefix(SymbolRange range, out string written)
        {
            var end = text.LastIndexOf('!', range.End - 1, range.Length);
            var prefixText = text.Substring(range.Start, end + 1 - range.Start);
            if (_prefixes is not null)
            {
                foreach (var (knownText, knownWritten) in _prefixes)
                {
                    if (knownText == prefixText)
                    {
                        written = knownWritten;
                        return true;
                    }
                }
            }

            if (!FormulaText.TryWriteSheetPrefix(prefixText, out var prefix))
            {
                written = string.Empty;
                return false;
            }

            (_prefixes ??= []).Add((prefixText, prefix));
            written = prefix;
            return true;
        }
    }

    /// <summary>The prefix a reference is written with.</summary>
    internal enum SlotPrefix
    {
        /// <summary>No prefix, as in <c>R1C1</c>.</summary>
        None,

        /// <summary>The <c>!</c> of a bang reference, as in <c>!R1C1</c>.</summary>
        Bang,

        /// <summary>A sheet prefix, with or without a book, as in <c>Sheet1!R1C1</c> or <c>[1]Sheet1:Sheet3!R1C1</c>.</summary>
        Sheet,
    }

    /// <summary>
    /// A part of the text that the conversion to A1 writes again: a reference, or a fixed text in its
    /// place.
    /// </summary>
    private readonly record struct Rewrite(SymbolRange Range, string? Replacement, ReferenceArea Area, SlotPrefix Prefix);

    /// <summary>
    /// Notes each reference, and each other part the conversion to A1 writes again. Every node is its
    /// range in the text: the template needs no tree, only where the parts are.
    /// </summary>
    private sealed class PartsFactory : IAstFactory<SymbolRange, SymbolRange, Builder>
    {
        public SymbolRange LogicalValue(Builder context, SymbolRange range, bool value) => range;

        public SymbolRange NumberValue(Builder context, SymbolRange range, double value) => range;

        public SymbolRange TextValue(Builder context, SymbolRange range, string text) => range;

        public SymbolRange ErrorValue(Builder context, SymbolRange range, ReadOnlySpan<char> error) => range;

        public SymbolRange ArrayNode(Builder context, SymbolRange range, int rows, int columns,
            IReadOnlyList<SymbolRange> elements) => range;

        public SymbolRange BlankNode(Builder context, SymbolRange range) => range;

        public SymbolRange LogicalNode(Builder context, SymbolRange range, bool value) => range;

        public SymbolRange ErrorNode(Builder context, SymbolRange range, ReadOnlySpan<char> error)
        {
            context.AddError(range, error);
            return range;
        }

        public SymbolRange SheetErrorNode(Builder context, SymbolRange range, int? workbookIndex, string sheet,
            ReadOnlySpan<char> error) => range;

        public SymbolRange NumberNode(Builder context, SymbolRange range, double value) => range;

        public SymbolRange TextNode(Builder context, SymbolRange range, string text) => range;

        public SymbolRange Reference(Builder context, SymbolRange range, ReferenceArea reference)
        {
            context.AddReference(range, reference, SlotPrefix.None);
            return range;
        }

        public SymbolRange SheetReference(Builder context, SymbolRange range, string sheet, ReferenceArea reference)
        {
            context.AddReference(range, reference, SlotPrefix.Sheet);
            return range;
        }

        public SymbolRange BangReference(Builder context, SymbolRange range, ReferenceArea reference)
        {
            context.AddReference(range, reference, SlotPrefix.Bang);
            return range;
        }

        public SymbolRange Reference3D(Builder context, SymbolRange range, string firstSheet, string lastSheet,
            ReferenceArea reference)
        {
            context.AddReference(range, reference, SlotPrefix.Sheet);
            return range;
        }

        public SymbolRange ExternalSheetReference(Builder context, SymbolRange range, int workbookIndex, string sheet,
            ReferenceArea reference)
        {
            context.AddReference(range, reference, SlotPrefix.Sheet);
            return range;
        }

        public SymbolRange ExternalReference3D(Builder context, SymbolRange range, int workbookIndex,
            string firstSheet, string lastSheet, ReferenceArea reference)
        {
            context.AddReference(range, reference, SlotPrefix.Sheet);
            return range;
        }

        public SymbolRange Function(Builder context, SymbolRange range, ReadOnlySpan<char> functionName,
            IReadOnlyList<SymbolRange> arguments) => range;

        public SymbolRange Function(Builder context, SymbolRange range, string sheetName,
            ReadOnlySpan<char> functionName, IReadOnlyList<SymbolRange> args) => range;

        public SymbolRange ExternalFunction(Builder context, SymbolRange range, int workbookIndex, string sheetName,
            ReadOnlySpan<char> functionName, IReadOnlyList<SymbolRange> arguments) => range;

        public SymbolRange ExternalFunction(Builder context, SymbolRange range, int workbookIndex,
            ReadOnlySpan<char> functionName, IReadOnlyList<SymbolRange> arguments) => range;

        public SymbolRange CellFunction(Builder context, SymbolRange range, RowCol cell,
            IReadOnlyList<SymbolRange> arguments)
        {
            context.AddCellFunction();
            return range;
        }

        public SymbolRange StructureReference(Builder context, SymbolRange range, StructuredReferenceArea area,
            string? firstColumn, string? lastColumn) => range;

        public SymbolRange StructureReference(Builder context, SymbolRange range, string table,
            StructuredReferenceArea area, string? firstColumn, string? lastColumn) => range;

        public SymbolRange ExternalStructureReference(Builder context, SymbolRange range, int workbookIndex,
            string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn) => range;

        public SymbolRange Name(Builder context, SymbolRange range, string name) => range;

        public SymbolRange SheetName(Builder context, SymbolRange range, string sheet, string name) => range;

        public SymbolRange BangName(Builder context, SymbolRange range, string name) => range;

        public SymbolRange ExternalName(Builder context, SymbolRange range, int workbookIndex, string name) => range;

        public SymbolRange ExternalSheetName(Builder context, SymbolRange range, int workbookIndex, string sheet,
            string name) => range;

        public SymbolRange ExternalDynamicDataExchange(Builder context, SymbolRange range, int workbookIndex,
            string item) => range;

        public SymbolRange DynamicDataExchange(Builder context, SymbolRange range, string application, string topic,
            string item) => range;

        public SymbolRange BinaryNode(Builder context, SymbolRange range, BinaryOperation operation,
            SymbolRange leftNode, SymbolRange rightNode) => range;

        public SymbolRange Unary(Builder context, SymbolRange range, UnaryOperation operation, SymbolRange node) => range;

        public SymbolRange Nested(Builder context, SymbolRange range, SymbolRange node) => range;
    }
}
