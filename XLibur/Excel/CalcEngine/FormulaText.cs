using System;
using System.Buffers;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Parser;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The notation formula text is written in.
/// </summary>
internal enum FormulaNotation
{
    A1,
    R1C1,
}

/// <summary>
/// A refusal of formula text, returned as a value instead of thrown. The references in a refused
/// formula are unknown, so each caller of <see cref="FormulaText"/> decides what a refusal means to it.
/// </summary>
/// <remarks>
/// Two things refuse text. The parser refuses text it cannot read, with <see cref="ParsingException"/>.
/// A factory that XLibur gives the parser refuses a part of the text that the parser reads and the calc
/// engine cannot take, with <see cref="ExpressionParseException"/>: a function with too many or too few
/// arguments, or an error value that <see cref="XLError"/> has no member for (#543).
/// </remarks>
internal readonly record struct FormulaRefusal
{
    /// <summary>The parser refused the text.</summary>
    internal FormulaRefusal(string text, ParsingException cause)
    {
        Text = text;
        Cause = cause;
    }

    /// <summary>A factory refused a part of the text while the parser read it.</summary>
    internal FormulaRefusal(string text, ExpressionParseException cause)
    {
        Text = text;
        Cause = cause;
    }

    /// <summary>The formula text that was refused, as the caller passed it.</summary>
    internal string Text { get; }

    /// <summary>
    /// What refused the text: the parser's own <see cref="ParsingException"/>, which says where in the
    /// text the parser stopped, or the <see cref="ExpressionParseException"/> a factory threw.
    /// </summary>
    internal Exception Cause { get; }

    /// <summary>The message of <see cref="Cause"/>. The parser's message has the position it stopped at.</summary>
    internal string Message => Cause.Message;

    /// <summary>
    /// The refusal as a public caller sees it. Only the public edges throw it: evaluation,
    /// <see cref="IXLCell.FormulaR1C1"/>, and copying a formula. The parser's exception stays inside
    /// it, so the position the parser reported is still there. A factory's exception is the public type
    /// already, so it is given as it is, and evaluation throws what it threw before #543.
    /// </summary>
    internal ExpressionParseException ToException()
        => Cause as ExpressionParseException ?? new ExpressionParseException(Cause.Message, Cause);
}

/// <summary>
/// The only code in XLibur that hands formula text to <c>ClosedXML.Parser</c>.
/// </summary>
/// <remarks>
/// Everything that must happen to formula text before the parser sees it, and after the parser is
/// done, happens here:
/// <list type="bullet">
/// <item>A colon inside a table column name, as in <c>Table1[Start: Date]</c>, is hidden from the
/// parser, which would read it as a range operator. It is put back in every column name the parser
/// reports and in every text it produces.</item>
/// <item>A future function gets its prefix on the way in (<see cref="AddFuturePrefixes"/>). On the way
/// to evaluation the prefix comes off, in any case (<see cref="TryStripFuturePrefix"/>).</item>
/// <item>Formula text has no leading <c>=</c>. Where a caller's text can have one, it comes off here
/// (<see cref="WithoutLeadingEquals"/>).</item>
/// <item>The parser refuses text by throwing <see cref="ParsingException"/>. A factory refuses a part
/// of the text that the calc engine cannot take, such as a function with the wrong number of
/// arguments, by throwing <see cref="ExpressionParseException"/> (#543). Both are caught in one place,
/// <see cref="TryParse{TState,TResult}"/>, and returned as a <see cref="FormulaRefusal"/>. A caller
/// makes one decision only: what a refused formula means to it. No other file in XLibur catches
/// <see cref="ParsingException"/>.</item>
/// </list>
/// </remarks>
internal static class FormulaText
{
    /// <summary>
    /// A placeholder character (fullwidth colon U+FF1A) that stands in for a colon inside a
    /// single-bracket column name while the parser reads the text. The parser accepts it in a column
    /// name and does not read it as a range operator. It is one UTF-16 character, the same width as
    /// the colon, so every <see cref="SymbolRange"/> the parser reports indexes the original text.
    /// </summary>
    private const char ColonPlaceholder = '：';

    /// <summary>The prefix a file puts on a future function, for example <c>_xlfn.CONCAT</c>.</summary>
    private const string FuturePrefix = "_xlfn.";

    /// <summary>
    /// The further prefix a worksheet-only future function has after <see cref="FuturePrefix"/>, for
    /// example <c>_xlfn._xlws.FILTER</c>.
    /// </summary>
    private const string WorksheetFuturePrefix = "_xlws.";

    private static readonly Lazy<PrefixTree> FutureFunctionSet =
        new(() => PrefixTree.Build(XLConstants.FutureFunctionMap.Value.Keys));

    private static readonly RenameFunctionsVisitor RemapFutureFunctions = new(XLConstants.FutureFunctionMap);

    /// <summary>
    /// Parses <paramref name="text"/> and hands each part of it to <paramref name="factory"/>.
    /// </summary>
    /// <param name="text">The formula text, without a leading <c>=</c>.</param>
    /// <param name="context">Passed to every method of <paramref name="factory"/>.</param>
    /// <param name="factory">Builds a node for each part of the formula.</param>
    /// <param name="notation">The notation <paramref name="text"/> is written in.</param>
    /// <param name="root">The node <paramref name="factory"/> built for the whole formula.</param>
    /// <param name="refusal">Why the parser refused the text, when it did.</param>
    /// <returns><c>false</c> when the parser refused the text.</returns>
    /// <remarks>
    /// A <see cref="SymbolRange"/> the factory receives indexes <paramref name="text"/> itself, and a
    /// column name reaches the factory as it is written, colon included.
    /// </remarks>
    internal static bool TryWalk<TScalar, TNode, TContext>(
        string text,
        TContext context,
        IAstFactory<TScalar, TNode, TContext> factory,
        FormulaNotation notation,
        [MaybeNullWhen(false)] out TNode root,
        out FormulaRefusal refusal)
    {
        var parseable = ProtectStructuredRefColons(text, out var wasProtected);

        // Only text that had a colon hidden needs the column names restored. Everything else goes to
        // the caller's factory directly, so it costs no allocation.
        if (wasProtected)
            factory = new ColonRestoringFactory<TScalar, TNode, TContext>(factory);

        return TryParse(
            text,
            (Text: parseable, Context: context, Factory: factory, Notation: notation),
            static s => s.Notation == FormulaNotation.A1
                ? FormulaParser<TScalar, TNode, TContext>.CellFormulaA1(s.Text, s.Context, s.Factory)
                : FormulaParser<TScalar, TNode, TContext>.CellFormulaR1C1(s.Text, s.Context, s.Factory),
            out root,
            out refusal);
    }

    /// <summary>
    /// Rewrites the parts of A1 <paramref name="text"/> that <paramref name="modifier"/> changes, and
    /// keeps the rest of the text as it is written.
    /// </summary>
    /// <param name="text">The formula text, in A1 notation.</param>
    /// <param name="sheetName">The sheet the formula is on.</param>
    /// <param name="origin">The cell the formula is in.</param>
    /// <param name="modifier">Says what changes.</param>
    /// <param name="rewritten">The rewritten text, or <paramref name="text"/> unchanged when refused.</param>
    /// <param name="refusal">Why the parser refused the text, when it did.</param>
    /// <returns><c>false</c> when the parser refused the text.</returns>
    internal static bool TryRewrite(string text, string sheetName, Point origin, FormulaModifier modifier,
        out string rewritten, out FormulaRefusal refusal)
    {
        var parseable = ProtectStructuredRefColons(text, out var wasProtected);
        if (!TryParse(
                text,
                (Text: parseable, Sheet: sheetName, Origin: origin, Modifier: modifier),
                static s => FormulaConverter.ModifyA1(s.Text, s.Sheet, s.Origin.Row, s.Origin.Column, s.Modifier),
                out var result,
                out refusal))
        {
            rewritten = text;
            return false;
        }

        rewritten = wasProtected ? result.Replace(ColonPlaceholder, ':') : result;
        return true;
    }

    /// <summary>
    /// Converts <paramref name="text"/> to the notation <paramref name="to"/>, relative to
    /// <paramref name="origin"/>.
    /// </summary>
    /// <param name="text">The formula text, in the other notation.</param>
    /// <param name="origin">The cell that relative references are relative to.</param>
    /// <param name="to">The notation to convert to.</param>
    /// <param name="converted">The converted text, or <paramref name="text"/> unchanged when refused.</param>
    /// <param name="refusal">Why the parser refused the text, when it did.</param>
    /// <returns><c>false</c> when the parser refused the text.</returns>
    internal static bool TryConvert(string text, Point origin, FormulaNotation to,
        out string converted, out FormulaRefusal refusal)
    {
        var parseable = ProtectStructuredRefColons(text, out var wasProtected);
        if (!TryParse(
                text,
                (Text: parseable, Origin: origin, To: to),
                static s => s.To == FormulaNotation.R1C1
                    ? FormulaConverter.ToR1C1(s.Text, s.Origin.Row, s.Origin.Column)
                    : FormulaConverter.ToA1(s.Text, s.Origin.Row, s.Origin.Column),
                out var result,
                out refusal))
        {
            converted = text;
            return false;
        }

        converted = wasProtected ? result.Replace(ColonPlaceholder, ':') : result;
        return true;
    }

    /// <summary>
    /// Adds the prefix a file needs to each future function that <paramref name="typed"/> calls
    /// without one, for example <c>acot(A5)/2</c> to <c>_xlfn.ACOT(A5)/2</c>.
    /// </summary>
    /// <remarks>
    /// A refused formula comes back unchanged: text the parser cannot read has no future function it
    /// can find. A caller can set such text, for example an external reference in the form the formula
    /// bar shows, <c>'[file.xlsx]Sheet'!A1</c>, where a file stores <c>[1]Sheet!A1</c>. Only a refusal
    /// is answered with the unchanged text. A defect in the remap reaches the caller.
    /// </remarks>
    internal static string AddFuturePrefixes(string typed, string sheetName, Point origin)
    {
        // A string check first. It is far cheaper than a parse, and most formulas have no future
        // function.
        if (!MightContainFutureFunction(typed.AsSpan()))
            return typed;

        return TryRewrite(typed, sheetName, origin, RemapFutureFunctions, out var prefixed, out _)
            ? prefixed
            : typed;
    }

    /// <summary>
    /// Takes the future-function prefix off a function name, in any case: <c>_xlfn.</c>, and the
    /// <c>_xlws.</c> of a worksheet-only function after it.
    /// </summary>
    /// <param name="name">A function name as the formula text writes it.</param>
    /// <param name="bare">The name without its prefix, or <paramref name="name"/> when it has none.</param>
    /// <returns><c>true</c> when <paramref name="name"/> had a future-function prefix.</returns>
    internal static bool TryStripFuturePrefix(ReadOnlySpan<char> name, out ReadOnlySpan<char> bare)
    {
        if (!name.StartsWith(FuturePrefix, StringComparison.OrdinalIgnoreCase))
        {
            bare = name;
            return false;
        }

        bare = name[FuturePrefix.Length..];
        if (bare.StartsWith(WorksheetFuturePrefix, StringComparison.OrdinalIgnoreCase))
            bare = bare[WorksheetFuturePrefix.Length..];

        return true;
    }

    /// <summary>
    /// Formula text has no leading <c>=</c>: Excel shows one, and a file does not store it. A caller
    /// can type one, so it comes off before the parser sees the text.
    /// </summary>
    internal static string WithoutLeadingEquals(string text)
        => text.Length > 0 && text[0] == '=' ? text[1..] : text;

    /// <summary>
    /// Runs one parse, and returns a refusal as a value. This is the only place in XLibur that catches
    /// <see cref="ParsingException"/>, which the parser throws for text it cannot read. It also catches
    /// <see cref="ExpressionParseException"/>, which a factory throws while the parser reads the text,
    /// for a part the calc engine cannot take (#543). Only <paramref name="parse"/> runs inside the
    /// <c>try</c>, so nothing else becomes a refusal, and any other exception reaches the caller.
    /// </summary>
    private static bool TryParse<TState, TResult>(string text, TState state, Func<TState, TResult> parse,
        [MaybeNullWhen(false)] out TResult result, out FormulaRefusal refusal)
    {
        try
        {
            result = parse(state);
            refusal = default;
            return true;
        }
        catch (ParsingException ex)
        {
            result = default;
            refusal = new FormulaRefusal(text, ex);
            return false;
        }
        catch (ExpressionParseException ex)
        {
            result = default;
            refusal = new FormulaRefusal(text, ex);
            return false;
        }
    }

    /// <summary>
    /// Replace colons inside single-bracket structured reference column names with
    /// <see cref="ColonPlaceholder"/> so the parser does not treat them as range operators.
    /// <para>
    /// Single-bracket references like <c>Table[Some Header: Other]</c> contain a literal
    /// column name. Double-bracket references like <c>Table[[Col1]:[Col2]]</c> use the colon
    /// as a range separator and are left untouched.
    /// </para>
    /// <para>
    /// Single pass over a <see cref="ReadOnlySpan{T}"/>; the working buffer is leased from
    /// <see cref="ArrayPool{T}"/> only after the first qualifying colon is found, so
    /// formulas that have a colon outside any single-bracket reference (the common case
    /// for plain ranges like <c>SUM(A1:A10)</c>) cost nothing beyond the scan.
    /// </para>
    /// </summary>
    private static string ProtectStructuredRefColons(string formula, out bool wasProtected)
    {
        wasProtected = false;

        // Quick check: if no colon at all, nothing to protect.
        if (formula.IndexOf(':') < 0)
            return formula;

        var input = formula.AsSpan();
        char[]? rentedArray = null;
        Span<char> buffer = default;
        var i = 0;
        while (i < input.Length)
        {
            var c = input[i];

            if (c == '"')
            {
                i = SkipQuoted(input, i, '"');
                continue;
            }

            if (c == '\'')
            {
                i = SkipQuoted(input, i, '\'');
                continue;
            }

            if (c == '[')
            {
                i = ProcessBracket(formula, input, i, ref rentedArray, ref buffer);
                continue;
            }

            i++;
        }

        if (rentedArray is null)
            return formula;

        wasProtected = true;
        var result = new string(buffer);
        ArrayPool<char>.Shared.Return(rentedArray);
        return result;
    }

    private static int SkipQuoted(ReadOnlySpan<char> formula, int i, char quoteChar)
    {
        i++;
        while (i < formula.Length && formula[i] != quoteChar)
            i++;
        return i + 1;
    }

    private static int ProcessBracket(string sourceFormula, ReadOnlySpan<char> input, int i,
        ref char[]? rentedArray, ref Span<char> buffer)
    {
        var next = i + 1;
        if (next < input.Length && input[next] != '[' && input[next] != '#')
        {
            var j = next;
            while (j < input.Length && input[j] != ']')
            {
                if (input[j] == ':')
                {
                    if (rentedArray is null)
                    {
                        rentedArray = ArrayPool<char>.Shared.Rent(input.Length);
                        buffer = rentedArray.AsSpan(0, input.Length);
                        sourceFormula.AsSpan().CopyTo(buffer);
                    }

                    buffer[j] = ColonPlaceholder;
                }

                j++;
            }

            return j + 1;
        }

        return i + 1;
    }

    private static bool MightContainFutureFunction(ReadOnlySpan<char> formula)
    {
        for (var i = 0; i < formula.Length; ++i)
        {
            if (FutureFunctionSet.Value.IsPrefixOf(formula[i..]))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Hands the caller's factory every column name with its colon put back. The parser read the text
    /// with <see cref="ColonPlaceholder"/> in place of each colon inside a single-bracket column name,
    /// and a column name is the only part of a formula that can hold one: a colon in a string, or in a
    /// quoted sheet name, was never replaced.
    /// </summary>
    private sealed class ColonRestoringFactory<TScalar, TNode, TContext>(IAstFactory<TScalar, TNode, TContext> inner)
        : IAstFactory<TScalar, TNode, TContext>
    {
        public TNode StructureReference(TContext context, SymbolRange range, StructuredReferenceArea area,
            string? firstColumn, string? lastColumn)
            => inner.StructureReference(context, range, area, Restore(firstColumn), Restore(lastColumn));

        public TNode StructureReference(TContext context, SymbolRange range, string table,
            StructuredReferenceArea area, string? firstColumn, string? lastColumn)
            => inner.StructureReference(context, range, table, area, Restore(firstColumn), Restore(lastColumn));

        public TNode ExternalStructureReference(TContext context, SymbolRange range, int workbookIndex,
            string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
            => inner.ExternalStructureReference(context, range, workbookIndex, table, area, Restore(firstColumn),
                Restore(lastColumn));

        public TScalar LogicalValue(TContext context, SymbolRange range, bool value)
            => inner.LogicalValue(context, range, value);

        public TScalar NumberValue(TContext context, SymbolRange range, double value)
            => inner.NumberValue(context, range, value);

        public TScalar TextValue(TContext context, SymbolRange range, string text)
            => inner.TextValue(context, range, text);

        public TScalar ErrorValue(TContext context, SymbolRange range, ReadOnlySpan<char> error)
            => inner.ErrorValue(context, range, error);

        public TNode ArrayNode(TContext context, SymbolRange range, int rows, int columns,
            IReadOnlyList<TScalar> elements)
            => inner.ArrayNode(context, range, rows, columns, elements);

        public TNode BlankNode(TContext context, SymbolRange range)
            => inner.BlankNode(context, range);

        public TNode LogicalNode(TContext context, SymbolRange range, bool value)
            => inner.LogicalNode(context, range, value);

        public TNode ErrorNode(TContext context, SymbolRange range, ReadOnlySpan<char> error)
            => inner.ErrorNode(context, range, error);

        public TNode SheetErrorNode(TContext context, SymbolRange range, int? workbookIndex, string sheet,
            ReadOnlySpan<char> error)
            => inner.SheetErrorNode(context, range, workbookIndex, sheet, error);

        public TNode NumberNode(TContext context, SymbolRange range, double value)
            => inner.NumberNode(context, range, value);

        public TNode TextNode(TContext context, SymbolRange range, string text)
            => inner.TextNode(context, range, text);

        public TNode Reference(TContext context, SymbolRange range, ReferenceArea reference)
            => inner.Reference(context, range, reference);

        public TNode SheetReference(TContext context, SymbolRange range, string sheet, ReferenceArea reference)
            => inner.SheetReference(context, range, sheet, reference);

        public TNode BangReference(TContext context, SymbolRange range, ReferenceArea reference)
            => inner.BangReference(context, range, reference);

        public TNode Reference3D(TContext context, SymbolRange range, string firstSheet, string lastSheet,
            ReferenceArea reference)
            => inner.Reference3D(context, range, firstSheet, lastSheet, reference);

        public TNode ExternalSheetReference(TContext context, SymbolRange range, int workbookIndex, string sheet,
            ReferenceArea reference)
            => inner.ExternalSheetReference(context, range, workbookIndex, sheet, reference);

        public TNode ExternalReference3D(TContext context, SymbolRange range, int workbookIndex, string firstSheet,
            string lastSheet, ReferenceArea reference)
            => inner.ExternalReference3D(context, range, workbookIndex, firstSheet, lastSheet, reference);

        public TNode Function(TContext context, SymbolRange range, ReadOnlySpan<char> functionName,
            IReadOnlyList<TNode> arguments)
            => inner.Function(context, range, functionName, arguments);

        public TNode Function(TContext context, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName,
            IReadOnlyList<TNode> args)
            => inner.Function(context, range, sheetName, functionName, args);

        public TNode ExternalFunction(TContext context, SymbolRange range, int workbookIndex, string sheetName,
            ReadOnlySpan<char> functionName, IReadOnlyList<TNode> arguments)
            => inner.ExternalFunction(context, range, workbookIndex, sheetName, functionName, arguments);

        public TNode ExternalFunction(TContext context, SymbolRange range, int workbookIndex,
            ReadOnlySpan<char> functionName, IReadOnlyList<TNode> arguments)
            => inner.ExternalFunction(context, range, workbookIndex, functionName, arguments);

        public TNode CellFunction(TContext context, SymbolRange range, RowCol cell, IReadOnlyList<TNode> arguments)
            => inner.CellFunction(context, range, cell, arguments);

        public TNode Name(TContext context, SymbolRange range, string name)
            => inner.Name(context, range, name);

        public TNode SheetName(TContext context, SymbolRange range, string sheet, string name)
            => inner.SheetName(context, range, sheet, name);

        public TNode BangName(TContext context, SymbolRange range, string name)
            => inner.BangName(context, range, name);

        public TNode ExternalName(TContext context, SymbolRange range, int workbookIndex, string name)
            => inner.ExternalName(context, range, workbookIndex, name);

        public TNode ExternalSheetName(TContext context, SymbolRange range, int workbookIndex, string sheet,
            string name)
            => inner.ExternalSheetName(context, range, workbookIndex, sheet, name);

        public TNode ExternalDynamicDataExchange(TContext context, SymbolRange range, int workbookIndex, string item)
            => inner.ExternalDynamicDataExchange(context, range, workbookIndex, item);

        public TNode DynamicDataExchange(TContext context, SymbolRange range, string application, string topic,
            string item)
            => inner.DynamicDataExchange(context, range, application, topic, item);

        public TNode BinaryNode(TContext context, SymbolRange range, BinaryOperation operation, TNode leftNode,
            TNode rightNode)
            => inner.BinaryNode(context, range, operation, leftNode, rightNode);

        public TNode Unary(TContext context, SymbolRange range, UnaryOperation operation, TNode node)
            => inner.Unary(context, range, operation, node);

        public TNode Nested(TContext context, SymbolRange range, TNode node)
            => inner.Nested(context, range, node);

        private static string? Restore(string? column) => column?.Replace(ColonPlaceholder, ':');
    }

    /// <summary>
    /// All functions must have chars in the <c>.</c>-<c>_</c> range (trie range).
    /// </summary>
    private readonly record struct PrefixTree
    {
        private const char LowestChar = '.';
        private const char HighestChar = '_';

        /// <summary>
        /// Indicates the node represents a full prefix. Leaves are always ends and middle nodes
        /// sometimes (e.g. AB and ABC).
        /// </summary>
        private bool IsEnd { get; init; }

        /// <summary>
        /// Something transitions to this tree.
        /// </summary>
        [MemberNotNullWhen(false, nameof(Transitions))]
        private bool IsLeaf => Transitions is null;

        /// <summary>
        /// Index is a character minus <see cref="LowestChar"/>. The possible range of characters
        /// is from <see cref="LowestChar"/> to <see cref="HighestChar"/>.
        /// </summary>
        private PrefixTree[]? Transitions { get; init; }

        public static PrefixTree Build(IEnumerable<string> names)
        {
            var root = new PrefixTree { Transitions = new PrefixTree[HighestChar - LowestChar + 1] };
            foreach (var name in names)
                root.Insert(name.AsSpan());

            return root;
        }

        public bool IsPrefixOf(ReadOnlySpan<char> text)
        {
            var current = this;
            foreach (var c in text)
            {
                if (current.IsEnd)
                    return true;

                if (current.Transitions is null)
                    return false;

                var upperChar = char.ToUpperInvariant(c);
                if (upperChar is < LowestChar or > HighestChar)
                    return false;

                current = current.Transitions[upperChar - LowestChar];
            }

            return current.IsEnd;
        }

        private void Insert(ReadOnlySpan<char> functionName)
        {
            // Prev is necessary to update previous list due to immutability
            Debug.Assert(functionName.Length > 0);
            var prevTransitions = System.Array.Empty<PrefixTree>();
            var prevIndex = -1;
            var curNode = this;
            foreach (var c in functionName)
            {
                // All future function names are uppercase and in range, no need to transform.
                var transitionIndex = c - LowestChar;
                if (curNode.IsLeaf)
                {
                    // Current node is a leaf and thus has no transitions. Add them (kind of complicated thanks to readonly struct).
                    var currentTransitions = new PrefixTree[HighestChar - LowestChar + 1];
                    prevTransitions[prevIndex] = prevTransitions[prevIndex] with { Transitions = currentTransitions };
                    prevTransitions = currentTransitions;

                    // Move along the to a new node
                    curNode = currentTransitions[transitionIndex];
                }
                else
                {
                    prevTransitions = curNode.Transitions;
                    curNode = curNode.Transitions[transitionIndex];
                }

                prevIndex = transitionIndex;
            }

            prevTransitions[prevIndex] = prevTransitions[prevIndex] with { IsEnd = true };
        }
    }
}
