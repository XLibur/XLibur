using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Diagnostics;
using ClosedXML.Parser;

namespace XLibur.Excel.CalcEngine.Visitors;

internal static class FormulaTransformation
{
    /// <summary>
    /// A placeholder character (fullwidth colon U+FF1A) used to temporarily replace colons
    /// inside single-bracket structured reference column names (e.g. <c>Table[Some Header: Other]</c>)
    /// so the parser does not misinterpret them as range operators. The fullwidth colon is valid
    /// in the parser's grammar for column names but won't be treated as a range separator.
    /// </summary>
    private const char ColonPlaceholder = '\uFF1A';

    private static readonly Lazy<PrefixTree> FutureFunctionSet =
        new(() => PrefixTree.Build(XLConstants.FutureFunctionMap.Value.Keys));

    private static readonly RenameFunctionsVisitor RemapFutureFunctions = new(XLConstants.FutureFunctionMap);

    /// <summary>
    /// Add the necessary prefixes to a user-supplied future function without a prefix (e.g.
    /// <c>acot(A5)/2</c> to <c>_xlfn.ACOT(A5)/2</c>).
    /// </summary>
    internal static string FixFutureFunctions(string formula, string sheetName, XLSheetPoint origin)
    {
        // A preliminary check that formula might contain future function. There are two reasons to do this first:
        // * Although parsing is relatively cheap, it's not free. Checking for string is far cheaper.
        // * Risk management, parser might fail for some formulas and limit fallout in such case.
        if (!MightContainFutureFunction(formula.AsSpan()))
            return formula;

        return SafeModifyA1(formula, sheetName, origin.Row, origin.Column, RemapFutureFunctions);
    }

    /// <summary>
    /// Wrapper around FormulaConverter.ModifyA1 that protects colons inside
    /// single-bracket structured reference column names from being misinterpreted as range operators.
    /// </summary>
    internal static string SafeModifyA1(string formula, string sheetName, int row, int column, RefModVisitor visitor)
    {
        var protected_ = ProtectStructuredRefColons(formula, out var wasProtected);
        var result = FormulaConverter.ModifyA1(protected_, sheetName, row, column, visitor);
        return wasProtected ? result.Replace(ColonPlaceholder, ':') : result;
    }

    /// <summary>
    /// Wrapper around <see cref="FormulaConverter.ToR1C1"/> that protects colons inside
    /// single-bracket structured reference column names.
    /// </summary>
    internal static string SafeToR1C1(string formulaA1, int row, int column)
    {
        var protected_ = ProtectStructuredRefColons(formulaA1, out var wasProtected);
        var result = FormulaConverter.ToR1C1(protected_, row, column);
        return wasProtected ? result.Replace(ColonPlaceholder, ':') : result;
    }

    /// <summary>
    /// Wrapper around <see cref="FormulaConverter.ToA1"/> that protects colons inside
    /// single-bracket structured reference column names.
    /// </summary>
    internal static string SafeToA1(string formulaR1C1, int row, int column)
    {
        var protected_ = ProtectStructuredRefColons(formulaR1C1, out var wasProtected);
        var result = FormulaConverter.ToA1(protected_, row, column);
        return wasProtected ? result.Replace(ColonPlaceholder, ':') : result;
    }

    /// <summary>
    /// Replace colons inside single-bracket structured reference column names with a
    /// placeholder so the formula parser does not treat them as range operators.
    /// <para>
    /// Single-bracket references like <c>Table[Some Header: Other]</c> contain a literal
    /// column name. Double-bracket references like <c>Table[[Col1]:[Col2]]</c> use the colon
    /// as a range separator and are left untouched.
    /// </para>
    /// </summary>
    private static string ProtectStructuredRefColons(string formula, out bool wasProtected)
    {
        wasProtected = false;

        // Quick check: if no colon, nothing to protect.
        if (formula.IndexOf(':') < 0)
            return formula;

        char[]? chars = null;
        var i = 0;
        while (i < formula.Length)
        {
            var c = formula[i];

            // Skip string literals.
            if (c == '"')
            {
                i++;
                while (i < formula.Length && formula[i] != '"')
                    i++;
                i++; // skip closing quote
                continue;
            }

            // Skip single-quoted sheet name references (e.g. '[Book.xlsx]Sheet'!A1).
            if (c == '\'')
            {
                i++;
                while (i < formula.Length && formula[i] != '\'')
                    i++;
                i++; // skip closing quote
                continue;
            }

            // When we see '[', check if this is a single-bracket column reference.
            if (c == '[')
            {
                var next = i + 1;
                if (next < formula.Length && formula[next] != '[' && formula[next] != '#')
                {
                    // Inside a single-bracket structured reference column name.
                    // Replace colons with placeholders until the closing ']'.
                    var j = next;
                    while (j < formula.Length && formula[j] != ']')
                    {
                        if (formula[j] == ':')
                        {
                            chars ??= formula.ToCharArray();
                            chars[j] = ColonPlaceholder;
                            wasProtected = true;
                        }

                        j++;
                    }

                    i = j + 1;
                    continue;
                }
            }

            i++;
        }

        return wasProtected ? new string(chars!) : formula;
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
