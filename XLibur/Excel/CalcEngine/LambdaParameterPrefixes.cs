using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLibur.Parser;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// Gives the parameters of <c>LET</c> and <c>LAMBDA</c> in a typed formula the prefixes Excel writes
/// in a file: <c>_xlpm.</c> before each parameter and before each use of it, and <c>_xlop.</c> for an
/// optional parameter, which is typed <c>[y]</c>. Excel refuses to open a workbook whose parameters
/// have none. One of Excel's functions written by its name as a value, <c>BYROW(D1:E2,SUM)</c>, gets
/// <c>_xleta.</c>, without which Excel reads the name as a defined name.
/// </summary>
/// <remarks>
/// <para>
/// A use is prefixed only inside the scope of its parameter, so a defined name of the same spelling
/// elsewhere in the formula stays as it is: a <c>LAMBDA</c>'s parameters are seen by its body, and
/// each <c>LET</c> name by the values after it and by the body. An inner parameter of the same name
/// hides an outer one. A name called as a function, as a <c>LET</c>-bound lambda is in
/// <c>LET(f,LAMBDA(x,x*2),f(3))</c>, is a use too, unless it is one of Excel's functions
/// (<see cref="ExcelFunctionNames"/>): <c>LET(SUM,LAMBDA(x,x*2),SUM(3))</c> calls <c>SUM</c>. These
/// are the rules Excel 365 followed for the formulas in <c>LambdaParameterPrefixTests</c>, typed over
/// COM.
/// </para>
/// <para>
/// The parser's <see cref="FormulaModifier"/> rewrites functions, sheets and references but not
/// names, and knows no scopes, so the formula is walked with a factory of its own. The factory
/// collects each name it meets, binds the names to the parameter lists of the <c>LET</c> or
/// <c>LAMBDA</c> around them as the parser closes each call, and records the text to insert. The
/// text is then edited at those places only, so the rest of the formula keeps the spelling it was
/// typed with.
/// </para>
/// </remarks>
internal static class LambdaParameterPrefixes
{
    private const string ParameterPrefix = "_xlpm.";
    private const string OptionalParameterPrefix = "_xlop.";

    /// <summary>
    /// <paramref name="formula"/>, A1 text without a leading <c>=</c>, with its parameters prefixed;
    /// or the text unchanged when it has no parameter to prefix, or the parser refuses it.
    /// </summary>
    /// <param name="formula">The formula.</param>
    /// <param name="isDefinedName">
    /// Whether a name is a defined name. A defined name with the spelling of one of Excel's functions,
    /// such as <c>Rate</c>, is that name, as Excel writes it. Null for no defined names.
    /// </param>
    internal static string Add(string formula, Func<string, bool>? isDefinedName = null)
    {
        // A string check first: a formula with none of these function names has no parameter, and one
        // with no name of Excel's functions written as a value, =SUM, has no function passed by its name.
        if (!FunctionsWithParameters.Any(name => formula.Contains(name, StringComparison.OrdinalIgnoreCase)) &&
            !MayPassFunctionByName(formula))
            return formula;

        var walk = new Walk(formula);
        if (!FormulaText.TryWalk(formula, walk, ScopeFactory.Instance, FormulaNotation.A1, out var root, out _))
            return formula;

        // A name no parameter binds that is one of Excel's functions is that function as a value, which
        // Excel writes in its own spelling: _xleta.RATE for a typed Rate that names nothing.
        foreach (var use in root.Free)
        {
            if (!use.IsCall && !use.IsPrefixed && IsFunctionName(use.Name) && isDefinedName?.Invoke(use.Name) != true)
                walk.Replace(use.Start, use.Name.Length, FormulaText.EtaPrefix + use.Name.ToUpperInvariant());
        }

        return walk.Edits.Count == 0 ? formula : walk.Apply();
    }

    /// <summary>
    /// Whether the text of <paramref name="formula"/> may hold one of Excel's functions written by its
    /// name as a value, such as <c>SUM</c> in <c>=ISERROR(SUM)</c> (#869): a name of one of them that is
    /// not called, not a sheet's, not in a string, a quoted sheet name or brackets, and not an error
    /// such as <c>#N/A</c>. A name that reads as a cell, <c>LOG10</c>, is left out, and so are
    /// <c>TRUE</c> and <c>FALSE</c>, which read as values. It only decides whether the formula is
    /// walked, which finds what the name is.
    /// </summary>
    private static bool MayPassFunctionByName(string formula)
    {
        var i = 0;
        while (i < formula.Length)
        {
            var start = i;
            i = SkipToken(formula, i, out var isName);
            if (isName && IsNameOfFunctionAsValue(formula, start, i))
                return true;
        }

        return false;
    }

    /// <summary>
    /// The index after the token at <paramref name="start"/>: a string or quoted sheet name, brackets,
    /// a name, or one other character. <paramref name="isName"/> is whether it is a name.
    /// </summary>
    private static int SkipToken(string formula, int start, out bool isName)
    {
        var c = formula[start];
        isName = char.IsLetter(c) || c == '_';
        if (c is '"' or '\'')
            return SkipQuoted(formula, start, c);

        if (c == '[')
            return SkipBrackets(formula, start);

        if (!isName)
            return start + 1;

        var i = start;
        while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '.' or '_'))
            i++;

        return i;
    }

    private static bool IsNameOfFunctionAsValue(string formula, int start, int end)
    {
        var isCalled = end < formula.Length && formula[end] is '(' or '!';
        var isOfSheetOrError = start > 0 && formula[start - 1] is '#' or '!' or ']';
        if (isCalled || isOfSheetOrError)
            return false;

        var name = formula.AsSpan(start, end - start);
        if (ReadsAsCell(name) ||
            name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            return false;

        return IsFunctionName(name);
    }

    /// <summary>Whether <paramref name="name"/> reads as a cell, as <c>LOG10</c> does: one to three letters, then digits.</summary>
    private static bool ReadsAsCell(ReadOnlySpan<char> name)
    {
        var letters = 0;
        while (letters < name.Length && char.IsAsciiLetter(name[letters]))
            letters++;

        return letters is >= 1 and <= 3 && letters < name.Length && name[letters..].IndexOfAnyExceptInRange('0', '9') < 0;
    }

    /// <summary>
    /// The index after the text that starts with <paramref name="quote"/> at <paramref name="start"/>, a
    /// doubled quote being one.
    /// </summary>
    private static int SkipQuoted(string formula, int start, char quote)
    {
        var i = start + 1;
        while (i < formula.Length)
        {
            if (formula[i] == quote && (i + 1 >= formula.Length || formula[i + 1] != quote))
                return i + 1;

            i += formula[i] == quote ? 2 : 1;
        }

        return i;
    }

    /// <summary>
    /// The index after the brackets that open at <paramref name="start"/>, which may nest:
    /// <c>Table1[[#Data],[Col]]</c>.
    /// </summary>
    private static int SkipBrackets(string formula, int start)
    {
        var depth = 0;
        var i = start;
        while (i < formula.Length)
        {
            if (formula[i] == '\'' && i + 1 < formula.Length)
            {
                // An apostrophe escapes the bracket after it: Table1[Col'[name].
                i += 2;
                continue;
            }

            depth += formula[i] switch { '[' => 1, ']' => -1, _ => 0 };
            i++;
            if (depth == 0)
                return i;
        }

        return i;
    }

    /// <summary>The functions that bind parameters, and those that take a function, by name or a <c>LAMBDA</c>.</summary>
    private static readonly string[] FunctionsWithParameters = ["LET", "LAMBDA", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY"];

    /// <summary>
    /// Whether <paramref name="name"/> is one of Excel's functions that can be a value. <c>LET</c> and
    /// <c>LAMBDA</c> cannot, and Excel writes them bare where a function goes.
    /// </summary>
    private static bool IsFunctionName(ReadOnlySpan<char> name) =>
        ExcelFunctionNames.IsBuiltIn(name) &&
        !name.Equals("LET", StringComparison.OrdinalIgnoreCase) &&
        !name.Equals("LAMBDA", StringComparison.OrdinalIgnoreCase);

    /// <summary>The formula being walked, and the edits that prefix its parameters.</summary>
    private sealed class Walk(string formula)
    {
        internal string Formula { get; } = formula;

        internal List<(int Start, int Length, string Text)> Edits { get; } = [];

        internal void Insert(int start, string text) => Edits.Add((start, 0, text));

        internal void Replace(int start, int length, string text) => Edits.Add((start, length, text));

        internal string Apply()
        {
            var sb = new StringBuilder(Formula.Length + (Edits.Count * ParameterPrefix.Length));
            var at = 0;
            foreach (var (start, length, text) in Edits.OrderBy(e => e.Start))
            {
                sb.Append(Formula, at, start - at).Append(text);
                at = start + length;
            }

            return sb.Append(Formula, at, Formula.Length - at).ToString();
        }
    }

    /// <summary>
    /// A name the walk met that no parameter list has bound yet, at <c>Start</c> in the formula, written
    /// as a value or, <paramref name="IsCall"/>, called as a function.
    /// </summary>
    private readonly record struct Use(int Start, string Name, bool IsPrefixed, bool IsCall = false);

    /// <summary>
    /// A node that could stand in a parameter list: a name, or <c>[y]</c>, an optional parameter
    /// (which the parser reads as a structured reference with no table).
    /// </summary>
    private readonly record struct Declaration(int Start, int Length, string Name, bool IsOptional, bool IsWritten);

    /// <summary>What a node of the formula holds for the walk: its unbound names, and what it declares if it is a parameter.</summary>
    private sealed class Part
    {
        internal static readonly Part Empty = new([], null);

        internal Part(IReadOnlyList<Use> free, Declaration? declaration)
        {
            Free = free;
            Declaration = declaration;
        }

        internal IReadOnlyList<Use> Free { get; }

        internal Declaration? Declaration { get; }

        internal static Part Merge(IReadOnlyList<Part> parts)
        {
            List<Use>? free = null;
            foreach (var part in parts)
            {
                if (part.Free.Count > 0)
                    (free ??= []).AddRange(part.Free);
            }

            return free is null ? Empty : new Part(free, null);
        }
    }

    private sealed class ScopeFactory : IAstFactory<object?, Part, Walk>
    {
        internal static readonly ScopeFactory Instance = new();

        public object? LogicalValue(Walk context, SymbolRange range, bool value) => null;

        public object? NumberValue(Walk context, SymbolRange range, double value) => null;

        public object? TextValue(Walk context, SymbolRange range, string text) => null;

        public object? ErrorValue(Walk context, SymbolRange range, ReadOnlySpan<char> error) => null;

        public Part ArrayNode(Walk context, SymbolRange range, int rows, int columns, IReadOnlyList<object?> elements) => Part.Empty;

        public Part BlankNode(Walk context, SymbolRange range) => Part.Empty;

        public Part LogicalNode(Walk context, SymbolRange range, bool value) => Part.Empty;

        public Part ErrorNode(Walk context, SymbolRange range, ReadOnlySpan<char> error) => Part.Empty;

        public Part SheetErrorNode(Walk context, SymbolRange range, int? workbookIndex, string sheet, ReadOnlySpan<char> error) => Part.Empty;

        public Part NumberNode(Walk context, SymbolRange range, double value) => Part.Empty;

        public Part TextNode(Walk context, SymbolRange range, string text) => Part.Empty;

        public Part Reference(Walk context, SymbolRange range, ReferenceArea reference) => Part.Empty;

        public Part SheetReference(Walk context, SymbolRange range, string sheet, ReferenceArea reference) => Part.Empty;

        public Part BangReference(Walk context, SymbolRange range, ReferenceArea reference) => Part.Empty;

        public Part Reference3D(Walk context, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference) => Part.Empty;

        public Part ExternalSheetReference(Walk context, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference) => Part.Empty;

        public Part ExternalReference3D(Walk context, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet,
            ReferenceArea reference) => Part.Empty;

        public Part Function(Walk context, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<Part> arguments)
        {
            FormulaText.TryStripFuturePrefix(functionName, out var bare);
            if (bare.Equals("LET", StringComparison.OrdinalIgnoreCase) && TryBindLet(context, arguments, out var let))
                return let;

            if (bare.Equals("LAMBDA", StringComparison.OrdinalIgnoreCase) && TryBindLambda(context, arguments, out var lambda))
                return lambda;

            // Any other function: its name may be a LET-bound lambda called by name, f(3).
            return Called(range.Start, functionName.ToString(), arguments);
        }

        public Part Function(Walk context, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName,
            IReadOnlyList<Part> args) => Part.Merge(args);

        public Part ExternalFunction(Walk context, SymbolRange range, int workbookIndex, string sheetName,
            ReadOnlySpan<char> functionName, IReadOnlyList<Part> arguments) => Part.Merge(arguments);

        public Part ExternalFunction(Walk context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName,
            IReadOnlyList<Part> arguments) => Part.Merge(arguments);

        // A name that reads as a cell called as a function, such as LOG10(x). Excel accepts a LET name
        // that reads as a cell (LET(A1,1,A1) is 1, written _xlpm.A1), but the parser reads such a
        // name as a reference, so a call of one is not taken for a parameter here; its arguments
        // may use one.
        public Part CellFunction(Walk context, SymbolRange range, RowCol cell, IReadOnlyList<Part> arguments) =>
            Part.Merge(arguments);

        public Part ExpressionCall(Walk context, SymbolRange range, Part target, IReadOnlyList<Part> arguments) =>
            Part.Merge([target, .. arguments]);

        public Part StructureReference(Walk context, SymbolRange range, StructuredReferenceArea area, string? firstColumn,
            string? lastColumn)
        {
            // [y], with no table and one column, is how an optional LAMBDA parameter is typed.
            return area == StructuredReferenceArea.None && firstColumn is not null && firstColumn == lastColumn
                ? new Part([], new Declaration(range.Start, range.Length, firstColumn, IsOptional: true, IsWritten: false))
                : Part.Empty;
        }

        public Part StructureReference(Walk context, SymbolRange range, string table, StructuredReferenceArea area,
            string? firstColumn, string? lastColumn) => Part.Empty;

        public Part ExternalStructureReference(Walk context, SymbolRange range, int workbookIndex, string table,
            StructuredReferenceArea area, string? firstColumn, string? lastColumn) => Part.Empty;

        public Part Name(Walk context, SymbolRange range, string name)
        {
            var isOptional = name.StartsWith(OptionalParameterPrefix, StringComparison.OrdinalIgnoreCase);
            var isPrefixed = isOptional || name.StartsWith(ParameterPrefix, StringComparison.OrdinalIgnoreCase);
            var bare = isPrefixed ? name[ParameterPrefix.Length..] : name;
            return new Part([new Use(range.Start, bare, isPrefixed)],
                new Declaration(range.Start, range.Length, bare, isOptional, IsWritten: isPrefixed));
        }

        public Part SheetName(Walk context, SymbolRange range, string sheet, string name) => Part.Empty;

        public Part BangName(Walk context, SymbolRange range, string name) => Part.Empty;

        public Part ExternalName(Walk context, SymbolRange range, int workbookIndex, string name) => Part.Empty;

        public Part ExternalSheetName(Walk context, SymbolRange range, int workbookIndex, string sheet, string name) => Part.Empty;

        public Part ExternalDynamicDataExchange(Walk context, SymbolRange range, int workbookIndex, string item) => Part.Empty;

        public Part DynamicDataExchange(Walk context, SymbolRange range, string application, string topic, string item) => Part.Empty;

        public Part BinaryNode(Walk context, SymbolRange range, BinaryOperation operation, Part leftNode, Part rightNode) =>
            Part.Merge([leftNode, rightNode]);

        public Part Unary(Walk context, SymbolRange range, UnaryOperation operation, Part node) => node;

        public Part Nested(Walk context, SymbolRange range, Part node) => node;

        /// <summary>
        /// A call of a function by name. The name is a use of a parameter, a lambda called by the name
        /// a LET gave it, unless it is one of Excel's functions: a call by such a name calls the
        /// function even where a parameter has the name.
        /// </summary>
        private static Part Called(int start, string functionName, IReadOnlyList<Part> arguments)
        {
            var merged = Part.Merge(arguments);
            FormulaText.TryStripFuturePrefix(functionName, out var unprefixed);
            if (ExcelFunctionNames.IsBuiltIn(unprefixed.ToString()))
                return merged;

            var isPrefixed = functionName.StartsWith(ParameterPrefix, StringComparison.OrdinalIgnoreCase);
            var bare = isPrefixed ? functionName[ParameterPrefix.Length..] : functionName;
            return new Part([new Use(start, bare, isPrefixed, IsCall: true), .. merged.Free], null);
        }

        /// <summary>
        /// <c>LET(name1, value1, name2, value2, …, body)</c>: each value sees the names before it, and
        /// the body sees them all. <c>false</c> when the arguments are not names and values in that
        /// shape, which leaves the call to be read as any other function's.
        /// </summary>
        private static bool TryBindLet(Walk walk, IReadOnlyList<Part> arguments, out Part bound)
        {
            bound = Part.Empty;
            if (arguments.Count < 3 || arguments.Count % 2 == 0)
                return false;

            for (var i = 0; i < arguments.Count - 1; i += 2)
            {
                if (arguments[i].Declaration is not { IsOptional: false })
                    return false;
            }

            var names = new List<string>();
            var free = new List<Use>();
            for (var i = 0; i < arguments.Count - 1; i += 2)
            {
                Bind(walk, arguments[i + 1].Free, names, free);
                var declaration = arguments[i].Declaration!.Value;
                Declare(walk, declaration);
                names.Add(declaration.Name);
            }

            Bind(walk, arguments[^1].Free, names, free);
            bound = new Part(free, null);
            return true;
        }

        /// <summary>
        /// <c>LAMBDA(parameter1, …, body)</c>: the body sees every parameter. <c>false</c> when an
        /// argument before the body is not a parameter.
        /// </summary>
        private static bool TryBindLambda(Walk walk, IReadOnlyList<Part> arguments, out Part bound)
        {
            bound = Part.Empty;
            if (arguments.Count == 0)
                return false;

            for (var i = 0; i < arguments.Count - 1; i++)
            {
                if (arguments[i].Declaration is null)
                    return false;
            }

            var names = new List<string>(arguments.Count - 1);
            for (var i = 0; i < arguments.Count - 1; i++)
            {
                var declaration = arguments[i].Declaration!.Value;
                Declare(walk, declaration);
                names.Add(declaration.Name);
            }

            var free = new List<Use>();
            Bind(walk, arguments[^1].Free, names, free);
            bound = new Part(free, null);
            return true;
        }

        /// <summary>Writes the prefix of a parameter where it is declared, unless it has one.</summary>
        private static void Declare(Walk walk, Declaration declaration)
        {
            if (declaration.IsWritten)
                return;

            if (declaration.IsOptional)
                walk.Replace(declaration.Start, declaration.Length, OptionalParameterPrefix + declaration.Name);
            else
                walk.Insert(declaration.Start, ParameterPrefix);
        }

        /// <summary>
        /// Prefixes each use in <paramref name="uses"/> that names one of <paramref name="names"/>, and
        /// adds the others to <paramref name="free"/> for an outer scope to bind.
        /// </summary>
        private static void Bind(Walk walk, IReadOnlyList<Use> uses, List<string> names, List<Use> free)
        {
            foreach (var use in uses)
            {
                if (!names.Contains(use.Name, StringComparer.OrdinalIgnoreCase))
                {
                    free.Add(use);
                    continue;
                }

                if (!use.IsPrefixed)
                    walk.Insert(use.Start, ParameterPrefix);
            }
        }
    }
}
