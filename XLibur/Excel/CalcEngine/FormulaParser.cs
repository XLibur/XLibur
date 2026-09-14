using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Parser;
using XLibur.Extensions;

namespace XLibur.Excel.CalcEngine;

internal sealed class FormulaParser
{
    private readonly AstFactory _nodeFactoryA1;
    private readonly AstFactory _nodeFactoryR1C1;

    public FormulaParser(FunctionRegistry functionRegistry)
    {
        _nodeFactoryA1 = new AstFactory(functionRegistry, true);
        _nodeFactoryR1C1 = new AstFactory(functionRegistry, false);
    }

    /// <summary>
    /// Parse a formula into an abstract syntax tree.
    /// </summary>
    /// <param name="formula">The formula text. A leading <c>=</c> is allowed.</param>
    /// <param name="isA1">Whether the text is in A1 notation rather than R1C1.</param>
    /// <exception cref="ExpressionParseException">The parser refused the formula.</exception>
    public Formula GetAst(string formula, bool isA1)
    {
        // Evaluation is a public edge: a refused formula reaches the caller as ExpressionParseException.
        if (!TryGetAst(formula, isA1, out var ast, out var refusal))
            throw refusal.ToException();

        return ast;
    }

    /// <summary>
    /// Parse a formula into an abstract syntax tree, and return the parser's refusal as a value.
    /// </summary>
    /// <param name="formula">The formula text. A leading <c>=</c> is allowed.</param>
    /// <param name="isA1">Whether the text is in A1 notation rather than R1C1.</param>
    /// <param name="ast">The tree, when the parser accepted the formula.</param>
    /// <param name="refusal">Why the parser refused the formula, when it did.</param>
    /// <returns><c>false</c> when the parser refused the formula.</returns>
    public bool TryGetAst(string formula, bool isA1, [NotNullWhen(true)] out Formula? ast, out FormulaRefusal refusal)
    {
        formula = FormulaText.WithoutLeadingEquals(formula);
        var factory = isA1 ? _nodeFactoryA1 : _nodeFactoryR1C1;
        var notation = isA1 ? FormulaNotation.A1 : FormulaNotation.R1C1;

        if (!FormulaText.TryWalk(formula, formula, factory, notation, out var root, out refusal))
        {
            ast = null;
            return false;
        }

        ast = new Formula(formula, root);
        return true;
    }

    /// <summary>
    /// Factory to create an abstract syntax tree for a formula in A1 notation.
    /// </summary>
    private sealed class AstFactory : IAstFactory<ScalarValue, ValueNode, string>
    {
        private readonly FunctionRegistry _functionRegistry;
        private readonly bool _isA1;

        internal AstFactory(FunctionRegistry functionRegistry, bool isA1)
        {
            _functionRegistry = functionRegistry;
            _isA1 = isA1;
        }

        public ScalarValue LogicalValue(string context, SymbolRange range, bool value) => value;

        public ScalarValue NumberValue(string context, SymbolRange range, double value) => value;

        public ScalarValue TextValue(string context, SymbolRange range, string text) => text;

        public ScalarValue ErrorValue(string context, SymbolRange range, ReadOnlySpan<char> error)
        {
            return GetErrorValue(error);
        }

        public ValueNode ArrayNode(string context, SymbolRange range, int rows, int columns,
            IReadOnlyList<ScalarValue> elements)
        {
            var array = new LiteralArray(rows, columns, elements);
            return new ArrayNode(array);
        }

        public ValueNode BlankNode(string context, SymbolRange range)
        {
            return new ScalarNode(ScalarValue.Blank);
        }

        public ValueNode LogicalNode(string context, SymbolRange range, bool value)
        {
            return new ScalarNode(value);
        }

        public ValueNode ErrorNode(string context, SymbolRange range, ReadOnlySpan<char> error)
        {
            return new ScalarNode(GetErrorValue(error));
        }

        public ValueNode SheetErrorNode(string context, SymbolRange range, int? workbookIndex, string sheet,
            ReadOnlySpan<char> error)
        {
            // The sheet only matters to a rename; Sheet1!#REF! evaluates to #REF! all the same.
            return new ScalarNode(GetErrorValue(error));
        }

        public ValueNode NumberNode(string context, SymbolRange range, double value)
        {
            return new ScalarNode(value);
        }

        public ValueNode TextNode(string context, SymbolRange range, string text)
        {
            return new ScalarNode(text);
        }

        public ValueNode Reference(string context, SymbolRange range, ReferenceArea reference)
        {
            return new ReferenceNode(null, reference, _isA1);
        }

        public ValueNode SheetReference(string context, SymbolRange range, string sheet, ReferenceArea reference)
        {
            var prefixNode = new PrefixNode(null, sheet, null, null);
            return new ReferenceNode(prefixNode, reference, _isA1);
        }

        public ValueNode BangReference(string context, SymbolRange range, ReferenceArea reference)
        {
            // `!A1` means "on the sheet the formula is evaluated for", which is what a reference
            // without a sheet already gets: evaluation falls back to the context's sheet, and a
            // defined name is evaluated in the context of the cell that uses it.
            return new ReferenceNode(null, reference, _isA1);
        }

        public ValueNode Reference3D(string context, SymbolRange range, string firstSheet, string lastSheet,
            ReferenceArea reference)
        {
            var prefixNode = new PrefixNode(null, null, firstSheet, lastSheet);
            return new ReferenceNode(prefixNode, reference, _isA1);
        }

        public ValueNode ExternalSheetReference(string context, SymbolRange range, int workbookIndex, string sheet,
            ReferenceArea reference)
        {
            var fileNode = new FileNode(workbookIndex);
            var prefixNode = new PrefixNode(fileNode, sheet, null, null);
            return new ReferenceNode(prefixNode, reference, _isA1);
        }

        public ValueNode ExternalReference3D(string context, SymbolRange range, int workbookIndex, string firstSheet,
            string lastSheet, ReferenceArea reference)
        {
            var fileNode = new FileNode(workbookIndex);
            var prefixNode = new PrefixNode(fileNode, null, firstSheet, lastSheet);
            return new ReferenceNode(prefixNode, reference, _isA1);
        }

        public ValueNode Function(string context, SymbolRange range, ReadOnlySpan<char> functionName,
            IReadOnlyList<ValueNode> arguments)
        {
            return GetFunctionNode(null, functionName.ToString(), arguments);
        }

        public ValueNode Function(string context, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName,
            IReadOnlyList<ValueNode> args)
        {
            var prefixNode = new PrefixNode(null, sheetName, null, null);
            return GetFunctionNode(prefixNode, functionName.ToString(), args);
        }

        public ValueNode ExternalFunction(string context, SymbolRange range, int workbookIndex, string sheetName,
            ReadOnlySpan<char> functionName, IReadOnlyList<ValueNode> arguments)
        {
            var prefixNode = new PrefixNode(new FileNode(workbookIndex), sheetName, null, null);
            return GetFunctionNode(prefixNode, functionName.ToString(), arguments);
        }

        public ValueNode ExternalFunction(string context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName,
            IReadOnlyList<ValueNode> arguments)
        {
            var prefixNode = new PrefixNode(new FileNode(workbookIndex), null, null, null);
            return GetFunctionNode(prefixNode, functionName.ToString(), arguments);
        }

        public ValueNode CellFunction(string context, SymbolRange range, RowCol cell,
            IReadOnlyList<ValueNode> arguments)
        {
            // Grammar technically allows evaluating a function from a different cell. The intended
            // usage is likely for lambda functions. Excel (as of 2022) doesn't do that, so use preference
            // as LOG10. Parser doesn't know about names of functions, so names such as LOG10 will always end up
            // here.
            var functionName = context.Substring(range.Start, context.IndexOf('(', range.Start) - range.Start);
            if (_functionRegistry.TryGetFunc(functionName, out _, out _))
                return new FunctionNode(functionName, arguments);

            // Nonexistent function is evaluated to #NAME?, but cell function should be evaluated to #REF!
            return new ScalarNode(XLError.CellReference);
        }

        public ValueNode StructureReference(string context, SymbolRange range, StructuredReferenceArea area,
            string? firstColumn, string? lastColumn)
        {
            return new StructuredReferenceNode(null, null, area, firstColumn, lastColumn);
        }

        public ValueNode StructureReference(string context, SymbolRange range, string table, StructuredReferenceArea area,
            string? firstColumn, string? lastColumn)
        {
            return new StructuredReferenceNode(null, table, area, firstColumn, lastColumn);
        }

        public ValueNode ExternalStructureReference(string context, SymbolRange range, int workbookIndex, string table,
            StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        {
            return new StructuredReferenceNode(new PrefixNode(new FileNode(workbookIndex), null, null, null), table,
                area, firstColumn, lastColumn);
        }

        public ValueNode Name(string context, SymbolRange range, string name)
        {
            return new NameNode(null, name);
        }

        public ValueNode SheetName(string context, SymbolRange range, string sheet, string name)
        {
            var prefixNode = new PrefixNode(null, sheet, null, null);
            return new NameNode(prefixNode, name);
        }

        public ValueNode BangName(string context, SymbolRange range, string name)
        {
            // `!Total` means the name as seen from the sheet the formula is evaluated for: that
            // sheet's own `Total` if it has one, otherwise the workbook's. A name without a prefix
            // already resolves that way, against the sheet of the cell using the defined name.
            return new NameNode(null, name);
        }

        public ValueNode ExternalName(string context, SymbolRange range, int workbookIndex, string name)
        {
            var prefixNode = new PrefixNode(new FileNode(workbookIndex), null, null, null);
            return new NameNode(prefixNode, name);
        }

        public ValueNode ExternalSheetName(string context, SymbolRange range, int workbookIndex, string sheet, string name)
        {
            var prefixNode = new PrefixNode(new FileNode(workbookIndex), sheet, null, null);
            return new NameNode(prefixNode, name);
        }

        public ValueNode ExternalDynamicDataExchange(string context, SymbolRange range, int workbookIndex, string item)
        {
            return new NotSupportedNode("dynamic data exchange");
        }

        public ValueNode DynamicDataExchange(string context, SymbolRange range, string application, string topic,
            string item)
        {
            return new NotSupportedNode("dynamic data exchange");
        }

        public ValueNode BinaryNode(string context, SymbolRange range, BinaryOperation operation, ValueNode leftNode,
            ValueNode rightNode)
        {
            var op = operation switch
            {
                BinaryOperation.Concat => BinaryOp.Concat,
                BinaryOperation.GreaterOrEqualThan => BinaryOp.Gte,
                BinaryOperation.LessOrEqualThan => BinaryOp.Lte,
                BinaryOperation.LessThan => BinaryOp.Lt,
                BinaryOperation.GreaterThan => BinaryOp.Gt,
                BinaryOperation.NotEqual => BinaryOp.Neq,
                BinaryOperation.Equal => BinaryOp.Eq,
                BinaryOperation.Addition => BinaryOp.Add,
                BinaryOperation.Subtraction => BinaryOp.Sub,
                BinaryOperation.Multiplication => BinaryOp.Mult,
                BinaryOperation.Division => BinaryOp.Div,
                BinaryOperation.Power => BinaryOp.Exp,
                BinaryOperation.Union => BinaryOp.Union,
                BinaryOperation.Intersection => BinaryOp.Intersection,
                BinaryOperation.Range => BinaryOp.Range,
                _ => throw new NotSupportedException($"'{operation}' is not a binary operation.")
            };

            return new BinaryNode(op, leftNode, rightNode);
        }

        public ValueNode Unary(string context, SymbolRange range, UnaryOperation operation, ValueNode node)
        {
            var op = operation switch
            {
                UnaryOperation.Plus => UnaryOp.Add,
                UnaryOperation.Minus => UnaryOp.Subtract,
                UnaryOperation.Percent => UnaryOp.Percentage,
                UnaryOperation.ImplicitIntersection => UnaryOp.ImplicitIntersection,
                UnaryOperation.SpillRange => UnaryOp.SpillRange,
                _ => throw new NotSupportedException($"'{operation}' is not a unary operation.")
            };
            return new UnaryNode(op, node);
        }

        public ValueNode Nested(string context, SymbolRange range, ValueNode node)
        {
            return node;
        }

        private FunctionNode GetFunctionNode(PrefixNode? prefixNode, string functionName,
            IReadOnlyList<ValueNode> argumentNodes)
        {
            var foundFunction = _functionRegistry.TryGetFunc(functionName, out var minParams, out var maxParams);

            // Functions are registered without the future-function prefix, so a prefixed name is looked
            // up again without it, whatever case the prefix is written in.
            if (!foundFunction && FormulaText.TryStripFuturePrefix(functionName, out var bareName))
            {
                functionName = bareName.ToString();
                foundFunction = _functionRegistry.TryGetFunc(functionName, out minParams, out maxParams);
            }

            // Even if we haven't found anything, don't crash. Missing function will be evaluated to `#NAME?`
            if (!foundFunction)
                return new FunctionNode(functionName, argumentNodes);

            if (minParams != -1 && argumentNodes.Count < minParams)
                throw new ExpressionParseException(
                    $"Too few parameters for function '{functionName}'. Expected a minimum of {minParams} and a maximum of {maxParams}.");

            if (maxParams != -1 && argumentNodes.Count > maxParams)
                throw new ExpressionParseException(
                    $"Too many parameters for function '{functionName}'.Expected a minimum of {minParams} and a maximum of {maxParams}.");

            return new FunctionNode(prefixNode, functionName, argumentNodes);
        }

        private static XLError GetErrorValue(ReadOnlySpan<char> error) => XLErrorParser.ParseFormulaError(error);
    }
}
