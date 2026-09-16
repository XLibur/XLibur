using System;
using System.Collections.Generic;
using XLibur.Parser;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// Collects the precedents of a formula while the parser reads it, so that the dependency tree
/// builds no AST for the formula (#513, change 6).
/// </summary>
/// <remarks>
/// <para>
/// A node is a handle to the areas of a reference, kept in <see cref="DependenciesContext"/>, or
/// <see cref="None"/> when the node is not a reference. A value, an operator on values and a function
/// that gives a value are <see cref="None"/>, and they allocate nothing. Building the AST was about
/// 64% of what parsing allocated in a build: 159 MB of 248 MB for 200,000 formulas.
/// </para>
/// <para>
/// The rule for each kind of node is the one that <see cref="DependenciesVisitor"/> uses, in a method
/// that both call, for example <see cref="DependenciesVisitor.ApplyFunction{TArguments}"/>. The visitor
/// still reads the AST of a shared formula. The factory checks a function as the parser's own factory
/// does (<see cref="FormulaParser.ResolveFunction"/>), so both refuse the same formulas. A function with
/// the wrong number of arguments is a refusal for both, as text the parser cannot read is (#543).
/// </para>
/// <para>
/// The factory gets the arguments of a node after the parser has read them, so it has added their
/// precedents before it knows what the node does with them. The visitor does not read the arguments of
/// two nodes: a function from another cell that is not a function (it is <c>#REF!</c>), and a defined
/// name whose text the parser refuses. For these, the factory sets
/// <see cref="DependenciesContext.NeedsAst"/>, and the tree reads that formula through its AST.
/// </para>
/// </remarks>
internal sealed class PrecedentsFactory : IAstFactory<int, int, DependenciesContext>
{
    /// <summary>
    /// The node is not a reference.
    /// </summary>
    internal const int None = -1;

    /// <summary>
    /// The defined names whose formulas are being read, on the path from the cell formula down to the
    /// current node. See <see cref="DependenciesVisitor"/>.
    /// </summary>
    private readonly HashSet<XLDefinedName> _namesOnPath = new(ReferenceEqualityComparer.Instance);

    /// <summary>
    /// Read <paramref name="text"/>, and add its precedents to the dependencies of <paramref name="context"/>.
    /// </summary>
    /// <returns><c>false</c> when the parser refused the text. It may have added precedents before it did.</returns>
    internal bool TryCollect(string text, DependenciesContext context)
    {
        text = FormulaText.WithoutLeadingEquals(text);
        context.BeginWalk(text);
        if (!FormulaText.TryWalk(text, context, this, FormulaNotation.A1, out var root, out _))
            return false;

        if (root != None)
            context.AddAreas(context.GetNode(root));

        return true;
    }

    public int LogicalValue(DependenciesContext context, SymbolRange range, bool value) => 0;

    public int NumberValue(DependenciesContext context, SymbolRange range, double value) => 0;

    public int TextValue(DependenciesContext context, SymbolRange range, string text) => 0;

    public int ErrorValue(DependenciesContext context, SymbolRange range, ReadOnlySpan<char> error)
    {
        // Read as the parser's own factory reads it, so that the same text fails.
        _ = XLErrorParser.ParseFormulaError(error);
        return 0;
    }

    public int ArrayNode(DependenciesContext context, SymbolRange range, int rows, int columns,
        IReadOnlyList<int> elements) => None;

    public int BlankNode(DependenciesContext context, SymbolRange range) => None;

    public int LogicalNode(DependenciesContext context, SymbolRange range, bool value) => None;

    public int ErrorNode(DependenciesContext context, SymbolRange range, ReadOnlySpan<char> error)
    {
        _ = XLErrorParser.ParseFormulaError(error);
        return None;
    }

    public int SheetErrorNode(DependenciesContext context, SymbolRange range, int? workbookIndex, string sheet,
        ReadOnlySpan<char> error)
    {
        _ = XLErrorParser.ParseFormulaError(error);
        return None;
    }

    public int NumberNode(DependenciesContext context, SymbolRange range, double value) => None;

    public int TextNode(DependenciesContext context, SymbolRange range, string text) => None;

    public int Reference(DependenciesContext context, SymbolRange range, ReferenceArea reference)
        => Push(context, DependenciesVisitor.ReferenceOnSheet(context, context.FormulaArea.Name, reference));

    public int SheetReference(DependenciesContext context, SymbolRange range, string sheet, ReferenceArea reference)
        => Push(context, DependenciesVisitor.ReferenceOnSheet(context, sheet, reference));

    public int BangReference(DependenciesContext context, SymbolRange range, ReferenceArea reference)
        => Push(context, DependenciesVisitor.ReferenceOnSheet(context, context.FormulaArea.Name, reference));

    // A 3D reference is not supported yet, so it gives nothing, as in the visitor.
    public int Reference3D(DependenciesContext context, SymbolRange range, string firstSheet, string lastSheet,
        ReferenceArea reference) => None;

    // Book index 0 is this workbook. Another workbook gives nothing, as in the visitor.
    public int ExternalSheetReference(DependenciesContext context, SymbolRange range, int workbookIndex, string sheet,
        ReferenceArea reference)
        => workbookIndex == 0 ? Push(context, DependenciesVisitor.ReferenceOnSheet(context, sheet, reference)) : None;

    public int ExternalReference3D(DependenciesContext context, SymbolRange range, int workbookIndex, string firstSheet,
        string lastSheet, ReferenceArea reference) => None;

    public int Function(DependenciesContext context, SymbolRange range, ReadOnlySpan<char> functionName,
        IReadOnlyList<int> arguments) => Call(context, functionName.ToString(), arguments);

    public int Function(DependenciesContext context, SymbolRange range, string sheetName,
        ReadOnlySpan<char> functionName, IReadOnlyList<int> args) => Call(context, functionName.ToString(), args);

    public int ExternalFunction(DependenciesContext context, SymbolRange range, int workbookIndex, string sheetName,
        ReadOnlySpan<char> functionName, IReadOnlyList<int> arguments) => Call(context, functionName.ToString(), arguments);

    public int ExternalFunction(DependenciesContext context, SymbolRange range, int workbookIndex,
        ReadOnlySpan<char> functionName, IReadOnlyList<int> arguments) => Call(context, functionName.ToString(), arguments);

    public int CellFunction(DependenciesContext context, SymbolRange range, RowCol cell, IReadOnlyList<int> arguments)
    {
        // As in the parser's own factory: a name such as LOG10 reads as a cell, and is that function if
        // there is one, with no check of its arguments.
        var text = context.Text;
        var functionName = text.Substring(range.Start, text.IndexOf('(', range.Start) - range.Start);
        if (context.Workbook.CalcEngine.Functions.TryGetFunc(functionName, out _, out _))
            return Push(context, DependenciesVisitor.ApplyFunction(context, functionName, new ParsedArguments(context, arguments)));

        // Anything else is #REF!, and the visitor does not read its arguments. Their precedents are
        // added already, so the tree reads this formula through its AST.
        context.NeedsAst = true;
        return None;
    }

    public int StructureReference(DependenciesContext context, SymbolRange range, StructuredReferenceArea area,
        string? firstColumn, string? lastColumn)
        => Push(context, DependenciesVisitor.ApplyStructuredReference(context,
            new StructuredReferenceNode(null, null, area, firstColumn, lastColumn)));

    public int StructureReference(DependenciesContext context, SymbolRange range, string table,
        StructuredReferenceArea area, string? firstColumn, string? lastColumn)
        => Push(context, DependenciesVisitor.ApplyStructuredReference(context,
            new StructuredReferenceNode(null, table, area, firstColumn, lastColumn)));

    public int ExternalStructureReference(DependenciesContext context, SymbolRange range, int workbookIndex,
        string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        var prefix = new PrefixNode(new FileNode(workbookIndex), null, null, null);
        return Push(context, DependenciesVisitor.ApplyStructuredReference(context,
            new StructuredReferenceNode(prefix, table, area, firstColumn, lastColumn)));
    }

    public int Name(DependenciesContext context, SymbolRange range, string name)
        => NameNode(context, null, isThisWorkbookScope: false, name);

    public int SheetName(DependenciesContext context, SymbolRange range, string sheet, string name)
        => NameNode(context, sheet, isThisWorkbookScope: false, name);

    public int BangName(DependenciesContext context, SymbolRange range, string name)
        => NameNode(context, null, isThisWorkbookScope: false, name);

    // [0]!Name is a workbook-scoped name of this workbook. A name in another workbook gives nothing.
    public int ExternalName(DependenciesContext context, SymbolRange range, int workbookIndex, string name)
        => workbookIndex == 0 ? NameNode(context, null, isThisWorkbookScope: true, name) : None;

    public int ExternalSheetName(DependenciesContext context, SymbolRange range, int workbookIndex, string sheet,
        string name)
        => workbookIndex == 0 ? NameNode(context, sheet, isThisWorkbookScope: false, name) : None;

    public int ExternalDynamicDataExchange(DependenciesContext context, SymbolRange range, int workbookIndex,
        string item) => None;

    public int DynamicDataExchange(DependenciesContext context, SymbolRange range, string application, string topic,
        string item) => None;

    public int BinaryNode(DependenciesContext context, SymbolRange range, BinaryOperation operation, int leftNode,
        int rightNode)
        => Push(context, DependenciesVisitor.ApplyBinary(context, FormulaParser.ToBinaryOp(operation),
            context.GetNode(leftNode), context.GetNode(rightNode)));

    public int Unary(DependenciesContext context, SymbolRange range, UnaryOperation operation, int node)
        => Push(context, DependenciesVisitor.ApplyUnary(context, FormulaParser.ToUnaryOp(operation),
            context.GetNode(node)));

    public int Nested(DependenciesContext context, SymbolRange range, int node) => node;

    private static int Call(DependenciesContext context, string functionName, IReadOnlyList<int> arguments)
    {
        FormulaParser.ResolveFunction(context.Workbook.CalcEngine.Functions, ref functionName, arguments.Count);
        return Push(context, DependenciesVisitor.ApplyFunction(context, functionName, new ParsedArguments(context, arguments)));
    }

    private int NameNode(DependenciesContext context, string? sheet, bool isThisWorkbookScope, string name)
    {
        var areas = DependenciesVisitor.ApplyName(context, _namesOnPath, sheet, isThisWorkbookScope, name,
            new WalkedName(this, context));
        return Push(context, areas);
    }

    /// <summary>
    /// Read the formula of a defined name with this factory, inside the walk of the cell formula.
    /// </summary>
    private ReferenceAreas WalkName(DependenciesContext context, XLDefinedName definedName)
    {
        var outerText = context.Text;
        var text = FormulaText.WithoutLeadingEquals(definedName.RefersTo);
        context.Text = text;
        try
        {
            if (FormulaText.TryWalk(text, context, this, FormulaNotation.A1, out var root, out _))
                return context.GetNode(root);

            // The visitor reads nothing of a name whose text the parser refuses. This walk may have
            // added precedents before the refusal, so the tree reads the formula through its AST.
            context.NeedsAst = true;
            return ReferenceAreas.None;
        }
        finally
        {
            context.Text = outerText;
        }
    }

    private static int Push(DependenciesContext context, ReferenceAreas areas)
        => areas.IsReference ? context.PushNode(areas) : None;

    /// <summary>
    /// The arguments of a function, which the parser has read already.
    /// </summary>
    private readonly struct ParsedArguments(DependenciesContext context, IReadOnlyList<int> nodes)
        : DependenciesVisitor.IFunctionArguments
    {
        public int Count => nodes.Count;

        public ReferenceAreas Evaluate(int index) => context.GetNode(nodes[index]);
    }

    private readonly struct WalkedName(PrecedentsFactory factory, DependenciesContext context)
        : DependenciesVisitor.INameFormula
    {
        public ReferenceAreas Evaluate(XLDefinedName definedName) => factory.WalkName(context, definedName);
    }
}
