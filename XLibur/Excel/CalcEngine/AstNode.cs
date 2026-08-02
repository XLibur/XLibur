using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Parser;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// Base class for all AST nodes. All AST nodes must be observably immutable: a node may
/// memoise something it derives from its own state (see <see cref="ReferenceNode"/>), but
/// nothing a visitor can see may change after construction.
/// </summary>
// S1694 suggests an interface, since Accept is the only member. The AST is deliberately a closed
// class hierarchy: an interface could be implemented by a struct, which would box on every visit,
// and ValueNode extends this to share that closure with its own subclasses.
#pragma warning disable S1694
internal abstract class AstNode
{
    /// <summary>
    /// Method to accept a visitor (=call a method of a visitor with the correct type of the node).
    /// </summary>
    public abstract TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor);
}

/// <summary>
/// A base class for all AST nodes that can be evaluated to produce a value.
/// </summary>
internal abstract class ValueNode : AstNode;

/// <summary>
/// AST node that contains a blank, logical, number, text or an error value.
/// </summary>
internal sealed class ScalarNode : ValueNode
{
    public ScalarNode(ScalarValue value)
    {
        Value = value;
    }

    public ScalarValue Value { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// AST node that contains a constant array. Array is at least 1x1.
/// </summary>
internal sealed class ArrayNode : ValueNode
{
    public ArrayNode(Array value)
    {
        Value = value;
    }

    public Array Value { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

internal enum UnaryOp
{
    Add,
    Subtract,
    Percentage,
    SpillRange,
    ImplicitIntersection
}

/// <summary>
/// Unary expression, e.g. +123
/// </summary>
internal sealed class UnaryNode : ValueNode
{
    public UnaryNode(UnaryOp operation, ValueNode expr)
    {
        Operation = operation;
        Expression = expr;
    }

    public UnaryOp Operation { get; }

    public ValueNode Expression { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

internal enum BinaryOp
{
    // Text operators
    Concat,
    // Arithmetic
    Add,
    Sub,
    Mult,
    Div,
    Exp,
    // Comparison operators
    Lt,
    Lte,
    Eq,
    Neq,
    Gte,
    Gt,
    // References operators
    Range,
    Union,
    Intersection
}

/// <summary>
/// Binary expression, e.g. 1+2
/// </summary>
internal sealed class BinaryNode : ValueNode
{
    public BinaryNode(BinaryOp operation, ValueNode exprLeft, ValueNode exprRight)
    {
        Operation = operation;
        LeftExpression = exprLeft;
        RightExpression = exprRight;
    }

    public BinaryOp Operation { get; }

    public ValueNode LeftExpression { get; }

    public ValueNode RightExpression { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// A function call, e.g. <c>SIN(0.5)</c>.
/// </summary>
internal sealed class FunctionNode : ValueNode
{
    public FunctionNode(string name, IReadOnlyList<ValueNode> parms) : this(null, name, parms)
    {
    }

    public FunctionNode(PrefixNode? prefix, string name, IReadOnlyList<ValueNode> parms)
    {
        Prefix = prefix;
        Name = name;
        Parameters = parms;
    }

    public PrefixNode? Prefix { get; }

    /// <summary>
    /// Name of the function.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// AST nodes for arguments of the function.
    /// </summary>
    public IReadOnlyList<ValueNode> Parameters { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// An placeholder node for AST nodes that are not yet supported in XLibur.
/// </summary>
internal sealed class NotSupportedNode : ValueNode
{
    public NotSupportedNode(string featureName)
    {
        FeatureName = featureName;
    }

    public string FeatureName { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// AST node for an reference to an external file in a formula.
/// </summary>
internal sealed class FileNode : AstNode
{
    /// <summary>
    /// If the file is references indirectly, numeric identifier of a file.
    /// </summary>
    public int? Numeric { get; }

    /// <summary>
    /// If a file is referenced directly, a path to the file on the disc/UNC/web link, .
    /// </summary>
    public string? Path { get; }

    public FileNode(string path)
    {
        Path = path;
    }

    public FileNode(int numeric)
    {
        Numeric = numeric;
    }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// AST node for prefix of a reference in a formula. Prefix is a specification where to look for a reference.
/// <list type="bullet">
/// <item>Prefix specifies a <c>Sheet</c> - used for references in the local workbook.</item>
/// <item>Prefix specifies a <c>FirstSheet</c> and a <c>LastSheet</c> - 3D reference, references uses all sheets between first and last.</item>
/// <item>Prefix specifies a <c>File</c>, no sheet is specified - used for named ranges in external file.</item>
/// <item>Prefix specifies a <c>File</c> and a <c>Sheet</c> - references looks for its address in the sheet of the file.</item>
/// </list>
/// </summary>
internal sealed class PrefixNode : AstNode
{
    public PrefixNode(FileNode? file, string? sheet, string? firstSheet, string? lastSheet)
    {
        File = file;
        Sheet = sheet;
        FirstSheet = firstSheet;
        LastSheet = lastSheet;
    }

    /// <summary>
    /// If prefix references data from another file, can be empty.
    /// </summary>
    public FileNode? File { get; }

    /// <summary>
    /// Name of the sheet, without ! or escaped quotes. Can be null in some cases e.g. reference to a named range in an another file).
    /// </summary>
    public string? Sheet { get; }

    /// <summary>
    /// If the prefix is for 3D reference, the name of the first sheet. Empty otherwise.
    /// </summary>
    public string? FirstSheet { get; }

    /// <summary>
    /// If the prefix is for 3D reference, the name of the last sheet. Empty otherwise.
    /// </summary>
    public string? LastSheet { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);

    internal OneOf<IXLWorksheet, XLError> GetWorksheet(XLWorkbook wb)
    {
        if (File is not null)
            return XLError.CellReference;

        if (FirstSheet is not null || LastSheet is not null)
            return XLError.CellReference;

        if (!wb.TryGetWorksheet(Sheet!, out XLWorksheet? worksheet))
            return XLError.CellReference;

        return OneOf<IXLWorksheet, XLError>.FromT0(worksheet);
    }
}

/// <summary>
/// AST node for a reference of an area in some sheet.
/// </summary>
internal sealed class ReferenceNode : ValueNode
{
    /// <summary>
    /// Resolved reference for the sheet-less form. The address does not depend on anything
    /// outside the node, so once built it is valid for the node's lifetime.
    /// </summary>
    private Reference? _sheetlessReference;

    /// <summary>
    /// Resolved reference for the prefixed form, together with the sheet it was resolved
    /// against, kept in one object so the pair is replaced as a unit.
    /// </summary>
    private SheetReference? _sheetReference;

    // Neither memo is synchronized. Evaluation is single-threaded — XLWorkbook is not
    // thread-safe — so this is not a claim about concurrent readers. It is worth noting that
    // both fields tolerate a lost update anyway, because neither is trusted on its own: the
    // sheet-less reference is immutable and every writer produces an equal one, and the
    // prefixed reference is only used after ReferenceEquals confirms it was resolved against
    // the sheet being asked about, so any other value is recomputed rather than misapplied.

    public ReferenceNode(PrefixNode? prefix, ReferenceArea referenceArea, bool isA1)
    {
        Prefix = prefix;
        Address = isA1 ? referenceArea.GetDisplayStringA1() : referenceArea.GetDisplayStringR1C1();
        ReferenceArea = referenceArea;
        IsA1 = isA1;
    }

    /// <summary>
    /// An optional prefix for reference item.
    /// </summary>
    public PrefixNode? Prefix { get; }

    /// <summary>
    /// An address of a reference that corresponds to <see cref="Type"/>. Always without a sheet (that is in the prefix).
    /// </summary>
    public string Address { get; }

    /// <summary>
    /// An area from a parser.
    /// </summary>
    public ReferenceArea ReferenceArea { get; }

    /// <summary>
    /// Is the reference in A1 style? If <c>false</c>, then it is R1C1.
    /// </summary>
    public bool IsA1 { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);

    public AnyValue GetReference(CalcContext ctx)
    {
        if (Prefix is null)
            return _sheetlessReference ??= new Reference(BuildAddress(null));

        if (!Prefix.GetWorksheet(ctx.Workbook).TryPickT0(out var ws, out var err))
            return err;

        // ASTs are shared between every cell holding the same formula text (see
        // ExpressionCache), and recalculation re-resolves every reference, so the same node
        // is resolved many times and almost always against the same sheet. Keyed on the
        // resolved sheet rather than cached outright: a rename or a delete-and-re-add changes
        // which sheet the prefix resolves to, and that must not serve the previous address.
        var sheet = (XLWorksheet)ws;
        var cached = _sheetReference;
        if (cached is not null && ReferenceEquals(cached.Sheet, sheet))
            return cached.Reference;

        var reference = new Reference(BuildAddress(sheet));
        _sheetReference = new SheetReference(sheet, reference);
        return reference;
    }

    /// <summary>
    /// Build the range address from the <see cref="ReferenceArea"/> the parser produced,
    /// rather than by re-parsing <see cref="Address"/> — which the constructor generated from
    /// that same area, so parsing it only recovers what is already known.
    /// </summary>
    /// <remarks>
    /// Only the A1 form is handled. <see cref="IsA1"/> is <c>false</c> only for ASTs built to
    /// rewrite R1C1 formula text, and those are never evaluated: the only caller that reaches
    /// here is <see cref="XLCalcEngine.Parse"/>, which always parses as A1. The previous
    /// string-parsing implementation could not resolve R1C1 either — <see cref="XLRangeAddress"/>
    /// does not parse that syntax — so this narrows nothing.
    /// </remarks>
    private XLRangeAddress BuildAddress(XLWorksheet? sheet)
    {
        var first = ReferenceArea.First;
        var second = ReferenceArea.Second;

        // An axis of type None means the other axis carries the reference (A:B has no row,
        // 1:5 has no column), so the missing axis spans the whole sheet.
        var (row1, fixedRow1) = Axis(first.RowType, first.RowValue, XLHelper.MinRowNumber);
        var (col1, fixedCol1) = Axis(first.ColumnType, first.ColumnValue, XLHelper.MinColumnNumber);
        var (row2, fixedRow2) = Axis(second.RowType, second.RowValue, XLHelper.MaxRowNumber);
        var (col2, fixedCol2) = Axis(second.ColumnType, second.ColumnValue, XLHelper.MaxColumnNumber);

        // The endpoints need not be the top-left and bottom-right corners (D4:A1, D1:A4), but
        // Reference requires a normalized address. Each axis is ordered independently, with
        // the fixed flag travelling with the coordinate it belongs to.
        if (row1 > row2)
        {
            (row1, row2) = (row2, row1);
            (fixedRow1, fixedRow2) = (fixedRow2, fixedRow1);
        }

        if (col1 > col2)
        {
            (col1, col2) = (col2, col1);
            (fixedCol1, fixedCol2) = (fixedCol2, fixedCol1);
        }

        return new XLRangeAddress(
            new XLAddress(sheet, row1, col1, fixedRow1, fixedCol1),
            new XLAddress(sheet, row2, col2, fixedRow2, fixedCol2));
    }

    private static (int Position, bool Fixed) Axis(ReferenceAxisType axisType, int value, int absent) => axisType switch
    {
        ReferenceAxisType.Absolute => (value, true),
        ReferenceAxisType.Relative => (value, false),
        ReferenceAxisType.None => (absent, false),
        _ => throw new NotSupportedException($"Unknown reference axis type {axisType}."),
    };

    private sealed record SheetReference(XLWorksheet Sheet, Reference Reference);
}

/// <summary>
/// A name node in the formula. Name can refers to a generic formula, in most cases a reference, but it can be any kind of calculation (e.g. <c>A1+7</c>).
/// </summary>
internal sealed class NameNode : ValueNode
{
    public NameNode(PrefixNode? prefix, string name)
    {
        Prefix = prefix;
        Name = name;
    }

    /// <summary>
    /// An optional prefix for reference item.
    /// </summary>
    public PrefixNode? Prefix { get; }

    public string Name { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);

    public AnyValue GetValue(XLWorksheet ctxWs, XLCalcEngine engine)
    {
        var worksheet = ctxWs;
        if (Prefix is not null)
        {
            if (!Prefix.GetWorksheet(ctxWs.Workbook).TryPickT0(out var ws, out var err))
                return err;

            worksheet = (XLWorksheet)ws;
        }

        if (!TryGetNameRange(worksheet, out var definedName))
            return XLError.NameNotRecognized;

        // Parser needs an equal sign for a union of ranges (or braces around formula)
        var nameFormula = definedName.RefersTo;
        nameFormula = nameFormula.StartsWith('=') ? nameFormula : "=" + nameFormula;
        return engine.EvaluateName(nameFormula, ctxWs);
    }

    internal bool TryGetNameRange(IXLWorksheet ws, [NotNullWhen(true)] out IXLDefinedName? definedName)
    {
        if (ws.DefinedNames.TryGetValue(Name, out var sheetDefinedName))
        {
            definedName = sheetDefinedName;
            return true;
        }

        if (ws.Workbook.DefinedNamesInternal.TryGetValue(Name, out var bookDefinedName))
        {
            definedName = bookDefinedName;
            return true;
        }

        definedName = null;
        return false;
    }
}

internal sealed class StructuredReferenceNode : ValueNode
{
    public StructuredReferenceNode(PrefixNode? prefix, string? table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        Prefix = prefix;
        Table = table;
        Area = area;
        FirstColumn = firstColumn;
        LastColumn = lastColumn;
    }

    /// <summary>
    /// Can be empty if no prefix available.
    /// </summary>
    public PrefixNode? Prefix { get; }

    /// <summary>
    /// Table of the reference. It can be empty, if formula using the reference is within
    /// the table itself (e.g. total formulas).
    /// </summary>
    public string? Table { get; }

    /// <summary>
    /// Area of the table that is considered for the range of cell of reference.
    /// </summary>
    public StructuredReferenceArea Area { get; }

    /// <summary>
    /// First column of column range. If the reference refers to the whole table,
    /// the value is null.
    /// </summary>
    public string? FirstColumn { get; }

    /// <summary>
    /// Last column of column range. If structured reference refers only to one column,
    /// it is same as <see cref="FirstColumn"/>. If the reference refers to the whole table,
    /// the value is null.
    /// </summary>
    public string? LastColumn { get; }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}
