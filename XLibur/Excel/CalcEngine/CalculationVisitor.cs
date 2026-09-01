using System;
using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using ClosedXML.Parser;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Tables;

namespace XLibur.Excel.CalcEngine;

internal sealed class CalculationVisitor : IFormulaVisitor<CalcContext, AnyValue>
{
    private readonly FunctionRegistry _functions;
    private readonly ArrayPool<AnyValue> _argsPool;

    public CalculationVisitor(FunctionRegistry functions)
    {
        _functions = functions;
        _argsPool = ArrayPool<AnyValue>.Create(XLConstants.MaxFunctionArguments, 100);
    }

    public AnyValue Visit(CalcContext context, ScalarNode node)
    {
        return node.Value.ToAnyValue();
    }

    public AnyValue Visit(CalcContext context, ArrayNode node)
    {
        return node.Value;
    }

    public AnyValue Visit(CalcContext context, UnaryNode node)
    {
        var arg = node.Expression.Accept(context, this);

        return node.Operation switch
        {
            UnaryOp.Add => arg.UnaryPlus(),
            UnaryOp.Subtract => arg.UnaryMinus(context),
            UnaryOp.Percentage => arg.UnaryPercent(context),
            UnaryOp.SpillRange => EvaluateSpillRange(context, arg),
            UnaryOp.ImplicitIntersection => throw new NotImplementedException(
                "Excel 2016 implicit intersection is different from @ intersection of E2019+"),
            _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
        };
    }

    public AnyValue Visit(CalcContext context, BinaryNode node)
    {
        var leftArg = node.LeftExpression.Accept(context, this);
        var rightArg = node.RightExpression.Accept(context, this);

        return node.Operation switch
        {
            BinaryOp.Range => AnyValue.ReferenceRange(leftArg, rightArg, context),
            BinaryOp.Union => AnyValue.ReferenceUnion(leftArg, rightArg),
            BinaryOp.Intersection => throw new NotImplementedException(
                "Evaluation of range intersection operator is not implemented."),
            BinaryOp.Concat => AnyValue.Concat(leftArg, rightArg, context),
            BinaryOp.Add => AnyValue.BinaryPlus(leftArg, rightArg, context),
            BinaryOp.Sub => AnyValue.BinaryMinus(leftArg, rightArg, context),
            BinaryOp.Mult => AnyValue.BinaryMult(leftArg, rightArg, context),
            BinaryOp.Div => AnyValue.BinaryDiv(leftArg, rightArg, context),
            BinaryOp.Exp => AnyValue.BinaryExp(leftArg, rightArg, context),
            BinaryOp.Lt => AnyValue.IsLessThan(leftArg, rightArg, context),
            BinaryOp.Lte => AnyValue.IsLessThanOrEqual(leftArg, rightArg, context),
            BinaryOp.Eq => AnyValue.IsEqual(leftArg, rightArg, context),
            BinaryOp.Neq => AnyValue.IsNotEqual(leftArg, rightArg, context),
            BinaryOp.Gte => AnyValue.IsGreaterThanOrEqual(leftArg, rightArg, context),
            BinaryOp.Gt => AnyValue.IsGreaterThan(leftArg, rightArg, context),
            _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
        };
    }

    public AnyValue Visit(CalcContext context, FunctionNode node)
    {
        if (!_functions.TryGetFunc(node.Name, out var fn))
            return XLError.NameNotRecognized;

        var parameters = node.Parameters;
        var pool = _argsPool.Rent(parameters.Count);
        var args = new Span<AnyValue>(pool, 0, parameters.Count);

        // D38. An argument is evaluated in the context its function supplies, not the one the
        // formula started in, so operand intersection stops at the argument boundary. Excel is
        // finer-grained than this — it intersects inside SUM and MIN as well, and stops only at a
        // genuinely array-typed parameter such as SUMPRODUCT's — but XLibur has no data telling
        // SUM apart from SUMPRODUCT, and the suite's SUMPRODUCT/AVERAGE expectations are written
        // for array semantics. See specs/53 for the evidence and what closing the gap would cost.
        // Whatever the boundary, it must be restored rather than assumed false: a function call can
        // appear inside an operand of a top-level operator (`=A1+SUM(B1:B3)*C1:C3`), and the
        // operator after the call still has to intersect.
        var outerIntersectOperands = context.IntersectOperands;
        try
        {
            context.IntersectOperands = false;
            for (var i = 0; i < parameters.Count; ++i)
                args[i] = parameters[i].Accept(context, this);

            // The flag stays off for the call itself, not just for the arguments. A few functions
            // apply an operator to their own arguments from inside the body — SWITCH compares with
            // AnyValue.IsEqual — and an operator reached that way is not a top-level operator of
            // the formula, whatever the enclosing context was. The finally restores it.
            return !context.IsArrayCalculation
                ? fn!.CallFunction(context, args)
                : fn!.CallAsArray(context, args);
        }
        finally
        {
            context.IntersectOperands = outerIntersectOperands;
            _argsPool.Return(pool);
        }
    }

    public AnyValue Visit(CalcContext context, ReferenceNode node)
    {
        return node.GetReference(context);
    }

    public AnyValue Visit(CalcContext context, NameNode node)
    {
        return node.GetValue(context.Worksheet, context.CalcEngine);
    }

    public AnyValue Visit(CalcContext context, NotSupportedNode node)
        => throw new NotImplementedException($"Evaluation of {node.FeatureName} is not implemented.");

    public AnyValue Visit(CalcContext context, StructuredReferenceNode node)
    {
        if (!StructuredReferenceResolver.TryResolve(context, node, out var worksheet, out var range, out var error))
            return error;

        // The table's own sheet, not the formula's — a table name is workbook scoped and can be
        // referenced from anywhere in the workbook.
        return new Reference(XLRangeAddress.FromSheetRange(worksheet, range));
    }

    public AnyValue Visit(CalcContext context, PrefixNode node)
        => throw new InvalidOperationException("Node should never be visited.");

    public AnyValue Visit(CalcContext context, FileNode node)
        => throw new InvalidOperationException("Node should never be visited.");

    /// <summary>
    /// Evaluates the <c>#</c> spill-range operator (e.g. <c>A1#</c>): resolves the operand to a
    /// spill anchor and returns a <see cref="Reference"/> to that dynamic array's current
    /// footprint. Returns <c>#REF!</c> when the operand cell is not a spill anchor.
    /// </summary>
    private static AnyValue EvaluateSpillRange(CalcContext context, AnyValue operand)
    {
        // The operand of `#` must resolve to a single-cell reference: the spill anchor. A
        // multi-cell area (e.g. A1:B3#) is not a valid anchor, so it is a #REF!.
        if (!operand.TryPickArea(out var anchorArea, out var error))
            return error;

        if (anchorArea.FirstAddress.RowNumber != anchorArea.LastAddress.RowNumber ||
            anchorArea.FirstAddress.ColumnNumber != anchorArea.LastAddress.ColumnNumber)
            return XLError.CellReference;

        var sheet = anchorArea.Worksheet as XLWorksheet ?? context.Worksheet;
        var anchorRow = anchorArea.FirstAddress.RowNumber;
        var anchorColumn = anchorArea.FirstAddress.ColumnNumber;

        // Force the anchor to be current before reading its footprint: for a dirty anchor this
        // throws GettingDataException so the calc chain evaluates the anchor (spilling it and
        // updating its Range) before this formula. The returned value itself is unused.
        _ = context.GetCellValue(sheet, anchorRow, anchorColumn);

        var formula = sheet.Internals.CellsCollection.FormulaSlice.Get(new Point(anchorRow, anchorColumn));
        if (formula is null || !formula.IsDynamicArray)
            return XLError.CellReference; // #REF! — the cell is not a spill anchor.

        var footprint = formula.Range;
        if (footprint == default)
            return XLError.CellReference; // Anchor exists but hasn't produced a footprint yet.

        var rangeAddress = new XLRangeAddress(
            new XLAddress(sheet, footprint.TopRow, footprint.LeftColumn, true, true),
            new XLAddress(sheet, footprint.BottomRow, footprint.RightColumn, true, true));
        return new Reference(rangeAddress);
    }
}
