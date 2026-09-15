using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Parser;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// <para>
/// Visit each node and determine all ranges that might affect the formula.
/// It uses concrete values (e.g. actual range for structured references) and
/// should be refreshed when a structured reference or name is changed in a workbook.
/// </para>
/// <para>
/// The areas found by the visitor shouldn't change when data on a worksheet changes,
/// so the output is a superset of areas, if necessary.
/// </para>
/// <para>
/// Precedents visitor is not completely accurate; in case of uncertainty, it uses
/// a larger area. At worst, the end result is unnecessary recalculation. For simple
/// cases, it works fine and freaks like <c>A1:IF(Other!B5,B7,Different!G3)</c>
/// will be marked as dirty more often than strictly necessary.
/// </para>
/// <para>
/// Each node visitor evaluates if the output is a reference or a value/array. If
/// the result is an array, it propagates to upper nodes, where there can be things like
/// range operator. The result is a <see cref="ReferenceAreas"/>, which holds one area without
/// allocating anything.
/// </para>
/// <para>
/// The rule for each kind of node is a static <c>Apply</c> method, which <see cref="PrecedentsFactory"/>
/// calls too. The factory collects precedents while the parser reads a formula, without an AST, and the
/// shared rules keep the two from giving different precedents (#513).
/// </para>
/// </summary>
internal sealed class DependenciesVisitor : IFormulaVisitor<DependenciesContext, ReferenceAreas>
{
    /// <summary>
    /// The arguments of a function, each read once by <see cref="ApplyFunction{TArguments}"/>. The
    /// visitor reads an argument when the rule asks for it, and the factory has read them already.
    /// </summary>
    internal interface IFunctionArguments
    {
        int Count { get; }

        ReferenceAreas Evaluate(int index);
    }

    /// <summary>
    /// Reads the formula of a defined name for <see cref="ApplyName{TFormula}"/>.
    /// </summary>
    internal interface INameFormula
    {
        ReferenceAreas Evaluate(XLDefinedName definedName);
    }

    /// <summary>
    /// The defined names whose formulas are being visited, on the path from the cell formula down to
    /// the node being visited.
    /// </summary>
    /// <remarks>
    /// A name met again on its own path refers to itself, directly or through other names, and
    /// following it again would never end (D79). The set is the path, not every name seen so far: a
    /// name used twice in one formula, or reached from two branches, is followed from each, so each
    /// use adds its precedents. A visitor belongs to one <see cref="DependencyTree"/>, which is used
    /// from one thread at a time.
    /// </remarks>
    private readonly HashSet<XLDefinedName> _namesOnPath = new(ReferenceEqualityComparer.Instance);

    public ReferenceAreas Visit(DependenciesContext context, ScalarNode node)
    {
        // Scalar node can't contain sub-nodes or references.
        return ReferenceAreas.None;
    }

    public ReferenceAreas Visit(DependenciesContext context, ArrayNode node)
    {
        // Array node can't contain sub-nodes or references.
        return ReferenceAreas.None;
    }

    public ReferenceAreas Visit(DependenciesContext context, UnaryNode node)
    {
        return ApplyUnary(context, node.Operation, node.Expression.Accept(context, this));
    }

    public ReferenceAreas Visit(DependenciesContext context, BinaryNode node)
    {
        var leftAreas = node.LeftExpression.Accept(context, this);
        var rightAreas = node.RightExpression.Accept(context, this);
        return ApplyBinary(context, node.Operation, leftAreas, rightAreas);
    }

    public ReferenceAreas Visit(DependenciesContext context, FunctionNode node)
    {
        return ApplyFunction(context, node.Name, new VisitedArguments(this, context, node.Parameters));
    }

    public ReferenceAreas Visit(DependenciesContext context, NotSupportedNode node)
    {
        return ReferenceAreas.None;
    }

    public ReferenceAreas Visit(DependenciesContext context, ReferenceNode node)
    {
        var prefix = node.Prefix;
        string sheetName;
        if (prefix is not null)
        {
            // We don't support external references, so there is no way to depend on something
            // in different workbook at the moment. Book index 0 is this workbook.
            if (prefix.IsInOtherWorkbook)
                return ReferenceAreas.None;

            // 3D references are not supported yet, so don't propagate anything.
            if (prefix.FirstSheet is not null || prefix.LastSheet is not null)
                return ReferenceAreas.None;

            sheetName = prefix.Sheet ?? throw new InvalidOperationException("Prefix doesn't contain sheet.");
        }
        else
        {
            sheetName = context.FormulaArea.Name;
        }

        return ReferenceOnSheet(context, sheetName, node.ReferenceArea);
    }

    public ReferenceAreas Visit(DependenciesContext context, NameNode node)
    {
        // External references are not supported for names. Book index 0 is this workbook.
        if (node.Prefix is { IsInOtherWorkbook: true })
            return ReferenceAreas.None;

        return ApplyName(context, _namesOnPath, node.Prefix?.Sheet, node.Prefix is { IsThisWorkbookScope: true },
            node.Name, new ParsedNameFormula(this, context));
    }

    public ReferenceAreas Visit(DependenciesContext context, StructuredReferenceNode node)
    {
        return ApplyStructuredReference(context, node);
    }

    public ReferenceAreas Visit(DependenciesContext context, PrefixNode node)
    {
        throw new InvalidOperationException("Should never be called.");
    }

    public ReferenceAreas Visit(DependenciesContext context, FileNode node)
    {
        throw new InvalidOperationException("Should never be called.");
    }

    /// <summary>
    /// A reference to <paramref name="area"/> on a sheet. A relative reference is resolved against the
    /// anchor of the formula.
    /// </summary>
    internal static ReferenceAreas ReferenceOnSheet(DependenciesContext context, string sheetName, ReferenceArea area)
    {
        var anchor = context.FormulaArea.Area.FirstPoint;
        var sheetRange = area.ToSheetRange(anchor);
        return ReferenceAreas.Of(new SheetArea(sheetName, sheetRange));
    }

    internal static ReferenceAreas ApplyUnary(DependenciesContext context, UnaryOp operation, ReferenceAreas operand)
    {
        // If the operand of unary node is not a reference -> end immediately,
        // the operator can't modify a non-reference into a reference.
        if (!operand.IsReference)
            return ReferenceAreas.None;

        // Operand is a reference
        if (operation is UnaryOp.ImplicitIntersection or UnaryOp.SpillRange)
        {
            // The operand's area is propagated as-is: for `#` that is the spill anchor cell,
            // which is enough for dirty propagation because spilled cells only change when the
            // anchor re-evaluates (and the anchor's whole footprint is registered as its formula
            // area — see DependencyTree.CreateFrom). Implicit intersection only ever shrinks the
            // area, so propagating the operand is a safe over-approximation there too.

            // The reference must be propagated upward, because there could be
            // a range operator (e.g. `B7:@A1:A5`)
            return operand;
        }

        // Some other operator is applied to the reference -> reference is converted
        // to an array
        context.AddAreas(operand);
        return ReferenceAreas.None;
    }

    internal static ReferenceAreas ApplyBinary(DependenciesContext context, BinaryOp operation,
        ReferenceAreas leftAreas, ReferenceAreas rightAreas)
    {
        // Reference operation only makes sense if both sides are references.
        // Otherwise, the reference operation results in an error.
        if (leftAreas.IsReference && rightAreas.IsReference)
            return VisitBinaryReferenceOp(context, operation, leftAreas, rightAreas);

        // Both children aren't references, or only one is -> binary operation transforms it
        // to a non-reference, either value or #REF!
        AddAreasIfReference(context, leftAreas);
        AddAreasIfReference(context, rightAreas);
        return ReferenceAreas.None;
    }

    internal static ReferenceAreas ApplyFunction<TArguments>(DependenciesContext context, string name,
        TArguments arguments)
        where TArguments : IFunctionArguments
    {
        // According to grammar, ref functions are: CHOOSE, IF, INDEX, INDIRECT, OFFSET
        // Only these functions are allowed to return references, per grammar.
        // However, OFFSET and INDIRECT are volatile functions that always have to be
        // recalculated (=are always marked dirty).
        if (XLHelper.FunctionComparer.Equals(name, "IF"))
            return ApplyIf(context, arguments);

        if (XLHelper.FunctionComparer.Equals(name, "INDEX"))
            return ApplyIndex(context, arguments);

        if (XLHelper.FunctionComparer.Equals(name, "CHOOSE"))
            return ApplyChoose(context, arguments);

        // INDIRECT references are dynamic — can't determine at parse time. The Volatile flag
        // ensures recalculation. Accept args for their dependencies. All other functions can have
        // references as arguments, but not as an output value.
        for (var i = 0; i < arguments.Count; ++i)
            AddAreasIfReference(context, arguments.Evaluate(i));

        return ReferenceAreas.None;
    }

    internal static ReferenceAreas ApplyName<TFormula>(DependenciesContext context,
        HashSet<XLDefinedName> namesOnPath, string? prefixSheet, bool isThisWorkbookScope, string name,
        TFormula formula)
        where TFormula : INameFormula
    {
        // [0]!Name is this workbook's workbook-scoped name, whatever sheet the formula is on.
        if (isThisWorkbookScope)
        {
            context.AddName(new XLName(name));
            return context.Workbook.DefinedNamesInternal.TryGetScopedValue(name, out var thisBookName)
                ? Follow(thisBookName)
                : ReferenceAreas.None;
        }

        context.AddName(prefixSheet is not null ? new XLName(prefixSheet, name) : new XLName(name));

        // First, try to interpret name as a sheet scoped name.
        var sheetName = prefixSheet ?? context.FormulaArea.Name;
        if (context.Workbook.TryGetWorksheet(sheetName, out XLWorksheet? sheet) &&
            sheet.DefinedNames.TryGetScopedValue(name, out var sheetDefinedName))
        {
            return Follow(sheetDefinedName);
        }

        // Name is not a sheet scoped one, try workbook scoped one
        if (context.Workbook.DefinedNamesInternal.TryGetScopedValue(name, out var bookNamedRange))
        {
            return Follow(bookNamedRange);
        }

        // Name is not found in the workbook
        return ReferenceAreas.None;

        ReferenceAreas Follow(XLDefinedName definedName)
        {
            // A circular name adds nothing more where it meets itself: its precedents are already
            // being collected further up the path. Evaluating the name is what reports the cycle.
            if (!namesOnPath.Add(definedName))
                return ReferenceAreas.None;

            try
            {
                return formula.Evaluate(definedName);
            }
            finally
            {
                namesOnPath.Remove(definedName);
            }
        }
    }

    internal static ReferenceAreas ApplyStructuredReference(DependenciesContext context, StructuredReferenceNode node)
    {
        // Resolve to the area the table currently covers, the same way evaluation does. Like a
        // defined name, the answer changes when the table is resized, renamed, added or
        // deleted — DependencyTree already requires rebuilding on all four.
        if (!StructuredReferenceResolver.TryResolve(context, node, out var worksheet, out var range, out _))
        {
            // Unresolvable today (missing table or column) means there is no precedent to
            // register. The formula evaluates to #REF!, and whatever later makes the table
            // resolve rebuilds the tree.
            return ReferenceAreas.None;
        }

        // The precedent is on the table's sheet, which need not be the formula's — a table name
        // is workbook scoped. Propagated rather than added to the context, so an enclosing range
        // operator can combine it — the same contract the other reference-producing nodes follow.
        return ReferenceAreas.Of(new SheetArea(worksheet.Name, range));
    }

    private static ReferenceAreas VisitBinaryReferenceOp(
        DependenciesContext context, BinaryOp operation,
        ReferenceAreas leftAreas, ReferenceAreas rightAreas)
    {
        // Both sides are references — calculate new ranges and propagate.
        if (operation == BinaryOp.Union)
            return leftAreas.Concat(rightAreas);

        if (operation == BinaryOp.Range)
            return CombineRangeAreas(leftAreas, rightAreas);

        if (operation == BinaryOp.Intersection)
            return IntersectAreas(leftAreas, rightAreas);

        // Operand is not a reference one, so the reference is turned to an array of values.
        context.AddAreas(leftAreas);
        context.AddAreas(rightAreas);
        return ReferenceAreas.None;
    }

    private static ReferenceAreas CombineRangeAreas(ReferenceAreas leftAreas, ReferenceAreas rightAreas)
    {
        var operandAreas = new List<SheetArea>(leftAreas.Count + rightAreas.Count);
        for (var i = 0; i < leftAreas.Count; ++i)
            operandAreas.Add(leftAreas[i]);
        for (var i = 0; i < rightAreas.Count; ++i)
            operandAreas.Add(rightAreas[i]);

        var rangeResult = new List<SheetArea>();

        // Create a new range from both operands. It must deal with
        // the situation where there are multiple sheets for both operands,
        // e.g. `IF(G4,Sheet1!A1,Sheet2!A2):IF(H3,Sheet2!C4,Sheet1!C5)`
        // that creates a valid range.
        var sheetGroups = operandAreas.GroupBy(area => area.Name, XLHelper.SheetComparer);

        // There is no simple way to go through all paths, so try to find
        // the largest possible ranges that could be the result. For normal
        // operands (A1:B2:C3), it will work fine, and for freaks, it will
        // find the largest possible range that is a superset of the actual result.
        foreach (var sheetGroup in sheetGroups)
        {
            var sheetAreas = sheetGroup.ToList();
            if (sheetAreas.Count == 1)
                continue;

            var rangeArea = sheetAreas[0].Area;
            for (var i = 1; i < sheetAreas.Count; ++i)
                rangeArea = rangeArea.Range(sheetAreas[i].Area);

            rangeResult.Add(new SheetArea(sheetGroup.Key, rangeArea));
        }

        // It's enough to return the result of a range operation. Operands can
        // be discarded because they are included in the result.
        return ReferenceAreas.Of(rangeResult);
    }

    private static ReferenceAreas IntersectAreas(ReferenceAreas leftAreas, ReferenceAreas rightAreas)
    {
        // Intersection makes the range smaller, so it's rather hard to optimize
        // areas. We make a special case for the most frequent case.
        if (leftAreas.Count == 1 && rightAreas.Count == 1)
        {
            var intersection = leftAreas[0].Intersect(rightAreas[0]);

            // Propagate only the intersection, not operands. Even if operands
            // change, it doesn't affect the formula, because cells outside
            // intersection are never used.
            return intersection is not null ? ReferenceAreas.Of(intersection.Value) : ReferenceAreas.None;
        }

        // Anything else is too complicated and thus just propagate all references.
        return leftAreas.Concat(rightAreas);
    }

    private static ReferenceAreas ApplyIf<TArguments>(DependenciesContext context, TArguments arguments)
        where TArguments : IFunctionArguments
    {
        // Tested value is not propagated, it's evaluated as an argument.
        AddAreasIfReference(context, arguments.Evaluate(0));

        // If argument is reference and test is evaluated to TRUE,
        // the reference is returned => propagate.
        var valueIfTrueReference = arguments.Evaluate(1);
        var valueIfFalseReference = arguments.Count == 3
            ? arguments.Evaluate(2)
            : ReferenceAreas.None;

        if (valueIfFalseReference.IsReference && valueIfTrueReference.IsReference)
            return valueIfTrueReference.Concat(valueIfFalseReference);

        return valueIfFalseReference.IsReference ? valueIfFalseReference : valueIfTrueReference;
    }

    private static ReferenceAreas ApplyIndex<TArguments>(DependenciesContext context, TArguments arguments)
        where TArguments : IFunctionArguments
    {
        // Add argument references, INDEX can have 2 or 3 arguments.
        for (var i = 1; i < arguments.Count; ++i)
            AddAreasIfReference(context, arguments.Evaluate(i));

        // If an INDEX function indexes into an area, it returns a reference,
        // not a value. Either way, return the whole reference that is indexed,
        // even though it's larger than the actual function result.
        return arguments.Evaluate(0);
    }

    private static ReferenceAreas ApplyChoose<TArguments>(DependenciesContext context, TArguments arguments)
        where TArguments : IFunctionArguments
    {
        // Index argument is used to select value, so don't propagate.
        AddAreasIfReference(context, arguments.Evaluate(0));

        // Any of arguments can be propagated -> propagate all.
        var parametersReference = ReferenceAreas.None;
        for (var i = 1; i < arguments.Count; ++i)
        {
            var parameterReference = arguments.Evaluate(i);
            if (!parameterReference.IsReference)
                continue;

            parametersReference = parametersReference.IsReference
                ? parametersReference.Concat(parameterReference)
                : parameterReference;
        }

        return parametersReference;
    }

    private static void AddAreasIfReference(DependenciesContext context, in ReferenceAreas areas)
    {
        if (areas.IsReference)
            context.AddAreas(areas);
    }

    /// <summary>
    /// The parameters of a function node, each visited when the rule reads it.
    /// </summary>
    private readonly struct VisitedArguments(
        DependenciesVisitor visitor,
        DependenciesContext context,
        IReadOnlyList<ValueNode> parameters) : IFunctionArguments
    {
        public int Count => parameters.Count;

        public ReferenceAreas Evaluate(int index) => parameters[index].Accept(context, visitor);
    }

    /// <summary>
    /// Parses the formula of a defined name and visits it.
    /// </summary>
    private readonly struct ParsedNameFormula(DependenciesVisitor visitor, DependenciesContext context)
        : INameFormula
    {
        public ReferenceAreas Evaluate(XLDefinedName definedName)
        {
            // A load keeps a name whose text the parser refuses, so that one bad name cannot stop
            // the workbook from opening. Its references are unknown, so the formula that uses it
            // is taken to depend on every cell, as a refused cell formula is, instead of failing
            // the tree (#489).
            // The named range is stored as A1 and thus parsed as A1, but should be interpreted as R1C1
            if (!context.Workbook.CalcEngine.TryParse(definedName.RefersTo, out var ast))
            {
                context.Dependencies.MarkPrecedentsUnknown();
                return ReferenceAreas.None;
            }

            // If the formula returned a reference, propagate it, rather
            // than add to the context (required for `A1:name` ).
            return ast.AstRoot.Accept(context, visitor);
        }
    }
}
