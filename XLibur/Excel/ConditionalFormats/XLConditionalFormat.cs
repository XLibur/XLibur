using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using XLibur.Excel.CalcEngine.Visitors;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel.ConditionalFormats;

internal sealed class XLConditionalFormat : XLStylizedBase, IXLConditionalFormat, IXLStylized
{
    private sealed class FullEqualityComparer : IEqualityComparer<IXLConditionalFormat>
    {
        private readonly bool _compareRange;
        private readonly DictionaryComparer<int, XLColor> _colorsComparer = new();
        private readonly EnumerableComparer<string?> _listComparer = new();
        private readonly DictionaryComparer<int, XLCFContentType> _contentsTypeComparer = new();
        private readonly DictionaryComparer<int, XLCFIconSetOperator> _iconSetTypeComparer = new();

        public FullEqualityComparer(bool compareRange)
        {
            _compareRange = compareRange;
        }

        public bool Equals(IXLConditionalFormat? x, IXLConditionalFormat? y)
        {
            if (x is not XLConditionalFormat xx) return y is null;
            if (y is not XLConditionalFormat yy) return false;
            if (ReferenceEquals(xx, yy)) return true;
            if (xx.GetType() != yy.GetType()) return false;

            var xxValues = xx.Values.Values.Where(v => v is not null and not { IsFormula: true }).Select(v => v.Value);
            var yyValues = yy.Values.Values.Where(v => v is not null and not { IsFormula: true }).Select(v => v.Value);

            var xStyle = xx.StyleValue;
            var yStyle = yy.StyleValue;

            return Equals(xStyle, yStyle)
                   && xx.CopyDefaultModify == yy.CopyDefaultModify
                   && xx.ConditionalFormatType == yy.ConditionalFormatType
                   && xx.TimePeriod == yy.TimePeriod
                   && xx.IconSetStyle == yy.IconSetStyle
                   && xx.Operator == yy.Operator
                   && xx.Bottom == yy.Bottom
                   && xx.Percent == yy.Percent
                   && xx.ReverseIconOrder == yy.ReverseIconOrder
                   && xx.StopIfTrue == yy.StopIfTrue
                   && xx.ShowIconOnly == yy.ShowIconOnly
                   && xx.ShowBarOnly == yy.ShowBarOnly
                   && xx.Gradient == yy.Gradient
                   && xx.BarAxisPosition == yy.BarAxisPosition
                   && xx.BarAxisColor == yy.BarAxisColor
                   && _listComparer.Equals(xxValues, yyValues)
                   && FormulasAreEqual(xx, yy)
                   && _colorsComparer.Equals(xx.Colors, yy.Colors)
                   && _contentsTypeComparer.Equals(xx.ContentTypes, yy.ContentTypes)
                   && _iconSetTypeComparer.Equals(xx.IconSetOperators, yy.IconSetOperators)
                   && (!_compareRange || Equals(xx.Ranges, yy.Ranges));
        }

        /// <summary>
        /// Do two formats hold the same formulas, each relative to its own anchor
        /// (<see cref="AnchorOf"/>)? A format that
        /// holds a formula the parser refuses equals no other: the references in that formula are
        /// unknown, so nothing says the two hold the same one, and consolidation never merges it.
        /// </summary>
        private bool FormulasAreEqual(XLConditionalFormat x, XLConditionalFormat y)
            => TryGetRelativeFormulas(x, out var xFormulas)
               && TryGetRelativeFormulas(y, out var yFormulas)
               && _listComparer.Equals(xFormulas, yFormulas);

        public int GetHashCode(IXLConditionalFormat obj)
        {
            var xx = (XLConditionalFormat)obj;
            var xValues = xx.Values.Values
                .Where(v => v is not null and not { IsFormula: true })
                .Select(v => v.Value);

            // A refused formula is hashed as it is written. A format holding one equals only itself,
            // so a value that does not change keeps the hash consistent with Equals.
            if (obj.Ranges.Count > 0)
            {
                IEnumerable<string> formulas = TryGetRelativeFormulas(xx, out var relative) && relative is not null
                    ? relative
                    : xx.Values.Values.Where(v => v is { IsFormula: true }).Select(v => v.Value);
                xValues = xValues.Union(formulas);
            }

            unchecked
            {
                var hashCode = xx.StyleValue.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.CopyDefaultModify.GetHashCode();
                hashCode = (hashCode * 397) ^ _listComparer.GetHashCode(xValues);
                hashCode = (hashCode * 397) ^ _colorsComparer.GetHashCode(xx.Colors);
                hashCode = (hashCode * 397) ^ _contentsTypeComparer.GetHashCode(xx.ContentTypes);
                hashCode = (hashCode * 397) ^ _iconSetTypeComparer.GetHashCode(xx.IconSetOperators);
                hashCode = (hashCode * 397) ^ (_compareRange ? xx.Ranges.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (int)xx.ConditionalFormatType;
                hashCode = (hashCode * 397) ^ (int)xx.TimePeriod;
                hashCode = (hashCode * 397) ^ (int)xx.IconSetStyle;
                hashCode = (hashCode * 397) ^ (int)xx.Operator;
                hashCode = (hashCode * 397) ^ xx.Bottom.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.Percent.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.ReverseIconOrder.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.ShowIconOnly.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.ShowBarOnly.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.Gradient.GetHashCode();
                hashCode = (hashCode * 397) ^ (int)xx.BarAxisPosition;
                hashCode = (hashCode * 397) ^ xx.BarAxisColor.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.StopIfTrue.GetHashCode();
                return hashCode;
            }
        }

        /// <summary>
        /// The format's formulas in R1C1, relative to its anchor (<see cref="AnchorOf"/>), or
        /// <c>null</c> when it has no range.
        /// </summary>
        /// <returns><c>false</c> when the parser refuses one of the formulas.</returns>
        private static bool TryGetRelativeFormulas(XLConditionalFormat format, out List<string>? formulas)
        {
            if (format.Areas.Count == 0)
            {
                formulas = null;
                return true;
            }

            var anchor = AnchorOf(format.Areas);
            formulas = [];
            foreach (var value in format.Values.Values)
            {
                if (value is not { IsFormula: true })
                    continue;

                if (!XLCellFormula.TryGetFormula(value.Value, FormulaConversionType.A1ToR1C1, anchor, out var r1c1, out _))
                    return false;

                formulas.Add(r1c1);
            }

            return true;
        }
    }

    /// <summary>
    /// Re-points each formula from <paramref name="baseCell"/> to <paramref name="targetCell"/>.
    /// </summary>
    /// <param name="baseCell">The cell the formulas are written relative to.</param>
    /// <param name="targetCell">The cell to write them relative to.</param>
    /// <param name="leaveRefusedUnchanged">
    /// What a formula the parser refuses means to the caller. Copying a format is a public edge, and
    /// there it throws <c>ExpressionParseException</c>, as copying a cell does. Consolidation leaves
    /// it exactly as it is (ADR 0002): its references are unknown, so there is nothing to re-point.
    /// </param>
    internal void AdjustFormulas(XLCell baseCell, XLCell targetCell, bool leaveRefusedUnchanged = false)
    {
        var keys = Values.Keys.ToList();
        foreach (var key in keys)
        {
            if (Values[key] == null || !Values[key].IsFormula)
                continue;

            if (!baseCell.TryGetFormulaR1C1(Values[key].Value, out var r1c1, out var refusal)
                || !targetCell.TryGetFormulaA1(r1c1, out var a1, out refusal))
            {
                if (leaveRefusedUnchanged)
                    continue;

                throw refusal.ToException();
            }

            Values[key] = new XLFormula { _value = a1, IsFormula = true };
        }
    }

    /// <summary>
    /// Rewrites each formula of the format the way <paramref name="rewrite"/> says, for a sheet rename
    /// or delete: an expression, and a scale's value whose type is <see cref="XLCFContentType.Formula"/>.
    /// Such a value keeps no <c>=</c>, so it is a formula by its type rather than by
    /// <see cref="XLFormula.IsFormula"/>. A formula the parser refuses keeps its text (ADR 0002).
    /// </summary>
    /// <param name="formulaSheetName">The sheet the format is on.</param>
    /// <param name="rewrite">What the rename or the delete does to formula text.</param>
    internal void RewriteSheet(string formulaSheetName, SheetRewrite rewrite)
    {
        foreach (var key in Values.Keys.ToList())
        {
            var formula = Values[key];
            if (!IsFormulaValue(key, formula))
                continue;

            // The rewrite does not move a reference, so any origin reads the formula the same way.
            if (!rewrite.TryRewrite(formula.Value, formulaSheetName, new Point(1, 1), out var rewritten)
                || rewritten == formula.Value)
                continue;

            Values[key] = new XLFormula { _value = rewritten, IsFormula = formula.IsFormula };
        }
    }

    /// <summary>
    /// Re-points the references in each formula of the format for a row or column insert or delete,
    /// through spec 25's shifter, as a cell formula's are (issue #499, D77). The formulas are the ones
    /// <see cref="RewriteSheet"/> rewrites.
    /// </summary>
    /// <remarks>
    /// A formula is written relative to the format's range, so a relative reference names the cell a
    /// cell formula in the range's anchor (<see cref="AnchorOf"/>) would: <c>A1&gt;0</c> on <c>E1</c>
    /// names <c>A1</c>. The
    /// shifter moves the reference with the cell it names, which keeps it relative when the range moves
    /// as well: a row inserted above row 1 gives <c>A2&gt;0</c> on <c>E2</c>.
    /// </remarks>
    internal void ShiftFormulas<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        foreach (var key in Values.Keys.ToList())
        {
            var formula = Values[key];
            if (!IsFormulaValue(key, formula)
                || !TryShiftFormula<TAxis>(formula.Value, _worksheet, in edit, out var shifted))
                continue;

            Values[key] = new XLFormula { _value = shifted, IsFormula = formula.IsFormula };
        }
    }

    /// <summary>
    /// Re-points the references in one formula of a conditional format on
    /// <paramref name="formulaSheet"/> for <paramref name="edit"/>, through spec 25's shifter. Used
    /// for a modelled format's formulas and for the text of an <c>x14</c> rule kept as it was loaded.
    /// </summary>
    /// <returns>
    /// <c>false</c> when the formula keeps its text: the edit reaches nothing it refers to, or the
    /// parser refuses it. A refused formula is skipped before it reaches the shifter, whose regex
    /// fallback would otherwise guess at it (ADR 0002), as <c>XLDefinedNames</c> skips one.
    /// </returns>
    internal static bool TryShiftFormula<TAxis>(string text, XLWorksheet formulaSheet, in SheetEdit edit,
        out string shifted)
        where TAxis : struct, IGridAxis
    {
        shifted = text;
        if (string.IsNullOrWhiteSpace(text))
            return false;

        // An edit on another sheet reaches only a reference that names that sheet, and a formula names
        // a sheet only by writing its name, so text without the name costs no parse. On most edits that
        // is every rule of every other sheet.
        if (edit.Sheet != formulaSheet && !MentionsSheet(text, edit.Sheet.Name))
            return false;

        // A formula that refers to no cell has nothing to move. Constants are common here: a
        // cell-value rule's operand, a top-N rule's rank.
        if (!FormulaReferences.TryForFormula(text, out var references, out _)
            || (references.References.Count == 0 && references.SheetReferences.Count == 0))
            return false;

        var result = default(TAxis).ShiftFormula(text, formulaSheet, edit.Range, edit.Shift);
        if (result == text)
            return false;

        shifted = result;
        return true;
    }

    /// <summary>
    /// Whether this format is a piece of a rule that a row or column edit cut apart (see
    /// <c>XLConditionalFormats.CutIntoPieces</c>). Consolidation leaves such a piece as it is: Excel
    /// writes each piece as a block of its own (<c>cf-partial-*.xlsx</c>), and merging two pieces whose
    /// formulas happen to read the same would change the blocks it wrote.
    /// </summary>
    internal bool IsSplitByEdit { get; set; }

    /// <summary>
    /// A copy of the format over <paramref name="areas"/>, with its priority, type, values and style:
    /// one piece of a rule an edit cut apart.
    /// </summary>
    internal XLConditionalFormat CopyOnto(XLAreaList areas)
    {
        var copy = new XLConditionalFormat(_worksheet)
        {
            Areas = areas,
            Priority = Priority,
            CopyDefaultModify = CopyDefaultModify,
        };
        copy.CopyFrom(this);
        return copy;
    }

    /// <summary>
    /// Rebases each formula of the format from the cell <paramref name="from"/> onto the cell
    /// <paramref name="to"/>: a relative reference keeps its offset from the cell, and an absolute one
    /// stays where it is. The formulas are the ones <see cref="ShiftFormulas{TAxis}"/> shifts.
    /// </summary>
    /// <remarks>
    /// For a row or column edit: the formulas are rebased from the range's anchor
    /// (<see cref="AnchorOf"/>) onto the origin of each piece the edit leaves, before they are shifted
    /// (see
    /// <c>XLConditionalFormats.CutIntoPieces</c>). <c>$A2&gt;5</c> on <c>A2:C10</c> is rebased onto
    /// <c>A3</c> as <c>$A3&gt;5</c>, and deleting row 2 shifts it back to <c>$A2&gt;5</c> on
    /// <c>A2:C9</c>, as Excel writes (<c>cf-anchor-*.xlsx</c>). Shifting without the rebase gives
    /// <c>#REF!&gt;5</c>.
    /// </remarks>
    internal void RebaseFormulas(Point from, Point to)
    {
        if (from == to)
            return;

        foreach (var key in Values.Keys.ToList())
        {
            var formula = Values[key];
            if (!IsFormulaValue(key, formula) || !TryRebaseFormula(formula.Value, from, to, out var rebased))
                continue;

            Values[key] = new XLFormula { _value = rebased, IsFormula = formula.IsFormula };
        }
    }

    /// <summary>
    /// Rebases one formula of a conditional format from the cell <paramref name="from"/> onto the cell
    /// <paramref name="to"/> (see <see cref="RebaseFormulas"/>). Used for a modelled format's formulas
    /// and for the text of an <c>x14</c> rule kept as it was loaded.
    /// </summary>
    /// <returns>
    /// <c>false</c> when the formula keeps its text: nothing in it is relative, or the parser refuses
    /// it, in which case its references are unknown and it is not guessed at (ADR 0002).
    /// </returns>
    internal static bool TryRebaseFormula(string text, Point from, Point to, out string rebased)
    {
        rebased = text;

        // Onto the cell it came from, a formula is what it was; the round trip could only respell it.
        if (from == to
            || string.IsNullOrWhiteSpace(text)
            || !XLCellFormula.TryGetFormula(text, FormulaConversionType.A1ToR1C1, from, out var r1c1, out _)
            || !XLCellFormula.TryGetFormula(r1c1, FormulaConversionType.R1C1ToA1, to, out var a1, out _)
            || a1 == text)
            return false;

        rebased = a1;
        return true;
    }

    /// <summary>
    /// The cell a conditional format's formulas are written relative to, its anchor: the top-left
    /// corner of the rectangle that bounds <paramref name="areas"/>. For one area, its first cell.
    /// </summary>
    /// <remarks>
    /// No Excel-written file has settled which cell Excel uses for a range of several areas: the
    /// <c>cf-anchor-*.xlsx</c> and <c>cf-partial-*.xlsx</c> fixtures hold only single-area ranges. The
    /// bounding top-left follows what consolidation already did on save
    /// (<c>ConditionalFormatsConsolidateTests.ConsolidateShiftsFormulaRelativelyToTopMostCell</c>), so
    /// that the equality comparer, consolidation and a row or column edit all read a format's formulas
    /// from one cell (issue #499).
    /// </remarks>
    /// <param name="areas">A range of at least one area.</param>
    internal static Point AnchorOf(XLAreaList areas)
    {
        var row = int.MaxValue;
        var column = int.MaxValue;
        foreach (var area in areas)
        {
            row = Math.Min(row, area.TopRow);
            column = Math.Min(column, area.LeftColumn);
        }

        return new Point(row, column);
    }

    /// <summary>
    /// Is the value at <paramref name="key"/> a formula? An expression's is, and so is a scale's value
    /// point whose type is <see cref="XLCFContentType.Formula"/>. Such a value keeps no <c>=</c>, so it
    /// is a formula by its type rather than by <see cref="XLFormula.IsFormula"/>.
    /// </summary>
    private bool IsFormulaValue(int key, [NotNullWhen(true)] XLFormula? formula)
        => formula is not null
           && (formula.IsFormula
               || (ContentTypes.TryGetValue(key, out var type) && type == XLCFContentType.Formula));

    /// <summary>Does <paramref name="text"/> write <paramref name="sheetName"/>, quoted or not?</summary>
    private static bool MentionsSheet(string text, string sheetName)
        => text.Contains(sheetName, StringComparison.OrdinalIgnoreCase)
           || (sheetName.Contains('\'')
               && text.Contains(sheetName.Replace("'", "''"), StringComparison.OrdinalIgnoreCase));

    private static readonly IEqualityComparer<IXLConditionalFormat> FullComparerInstance =
        new FullEqualityComparer(true);

    public static IEqualityComparer<IXLConditionalFormat> FullComparer => FullComparerInstance;

    private static readonly IEqualityComparer<IXLConditionalFormat> NoRangeComparerInstance =
        new FullEqualityComparer(false);

    public static IEqualityComparer<IXLConditionalFormat> NoRangeComparer => NoRangeComparerInstance;

    #region Constructors

    private readonly XLWorksheet _worksheet;

    private XLConditionalFormat(XLWorksheet worksheet)
        : base(XLStyle.Default.Value)
    {
        _worksheet = worksheet;
        Id = Guid.NewGuid();
        Areas = XLAreaList.Empty;
        Values = new XLDictionary<XLFormula>();
        Colors = new XLDictionary<XLColor>();
        ContentTypes = new XLDictionary<XLCFContentType>();
        IconSetOperators = new XLDictionary<XLCFIconSetOperator>();
    }

    public XLConditionalFormat(XLRange range, bool copyDefaultModify = false)
        : this(range.Worksheet)
    {
        Areas = new XLAreaList(Area.FromRangeAddress(range.RangeAddress));
        CopyDefaultModify = copyDefaultModify;
    }

    public XLConditionalFormat(IEnumerable<XLRange> ranges, bool copyDefaultModify = false)
        : this(WorksheetOf(ranges))
    {
        Areas = XLAreaList.FromRanges(ranges);
        CopyDefaultModify = copyDefaultModify;
    }

    public XLConditionalFormat(XLConditionalFormat conditionalFormat, IXLRange targetRange)
        : this(conditionalFormat, [targetRange])
    {
    }

    public XLConditionalFormat(XLConditionalFormat conditionalFormat, IEnumerable<IXLRange> targetRanges)
        : this(WorksheetOf(targetRanges))
    {
        Areas = XLAreaList.FromRanges(targetRanges);
        CopyFrom(conditionalFormat);
    }

    private static XLWorksheet WorksheetOf(IEnumerable<IXLRange> ranges)
    {
        var first = ranges.FirstOrDefault()
                    ?? throw new InvalidOperationException("XLConditionalFormat requires at least one range.");
        return (XLWorksheet)first.Worksheet;
    }

    #endregion Constructors

    public Guid Id { get; internal set; }

    /// <summary>
    /// Priority of formatting rule. Lower values have higher priority than higher values.
    /// Minimum value is 1. It is basically used for ordering of CF during saving.
    /// </summary>
    internal int Priority { get; set; }

    public bool CopyDefaultModify { get; set; }

    protected override IEnumerable<XLStylizedBase> Children
    {
        get { yield break; }
    }

    public override IXLRanges RangesUsed => new XLRanges();

    public XLDictionary<XLFormula> Values { get; private set; }

    public XLDictionary<XLColor> Colors { get; private set; }

    public XLDictionary<XLCFContentType> ContentTypes { get; private set; }

    public XLDictionary<XLCFIconSetOperator> IconSetOperators { get; private set; }

    public IXLRange Range
    {
        get => Areas.Count > 0
            ? MaterializeRange(Areas[0])
            : throw new InvalidOperationException("XLConditionalFormat requires at least one Range.");
        set => Areas = new XLAreaList(Area.FromRangeAddress(value.RangeAddress));
    }

    /// <summary>
    /// The conditional format's coverage, as a value-typed <see cref="XLAreaList"/>. This is the
    /// source of truth: coverage lives here rather than as live repository ranges, so structural
    /// (row/column insert &amp; delete) shifts run as pure area transforms and can never alias or
    /// double-shift (ClosedXML issue #2850). <see cref="Ranges"/> and <see cref="Range"/> are
    /// projections of it.
    /// </summary>
    internal XLAreaList Areas { get; private set; }

    /// <summary>
    /// Coverage materialized as ranges on the owning worksheet. A fresh snapshot each call —
    /// mutating the returned collection has no effect; change coverage via <see cref="Range"/> or
    /// <see cref="SetAreas"/>.
    /// </summary>
    public IXLRanges Ranges
    {
        get
        {
            var ranges = new XLRanges();
            foreach (var area in Areas)
                ranges.Add(MaterializeRange(area));
            return ranges;
        }
    }

    /// <summary>
    /// Replace the coverage. Used by the range shifter to write back a value-typed area transform.
    /// </summary>
    internal void SetAreas(XLAreaList areas) => Areas = areas;

    public IXLConditionalFormat SetRanges(IEnumerable<IXLRange> ranges)
    {
        ArgumentNullException.ThrowIfNull(ranges);

        var materialized = ranges as IReadOnlyCollection<IXLRange> ?? ranges.ToList();

        if (materialized.Count == 0)
            throw new ArgumentException("A conditional format must cover at least one range.", nameof(ranges));

        foreach (var range in materialized)
        {
            // Areas are bare rectangles interpreted against this rule's own sheet, so a range from
            // elsewhere would silently move the rule rather than being rejected.
            if (!ReferenceEquals(range.Worksheet, _worksheet))
            {
                throw new ArgumentException(
                    $"Range '{range.RangeAddress}' belongs to a different worksheet than the conditional format.",
                    nameof(ranges));
            }
        }

        Areas = XLAreaList.FromRanges(materialized);
        return this;
    }

    private XLRange MaterializeRange(Area area)
        => _worksheet.Range(area.TopRow, area.LeftColumn, area.BottomRow, area.RightColumn);

    public XLConditionalFormatType ConditionalFormatType { get; set; }

    public XLTimePeriod TimePeriod { get; set; }

    public XLIconSetStyle IconSetStyle { get; set; }

    public XLCFOperator Operator { get; set; }

    public bool Bottom { get; set; }

    public bool Percent { get; set; }

    public bool ReverseIconOrder { get; set; }

    public bool ShowIconOnly { get; set; }

    public bool ShowBarOnly { get; set; }

    public bool Gradient { get; set; } = true;

    public XLDataBarAxisPosition BarAxisPosition { get; set; } = XLDataBarAxisPosition.Automatic;

    public XLColor BarAxisColor { get; set; } = XLColor.Black;

    public bool StopIfTrue { get; set; }

    public IXLConditionalFormat SetStopIfTrue()
    {
        return SetStopIfTrue(true);
    }

    public IXLConditionalFormat SetStopIfTrue(bool value)
    {
        StopIfTrue = value;
        return this;
    }

    public IXLConditionalFormat CopyTo(IXLWorksheet targetSheet)
    {
        if (targetSheet == Range.Worksheet)
            throw new InvalidOperationException(
                "Cannot copy conditional format to the worksheet it already belongs to.");
        var targetRanges = Ranges.Select(r => targetSheet.Range(((XLRangeAddress)r.RangeAddress).WithoutWorksheet()));
        var newCf = new XLConditionalFormat(this, targetRanges);
        targetSheet.ConditionalFormats.Add(newCf);
        return newCf;
    }

    public void CopyFrom(IXLConditionalFormat other)
    {
        InnerStyle = other.Style;
        ConditionalFormatType = other.ConditionalFormatType;
        TimePeriod = other.TimePeriod;
        IconSetStyle = other.IconSetStyle;
        Operator = other.Operator;
        Bottom = other.Bottom;
        Percent = other.Percent;
        ReverseIconOrder = other.ReverseIconOrder;
        ShowIconOnly = other.ShowIconOnly;
        ShowBarOnly = other.ShowBarOnly;
        Gradient = other.Gradient;
        BarAxisPosition = other.BarAxisPosition;
        BarAxisColor = other.BarAxisColor;
        StopIfTrue = other.StopIfTrue;

        Values.Clear();
        other.Values.Where(kp => kp.Value != null).ForEach(kp => Values.Add(kp.Key, new XLFormula(kp.Value)));
        CopyDictionary(Colors, other.Colors);
        CopyDictionary(ContentTypes, other.ContentTypes);
        CopyDictionary(IconSetOperators, other.IconSetOperators);
    }

    private static void CopyDictionary<T>(XLDictionary<T> target, XLDictionary<T> source) where T : notnull
    {
        target.Clear();
        source.ForEach(kp => target.Add(kp.Key, kp.Value));
    }

    public IXLStyle WhenIsBlank()
    {
        ConditionalFormatType = XLConditionalFormatType.IsBlank;
        return Style;
    }

    public IXLStyle WhenNotBlank()
    {
        ConditionalFormatType = XLConditionalFormatType.NotBlank;
        return Style;
    }

    public IXLStyle WhenIsError()
    {
        ConditionalFormatType = XLConditionalFormatType.IsError;
        return Style;
    }

    public IXLStyle WhenNotError()
    {
        ConditionalFormatType = XLConditionalFormatType.NotError;
        return Style;
    }

    public IXLStyle WhenDateIs(XLTimePeriod timePeriod)
    {
        TimePeriod = timePeriod;
        ConditionalFormatType = XLConditionalFormatType.TimePeriod;
        return Style;
    }

    public IXLStyle WhenContains(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        ConditionalFormatType = XLConditionalFormatType.ContainsText;
        Operator = XLCFOperator.Contains;
        return Style;
    }

    public IXLStyle WhenNotContains(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        ConditionalFormatType = XLConditionalFormatType.NotContainsText;
        Operator = XLCFOperator.NotContains;
        return Style;
    }

    public IXLStyle WhenStartsWith(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        ConditionalFormatType = XLConditionalFormatType.StartsWith;
        Operator = XLCFOperator.StartsWith;
        return Style;
    }

    public IXLStyle WhenEndsWith(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        ConditionalFormatType = XLConditionalFormatType.EndsWith;
        Operator = XLCFOperator.EndsWith;
        return Style;
    }

    public IXLStyle WhenEquals(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        Operator = XLCFOperator.Equal;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenEquals(double value)
    {
        Values.Initialize(new XLFormula(value));
        Operator = XLCFOperator.Equal;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenNotEquals(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        Operator = XLCFOperator.NotEqual;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenNotEquals(double value)
    {
        Values.Initialize(new XLFormula(value));
        Operator = XLCFOperator.NotEqual;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenGreaterThan(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        Operator = XLCFOperator.GreaterThan;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenGreaterThan(double value)
    {
        Values.Initialize(new XLFormula(value));
        Operator = XLCFOperator.GreaterThan;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenLessThan(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        Operator = XLCFOperator.LessThan;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenLessThan(double value)
    {
        Values.Initialize(new XLFormula(value));
        Operator = XLCFOperator.LessThan;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenEqualOrGreaterThan(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        Operator = XLCFOperator.EqualOrGreaterThan;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenEqualOrGreaterThan(double value)
    {
        Values.Initialize(new XLFormula(value));
        Operator = XLCFOperator.EqualOrGreaterThan;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenEqualOrLessThan(string value)
    {
        Values.Initialize(new XLFormula { Value = value });
        Operator = XLCFOperator.EqualOrLessThan;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenEqualOrLessThan(double value)
    {
        Values.Initialize(new XLFormula(value));
        Operator = XLCFOperator.EqualOrLessThan;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenBetween(string minValue, string maxValue)
    {
        Values.Initialize(new XLFormula { Value = minValue });
        Values.Add(new XLFormula { Value = maxValue });
        Operator = XLCFOperator.Between;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenBetween(double minValue, double maxValue)
    {
        Values.Initialize(new XLFormula(minValue));
        Values.Add(new XLFormula(maxValue));
        Operator = XLCFOperator.Between;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenNotBetween(string minValue, string maxValue)
    {
        Values.Initialize(new XLFormula { Value = minValue });
        Values.Add(new XLFormula { Value = maxValue });
        Operator = XLCFOperator.NotBetween;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenNotBetween(double minValue, double maxValue)
    {
        Values.Initialize(new XLFormula(minValue));
        Values.Add(new XLFormula(maxValue));
        Operator = XLCFOperator.NotBetween;
        ConditionalFormatType = XLConditionalFormatType.CellIs;
        return Style;
    }

    public IXLStyle WhenIsDuplicate()
    {
        ConditionalFormatType = XLConditionalFormatType.IsDuplicate;
        return Style;
    }

    public IXLStyle WhenIsUnique()
    {
        ConditionalFormatType = XLConditionalFormatType.IsUnique;
        return Style;
    }

    public IXLStyle WhenIsTrue(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
            throw new ArgumentException("Formula cannot be null, empty, or whitespace.", nameof(formula));

        var trimmed = formula.TrimStart();
        var f = trimmed[0] == '=' ? formula : "=" + formula;
        Values.Initialize(new XLFormula { Value = f });
        ConditionalFormatType = XLConditionalFormatType.Expression;
        return Style;
    }

    public IXLStyle WhenIsTop(int value, XLTopBottomType topBottomType = XLTopBottomType.Items)
    {
        Values.Initialize(new XLFormula(value));
        Percent = topBottomType == XLTopBottomType.Percent;
        ConditionalFormatType = XLConditionalFormatType.Top10;
        Bottom = false;
        return Style;
    }

    public IXLStyle WhenIsBottom(int value, XLTopBottomType topBottomType)
    {
        Values.Initialize(new XLFormula(value));
        Percent = topBottomType == XLTopBottomType.Percent;
        ConditionalFormatType = XLConditionalFormatType.Top10;
        Bottom = true;
        return Style;
    }

    public IXLCFColorScaleMin ColorScale()
    {
        ConditionalFormatType = XLConditionalFormatType.ColorScale;
        return new XLCFColorScaleMin(this);
    }

    public IXLCFDataBarMin DataBar(XLColor color, bool showBarOnly = false, bool gradient = true)
    {
        Colors.Initialize(color);
        ShowBarOnly = showBarOnly;
        Gradient = gradient;
        ConditionalFormatType = XLConditionalFormatType.DataBar;
        return new XLCFDataBarMin(this);
    }

    public IXLCFDataBarMin DataBar(XLColor positiveColor, XLColor negativeColor, bool showBarOnly = false,
        bool gradient = true)
    {
        Colors.Initialize(positiveColor);
        Colors.Add(negativeColor);
        ShowBarOnly = showBarOnly;
        Gradient = gradient;
        ConditionalFormatType = XLConditionalFormatType.DataBar;
        return new XLCFDataBarMin(this);
    }

    public IXLCFIconSet IconSet(XLIconSetStyle iconSetStyle, bool reverseIconOrder = false, bool showIconOnly = false)
    {
        IconSetOperators.Clear();
        Values.Clear();
        ContentTypes.Clear();
        ConditionalFormatType = XLConditionalFormatType.IconSet;
        IconSetStyle = iconSetStyle;
        ReverseIconOrder = reverseIconOrder;
        ShowIconOnly = showIconOnly;
        return new XLCFIconSet(this);
    }
}

internal sealed class DictionaryComparer<TKey, TValue> :
    IEqualityComparer<Dictionary<TKey, TValue>>
    where TKey : notnull
{
    private readonly IEqualityComparer<TValue> _valueComparer;

    public DictionaryComparer(IEqualityComparer<TValue>? valueComparer = null)
    {
        _valueComparer = valueComparer ?? EqualityComparer<TValue>.Default;
    }

    public bool Equals(Dictionary<TKey, TValue>? x, Dictionary<TKey, TValue>? y)
    {
        if (x is null) return y is null;
        if (y is null) return false;
        if (x.Count != y.Count)
            return false;
        if (x.Keys.Except(y.Keys).Any())
            return false;
        if (y.Keys.Except(x.Keys).Any())
            return false;
        return x.All(pair => _valueComparer.Equals(pair.Value, y[pair.Key]));
    }

    public int GetHashCode(Dictionary<TKey, TValue>? obj)
    {
        if (obj is null)
            return 0;

        unchecked
        {
            var hash = 0;
            foreach (var pair in obj)
            {
                var entryHash = pair.Key.GetHashCode();
                entryHash = (entryHash * 397) ^ (pair.Value is not null ? _valueComparer.GetHashCode(pair.Value) : 0);
                hash ^= entryHash;
            }
            return hash;
        }
    }
}

internal sealed class EnumerableComparer<T> : IEqualityComparer<IEnumerable<T>>
{
    private readonly IEqualityComparer<T> _valueComparer;

    public EnumerableComparer(IEqualityComparer<T>? valueComparer = null)
    {
        _valueComparer = valueComparer ?? EqualityComparer<T>.Default;
    }

    public bool Equals(IEnumerable<T>? x, IEnumerable<T>? y)
    {
        if (x is null) return y is null;
        return y is not null && SetEquals(x, y, _valueComparer);
    }

    public int GetHashCode(IEnumerable<T>? obj)
    {
        if (obj is null)
            return 0;

        unchecked
        {
            var hash = 0;
            foreach (var item in obj)
                hash ^= item is not null ? _valueComparer.GetHashCode(item) : 0;
            return hash;
        }
    }

    private static bool SetEquals(IEnumerable<T> first, IEnumerable<T> second,
        IEqualityComparer<T>? comparer)
    {
        return new HashSet<T>(second, comparer ?? EqualityComparer<T>.Default)
            .SetEquals(first);
    }
}
