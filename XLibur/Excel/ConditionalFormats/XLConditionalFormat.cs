using System;
using System.Collections.Generic;
using System.Linq;
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
        /// Do two formats hold the same formulas, each relative to its own first cell? A format that
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
        /// The format's formulas in R1C1, relative to its first cell, or <c>null</c> when it has no
        /// range.
        /// </summary>
        /// <returns><c>false</c> when the parser refuses one of the formulas.</returns>
        private static bool TryGetRelativeFormulas(XLConditionalFormat format, out List<string>? formulas)
        {
            if (format.Ranges.Count == 0)
            {
                formulas = null;
                return true;
            }

            var anchor = (XLCell)format.Ranges.First().FirstCell();
            formulas = [];
            foreach (var value in format.Values.Values)
            {
                if (value is not { IsFormula: true })
                    continue;

                if (!anchor.TryGetFormulaR1C1(value.Value, out var r1c1, out _))
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
