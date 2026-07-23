using System;
using System.Collections.Generic;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine
{
    /// <summary>
    /// Reference is a collection of cells in the workbook. It's used in formula evaluation.
    /// Every reference has at least one cell.
    /// </summary>
    /// <remarks>
    /// Inline-first storage: the typical reference has exactly one area, so the first area
    /// lives on the instance directly and only multi-area references allocate the
    /// <see cref="_additionalAreas"/> array. Saves the per-Reference list allocation that
    /// previously cost ~64 bytes (List header and one-element backing array) for the common
    /// single-area path.
    /// </remarks>
    internal sealed class Reference
    {
        private readonly XLRangeAddress _firstArea;
        private readonly XLRangeAddress[]? _additionalAreas;

        public Reference(XLRangeAddress area)
        {
            if (!area.IsNormalized)
                throw new ArgumentException("Range address must be normalized.", nameof(area));

            _firstArea = area;
        }

        /// <summary>
        /// Constructor that copies from a list. Pass a list with at least one normalized area.
        /// </summary>
        public Reference(List<XLRangeAddress> areas)
        {
            ArgumentNullException.ThrowIfNull(areas);
            if (areas.Count < 1)
                throw new ArgumentException("Reference must contain at least one area.", nameof(areas));

            _firstArea = areas[0];
            if (areas.Count > 1)
            {
                _additionalAreas = new XLRangeAddress[areas.Count - 1];
                for (var i = 1; i < areas.Count; i++)
                    _additionalAreas[i - 1] = areas[i];
            }
        }

        /// <summary>
        /// Constructor for the inline-first storage. <paramref name="additionalAreas"/> may be
        /// <c>null</c> or empty for a single-area reference; otherwise the array is taken
        /// over verbatim (do not mutate after construction).
        /// </summary>
        internal Reference(XLRangeAddress firstArea, XLRangeAddress[]? additionalAreas)
        {
            _firstArea = firstArea;
            _additionalAreas = additionalAreas is { Length: > 0 } ? additionalAreas : null;
        }

        public Reference(IXLRanges ranges)
        {
            ArgumentNullException.ThrowIfNull(ranges);
            var count = ranges.Count;
            if (count < 1)
                throw new ArgumentException("Reference must contain at least one range.", nameof(ranges));

            using var enumerator = ranges.GetEnumerator();
            enumerator.MoveNext();
            _firstArea = (XLRangeAddress)enumerator.Current.RangeAddress;

            if (count > 1)
            {
                _additionalAreas = new XLRangeAddress[count - 1];
                for (var i = 0; i < count - 1; i++)
                {
                    enumerator.MoveNext();
                    _additionalAreas[i] = (XLRangeAddress)enumerator.Current.RangeAddress;
                }
            }
        }

        /// <summary>
        /// Number of areas in the reference (always at least 1).
        /// </summary>
        public int AreaCount => _additionalAreas is null ? 1 : 1 + _additionalAreas.Length;

        /// <summary>
        /// Get the area at the specified index. Index 0 is always valid.
        /// </summary>
        public XLRangeAddress this[int index] =>
            index == 0 ? _firstArea : _additionalAreas![index - 1];

        /// <summary>
        /// Allocation-free struct enumerator over all areas.
        /// </summary>
        public Enumerator GetEnumerator() => new(this);

        /// <summary>
        /// Get the total number of cells covered by all areas (double counts overlapping areas).
        /// </summary>
        internal int NumberOfCells
        {
            get
            {
                var size = _firstArea.NumberOfCells;
                if (_additionalAreas is not null)
                {
                    foreach (var area in _additionalAreas)
                        size += area.NumberOfCells;
                }
                return size;
            }
        }

        /// <summary>
        /// An iterator over all nonblank cells of the range. Some cells can be iterated
        /// over multiple times (e.g., a union of two ranges with overlapping cells).
        /// </summary>
        public IEnumerable<ScalarValue> GetCellsValues(CalcContext ctx)
        {
            foreach (var area in this)
            {
                for (var row = area.FirstAddress.RowNumber; row <= area.LastAddress.RowNumber; ++row)
                {
                    for (var column = area.FirstAddress.ColumnNumber; column <= area.LastAddress.ColumnNumber; ++column)
                    {
                        var cellValue = ctx.GetCellValue(area.Worksheet, row, column);
                        if (!cellValue.IsBlank)
                        {
                            yield return cellValue;
                        }
                    }
                }
            }
        }

        public static OneOf<Reference, XLError> RangeOp(Reference lhs, Reference rhs, XLWorksheet contextWorksheet)
        {
            // Resolve the unique non-null worksheet on each side via direct loops; the original
            // Select/Where/Distinct/ToList chain allocated ~6 LINQ enumerators per binary op.
            // For single-area references the worksheet defaults to the area's own (no fallback
            // to context), preserving the previous behaviour of the lhs.Count == 1 branch.
            if (!TryResolveSingleWorksheet(lhs, contextWorksheet, out var lhsWorksheet))
                return XLError.IncompatibleValue;

            if (!TryResolveSingleWorksheet(rhs, contextWorksheet, out var rhsWorksheet))
                return XLError.IncompatibleValue;

            if (rhsWorksheet is not null && (lhsWorksheet ?? contextWorksheet) != rhsWorksheet)
                return XLError.IncompatibleValue;

            var minCol = XLHelper.MaxColumnNumber;
            var maxCol = 1;
            var minRow = XLHelper.MaxRowNumber;
            var maxRow = 1;
            foreach (var area in lhs)
                ExpandBoundingBox(area, ref minRow, ref maxRow, ref minCol, ref maxCol);
            foreach (var area in rhs)
                ExpandBoundingBox(area, ref minRow, ref maxRow, ref minCol, ref maxCol);

            var sheet = lhsWorksheet ?? rhsWorksheet;
            return new Reference(new XLRangeAddress(
                new XLAddress(sheet, minRow, minCol, false, false),
                new XLAddress(sheet, maxRow, maxCol, false, false)));

            static void ExpandBoundingBox(in XLRangeAddress area,
                ref int minRow, ref int maxRow, ref int minCol, ref int maxCol)
            {
                // Areas are normalized, so opposite corners don't have to be checked.
                if (area.FirstAddress.RowNumber < minRow) minRow = area.FirstAddress.RowNumber;
                if (area.LastAddress.RowNumber > maxRow) maxRow = area.LastAddress.RowNumber;
                if (area.FirstAddress.ColumnNumber < minCol) minCol = area.FirstAddress.ColumnNumber;
                if (area.LastAddress.ColumnNumber > maxCol) maxCol = area.LastAddress.ColumnNumber;
            }
        }

        public static Reference UnionOp(Reference lhs, Reference rhs)
        {
            // Build the inline-storage form directly: first area inline, rest in a single array.
            var totalCount = lhs.AreaCount + rhs.AreaCount;
            var additional = new XLRangeAddress[totalCount - 1];
            var index = 0;
            for (var i = 1; i < lhs.AreaCount; i++)
                additional[index++] = lhs[i];
            foreach (var area in rhs)
                additional[index++] = area;

            return new Reference(lhs[0], additional);
        }

        public static OneOf<Reference, XLError> Intersect(Reference lhs, Reference rhs, CalcContext ctx)
        {
            // The two references must share a single worksheet (default = context). Walk both
            // sides with a direct loop instead of Concat/Distinct/ToList.
            XLWorksheet? sheet = null;
            if (!TryResolveSharedSheet(lhs, ctx, ref sheet) || !TryResolveSharedSheet(rhs, ctx, ref sheet))
                return XLError.IncompatibleValue;

            // Sheet is non-null here because both references have at least one area and
            // ctx.Worksheet throws MissingContextException rather than returning null —
            // so each `area.Worksheet ?? ctx.Worksheet` resolves to a non-null value.
            var resolvedSheet = sheet!;
            List<XLRangeAddress>? intersections = null;
            foreach (var leftArea in lhs)
            {
                var intersectedArea = leftArea.WithWorksheet(resolvedSheet);
                foreach (var rightArea in rhs)
                {
                    intersectedArea = intersectedArea.Intersection(rightArea.WithWorksheet(resolvedSheet));
                    if (!intersectedArea.IsValid)
                        break;
                }

                if (intersectedArea.IsValid)
                    (intersections ??= []).Add(intersectedArea);
            }

            return intersections is { Count: > 0 } ? new Reference(intersections) : XLError.NullValue;
        }

        /// <summary>
        /// Resolve the single worksheet shared by all areas of <paramref name="reference"/> (areas
        /// without an explicit worksheet default to the context worksheet), folding the result into
        /// <paramref name="sheet"/>. Returns <c>false</c> if an area belongs to a different worksheet.
        /// </summary>
        private static bool TryResolveSharedSheet(Reference reference, CalcContext ctx, ref XLWorksheet? sheet)
        {
            foreach (var area in reference)
            {
                var ws = area.Worksheet ?? ctx.Worksheet;
                if (sheet is null) sheet = ws;
                else if (sheet != ws) return false;
            }

            return true;
        }

        /// <summary>
        /// Do an implicit intersection of an address.
        /// </summary>
        /// <param name="formulaAddress"></param>
        /// <returns>An address of the intersection or error if the intersection failed.</returns>
        public OneOf<Reference, XLError> ImplicitIntersection(IXLAddress formulaAddress)
        {
            if (AreaCount != 1)
                return XLError.IncompatibleValue;

            var area = _firstArea;
            if (area.RowSpan == 1 && area.ColumnSpan == 1)
                return this;

            var column = formulaAddress.ColumnNumber;
            var row = formulaAddress.RowNumber;

            if (area.ColumnSpan == 1 && area.FirstAddress.RowNumber <= row && row <= area.LastAddress.RowNumber)
            {
                var intersection = new XLAddress(area.Worksheet, row, area.FirstAddress.ColumnNumber, false, false);
                return new Reference(new XLRangeAddress(intersection, intersection));
            }

            if (area.RowSpan == 1 && area.FirstAddress.ColumnNumber <= column && column <= area.LastAddress.ColumnNumber)
            {
                var intersection = new XLAddress(area.Worksheet, area.FirstAddress.RowNumber, column, false, false);
                return new Reference(new XLRangeAddress(intersection, intersection));
            }

            return XLError.IncompatibleValue;
        }

        internal bool IsSingleCell()
        {
            return AreaCount == 1 && _firstArea.IsSingleCell();
        }

        internal bool TryGetSingleCellValue(out ScalarValue value, CalcContext ctx)
        {
            if (!IsSingleCell())
            {
                value = default;
                return false;
            }

            value = ctx.GetCellValue(_firstArea.Worksheet, _firstArea.FirstAddress.RowNumber, _firstArea.FirstAddress.ColumnNumber);
            return true;
        }

        internal OneOf<Array, XLError> ToArray(CalcContext context)
        {
            if (AreaCount != 1)
                return XLError.IncompatibleValue;

            return new ReferenceArray(_firstArea, context);
        }

        public OneOf<Array, XLError> Apply(Func<ScalarValue, ScalarValue> op, CalcContext context)
        {
            if (AreaCount != 1)
                return XLError.IncompatibleValue;

            var area = _firstArea;
            var width = area.ColumnSpan;
            var height = area.RowSpan;
            var startColumn = area.FirstAddress.ColumnNumber;
            var startRow = area.FirstAddress.RowNumber;
            var data = new ScalarValue[height, width];
            for (int y = 0; y < height; ++y)
            {
                for (int x = 0; x < width; ++x)
                {
                    var row = startRow + y;
                    var column = startColumn + x;
                    var cellValue = context.GetCellValue(area.Worksheet, row, column);
                    data[y, x] = op(cellValue);
                }
            }

            return new ConstArray(data);
        }

        /// <summary>
        /// Attempts to find a single non-null worksheet across all areas of <paramref name="reference"/>.
        /// Returns false if more than one distinct worksheet is referenced.
        /// </summary>
        /// <remarks>
        /// Single-area references skip the context-worksheet fallback (matching the original
        /// special case for <c>Areas.Count == 1</c>); multi-area references substitute the
        /// context worksheet for any null entries before deduplicating.
        /// </remarks>
        private static bool TryResolveSingleWorksheet(Reference reference, XLWorksheet contextWorksheet,
            out XLWorksheet? worksheet)
        {
            worksheet = null;
            if (reference.AreaCount == 1)
            {
                worksheet = reference._firstArea.Worksheet;
                return true;
            }

            foreach (var area in reference)
            {
                var ws = area.Worksheet ?? contextWorksheet;
                if (ws is null) continue;
                if (worksheet is null) worksheet = ws;
                else if (worksheet != ws) return false;
            }

            return true;
        }

        /// <summary>
        /// Allocation-free struct enumerator over the areas of a <see cref="Reference"/>.
        /// </summary>
        public struct Enumerator
        {
            private readonly Reference _reference;
            private readonly int _count;
            private int _index;

            internal Enumerator(Reference reference)
            {
                _reference = reference;
                _count = reference.AreaCount;
                _index = -1;
            }

            public XLRangeAddress Current => _reference[_index];

            public bool MoveNext() => ++_index < _count;
        }
    }
}
