using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel.Coordinates;

/// <summary>
/// Consolidates an <see cref="XLAreaList"/> into an equivalent list of non-overlapping areas,
/// merging overlapping and adjacent rectangles into maximal blocks (rows first, then columns).
/// Operates purely on <see cref="XLSheetRange"/> value structs via a sparse bit matrix keyed by
/// the boundary rows of the input areas. Separate from the <see cref="XLibur.Excel.Ranges"/>
/// <c>XLRangeConsolidationEngine</c>, which works on live <see cref="IXLRanges"/>.
/// </summary>
internal static class XLAreaConsolidator
{
    internal static XLAreaList Consolidate(XLAreaList areas)
    {
        if (areas.Count == 0)
            return areas;

        var matrix = new XLAreaConsolidationMatrix(areas);
        var consolidated = matrix.GetConsolidatedRanges().ToList();
        return new XLAreaList(consolidated);
    }

    /// <summary>
    /// Represents the covered cells as a set of bit rows (one per boundary row of the input),
    /// then reads maximal rectangles back out.
    /// </summary>
    private sealed class XLAreaConsolidationMatrix
    {
        private readonly Dictionary<int, BitArray> _bitMatrix;
        private readonly int _minColumn;

        internal XLAreaConsolidationMatrix(XLAreaList areas)
        {
            (_bitMatrix, _minColumn) = PrepareBitMatrix(areas);
            FillBitMatrix(areas);
        }

        public IEnumerable<XLSheetRange> GetConsolidatedRanges()
        {
            var rowNumbers = _bitMatrix.Keys.OrderBy(k => k).ToArray();
            for (var i = 0; i < rowNumbers.Length; i++)
            {
                var startRow = rowNumbers[i];
                var startings = GetRangesBoundariesStartingByRow(_bitMatrix[startRow]);

                foreach (var starting in startings)
                {
                    var j = i + 1;
                    while (j < rowNumbers.Length && RowIncludesRange(_bitMatrix[rowNumbers[j]], starting)) j++;

                    var endRow = rowNumbers[j - 1];
                    var startColumn = starting.Item1 + _minColumn - 1;
                    var endColumn = starting.Item2 + _minColumn - 1;

                    yield return new XLSheetRange(startRow, startColumn, endRow, endColumn);

                    while (j > i)
                    {
                        ClearRangeInRow(_bitMatrix[rowNumbers[j - 1]], starting);
                        j--;
                    }
                }
            }
        }

        private void AddToBitMatrix(XLSheetRange area)
        {
            var rows = _bitMatrix.Keys
                .Where(k => k >= area.TopRow && k <= area.BottomRow);

            var minIndex = area.LeftColumn - _minColumn + 1;
            var maxIndex = area.RightColumn - _minColumn + 1;

            foreach (var rowNum in rows)
            {
                for (var i = minIndex; i <= maxIndex; i++)
                {
                    _bitMatrix[rowNum][i] = true;
                }
            }
        }

        private static void ClearRangeInRow(BitArray rowArray, Tuple<int, int> rangeBoundaries)
        {
            for (var i = rangeBoundaries.Item1; i <= rangeBoundaries.Item2; i++)
            {
                rowArray[i] = false;
            }
        }

        private void FillBitMatrix(IEnumerable<XLSheetRange> areas)
        {
            foreach (var area in areas)
            {
                AddToBitMatrix(area);
            }
        }

        private static IEnumerable<Tuple<int, int>> GetRangesBoundariesStartingByRow(BitArray rowArray)
        {
            var startIdx = 0;
            for (var i = 1; i < rowArray.Length - 1; i++)
            {
                if (!rowArray[i - 1] && rowArray[i])
                    startIdx = i;
                if (rowArray[i] && !rowArray[i + 1])
                    yield return new Tuple<int, int>(startIdx, i);
            }
        }

        private static (Dictionary<int, BitArray> BitMatrix, int MinColumn) PrepareBitMatrix(XLAreaList areas)
        {
            var minColumn = XLHelper.MaxColumnNumber + 1;
            var maxColumn = 0;
            foreach (var area in areas)
            {
                minColumn = minColumn <= area.LeftColumn ? minColumn : area.LeftColumn;
                maxColumn = maxColumn >= area.RightColumn ? maxColumn : area.RightColumn;
            }

            // Two guard columns (indices 0 and Length-1) stay clear so boundary detection is uniform.
            var bitMaskSize = maxColumn - minColumn + 3;
            var bitMatrix = new Dictionary<int, BitArray>();
            foreach (var area in areas)
            {
                AddRowBitmask(bitMatrix, area.TopRow, bitMaskSize);
                AddRowBitmask(bitMatrix, area.BottomRow, bitMaskSize);
                AddRowBitmask(bitMatrix, area.BottomRow + 1, bitMaskSize);
            }

            return (bitMatrix, minColumn);

            static void AddRowBitmask(Dictionary<int, BitArray> bitMatrix, int rowNum, int bitMaskSize)
            {
                if (!bitMatrix.ContainsKey(rowNum))
                    bitMatrix.Add(rowNum, new BitArray(bitMaskSize, false));
            }
        }

        private static bool RowIncludesRange(BitArray rowArray, Tuple<int, int> rangeBoundaries)
        {
            for (var i = rangeBoundaries.Item1; i <= rangeBoundaries.Item2; i++)
            {
                if (!rowArray[i])
                    return false;
            }

            return true;
        }
    }
}
