using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel.Coordinates;

/// <summary>
/// A representation of a <c>ST_Ref</c>, i.e., an area in a sheet (no reference to the sheet).
/// </summary>
internal readonly struct Area : IEquatable<Area>, IEnumerable<Point>
{
    internal Area(Point point)
        : this(point, point)
    {
    }

    internal Area(Point firstPoint, Point lastPoint)
    {
        FirstPoint = firstPoint;
        LastPoint = lastPoint;
    }

    public Area(int rowStart, int columnStart, int rowEnd, int columnEnd)
        : this(new Point(rowStart, columnStart), new Point(rowEnd, columnEnd))
    {
    }

    /// <summary>
    /// A range that covers whole worksheet.
    /// </summary>
    public static readonly Area Full = new(
        new Point(XLHelper.MinRowNumber, XLHelper.MinColumnNumber),
        new Point(XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber));

    /// <summary>
    /// Top-left point of the sheet range.
    /// </summary>
    public readonly Point FirstPoint;

    /// <summary>
    /// Bottom-right point of the sheet range.
    /// </summary>
    public readonly Point LastPoint;

    public int Width => LastPoint.Column - FirstPoint.Column + 1;

    public int Height => LastPoint.Row - FirstPoint.Row + 1;

    /// <summary>
    /// The left column number of the range. From 1 to <see cref="XLHelper.MaxColumnNumber"/>.
    /// </summary>
    public int LeftColumn => FirstPoint.Column;

    /// <summary>
    /// The right column number of the range. From 1 to <see cref="XLHelper.MaxColumnNumber"/>.
    /// Greater or equal to <see cref="LeftColumn"/>.
    /// </summary>
    public int RightColumn => LastPoint.Column;

    /// <summary>
    /// The top row number of the range. From 1 to <see cref="XLHelper.MaxRowNumber"/>.
    /// </summary>
    public int TopRow => FirstPoint.Row;

    /// <summary>
    /// The bottom row number of the range. From 1 to <see cref="XLHelper.MaxRowNumber"/>.
    /// Greater or equal to <see cref="TopRow"/>.
    /// </summary>
    public int BottomRow => LastPoint.Row;

    /// <summary>
    /// Does the range span the full width of a sheet (first column to last column)?
    /// </summary>
    internal bool HasFullRowWidth => LeftColumn == XLHelper.MinColumnNumber && RightColumn == XLHelper.MaxColumnNumber;

    /// <summary>
    /// Does the range span the full height of a sheet (first row to last row)?
    /// </summary>
    internal bool HasFullColumnHeight => TopRow == XLHelper.MinRowNumber && BottomRow == XLHelper.MaxRowNumber;

    public override bool Equals(object? obj)
    {
        return obj is Area range && Equals(range);
    }

    public bool Equals(Area other)
    {
        return FirstPoint.Equals(other.FirstPoint) && LastPoint.Equals(other.LastPoint);
    }

    /// <summary>
    /// Combines the two corners rather than XOR-ing them, because a XOR self-cancels on the
    /// commonest area there is.
    /// </summary>
    /// <remarks>
    /// A single-cell area has <see cref="FirstPoint"/> equal to <see cref="LastPoint"/>, so
    /// <c>first ^ last</c> was <b>zero for every one of them</b> — and a single cell is what most
    /// references in a workbook are. Any <c>Dictionary</c> keyed on <see cref="Area"/> therefore put
    /// every distinct single-cell key in one bucket and degraded to a linear scan.
    /// <para>
    /// That made every consumer keyed on an area quadratic in the number of distinct single-cell
    /// keys — the dependency tree over a workbook's formula precedents, and
    /// <c>XLHyperlinks</c>, whose keys are all single cells. Measurements are in
    /// <c>docs/specs/19-benchmark-hotspot-survey.md</c>, area 5, rather than here.
    /// </para>
    /// <para>
    /// A XOR is also symmetric, so it collapsed each rectangle onto its own reversal. Only the
    /// normalised order is ever constructed, so nothing depended on telling them apart, but
    /// combining fixes that too. <c>AreaHashCodeTests</c> pins both properties as distributions.
    /// </para>
    /// </remarks>
    public override int GetHashCode()
    {
        return HashCode.Combine(FirstPoint, LastPoint);
    }

    public static bool operator ==(Area left, Area right) => left.Equals(right);

    public static bool operator !=(Area left, Area right) => !(left == right);


    /// <inheritdoc cref="Parse(ReadOnlySpan{char})"/>
    public static Area Parse(string input) => Parse(input.AsSpan());

    /// <summary>
    /// Parse point per type <c>ST_Ref</c> from
    /// <a href="https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/e7f22870-88a1-4c06-8e5f-d035b1179c50">2.1.1119 Part 4 Section 3.18.64, ST_Ref (Cell Range Reference)</a>
    /// </summary>
    /// <remarks>Can be one cell reference (A1) or two separated by a colon (A1:B2). First reference is always in top left corner</remarks>
    /// <param name="input">Input text</param>
    /// <exception cref="FormatException">If the input doesn't match expected grammar.</exception>
    public static Area Parse(ReadOnlySpan<char> input)
    {
        if (!TryParse(input, out var area))
            throw new FormatException($"Area reference doesn't have correct format: '{input.ToString()}'.");

        return area;
    }

    /// <summary>
    /// Try to parse area. Doesn't accept any extra whitespace anywhere in the input. Letters
    /// must be upper case. Area can specify one corner (<c>A1</c>) or both corners (<c>A1:B3</c>).
    /// </summary>
    public static bool TryParse(ReadOnlySpan<char> input, out Area area)
    {
        var separatorIndex = input.IndexOf(':');
        if (separatorIndex == -1)
        {
            if (!Point.TryParse(input, out var sheetPoint))
            {
                area = default;
                return false;
            }

            area = new Area(sheetPoint, sheetPoint);
            return true;
        }

        if (!Point.TryParse(input[..separatorIndex], out var first) ||
            !Point.TryParse(input[(separatorIndex + 1)..], out var second) ||
            first.Column > second.Column || first.Row > second.Row)
        {
            area = default;
            return false;
        }

        area = new Area(first, second);
        return true;
    }

    /// <summary>
    /// Write the sheet range to the span. If range has only one cell, write only the cell.
    /// </summary>
    /// <param name="output">Must be at least 21 chars long.</param>
    /// <returns>Number of written characters.</returns>
    public int Format(Span<char> output)
    {
        if (FirstPoint == LastPoint)
            return FirstPoint.Format(output);

        var firstPointLen = FirstPoint.Format(output);
        output[firstPointLen] = ':';
        var lastPointLen = LastPoint.Format(output.Slice(firstPointLen + 1));
        return firstPointLen + 1 + lastPointLen;
    }

    public override string ToString()
    {
        Span<char> text = stackalloc char[21];
        var len = Format(text);
        return text.Slice(0, len).ToString();
    }

    /// <summary>
    /// Return a range that contains all cells below the current range.
    /// </summary>
    /// <exception cref="InvalidOperationException">The range touches the bottom border of the sheet.</exception>
    internal Area BelowRange()
    {
        return BelowRange(XLHelper.MaxRowNumber);
    }

    /// <summary>
    /// Get a range below the current one <paramref name="rows"/> rows.
    /// If there isn't enough rows, use as many as possible.
    /// </summary>
    /// <exception cref="InvalidOperationException">The range touches the bottom border of the sheet.</exception>
    internal Area BelowRange(int rows)
    {
        if (LastPoint.Row >= XLHelper.MaxRowNumber)
            throw new InvalidOperationException("No cells below.");

        rows = Math.Min(rows, XLHelper.MaxRowNumber - LastPoint.Row);
        return new Area(
            new Point(LastPoint.Row + 1, FirstPoint.Column),
            new Point(LastPoint.Row + rows, LastPoint.Column));
    }

    /// <summary>
    /// Return a range that contains all cells to the right of the range.
    /// </summary>
    /// <exception cref="InvalidOperationException">The range touches the right border of the sheet.</exception>
    internal Area RightRange()
    {
        if (LastPoint.Column == XLHelper.MaxColumnNumber)
            throw new InvalidOperationException("No cells to the left.");

        return new Area(
            new Point(FirstPoint.Row, LastPoint.Column + 1),
            new Point(LastPoint.Row, XLHelper.MaxColumnNumber));
    }

    /// <summary>
    /// Return a range that contains additional number of rows below.
    /// </summary>
    internal Area ExtendBelow(int rows)
    {
        Debug.Assert(rows >= 0);
        var row = Math.Min(LastPoint.Row + rows, XLHelper.MaxRowNumber);
        return new Area(FirstPoint, new Point(row, LastPoint.Column));
    }

    /// <summary>
    /// Return a range that contains additional number of columns to the right.
    /// </summary>
    internal Area ExtendRight(int columns)
    {
        Debug.Assert(columns >= 0);
        var column = Math.Min(LastPoint.Column + columns, XLHelper.MaxColumnNumber);
        return new Area(FirstPoint, new Point(LastPoint.Row, column));
    }

    internal static Area FromRangeAddress<T>(T address)
        where T : IXLRangeAddress
    {
        var firstPoint = Point.FromAddress(address.FirstAddress);
        var lastPoint = Point.FromAddress(address.LastAddress);
        if (firstPoint.Row > lastPoint.Row || firstPoint.Column > lastPoint.Column)
            return new Area(lastPoint, firstPoint);

        return new Area(firstPoint, lastPoint);
    }

    public bool Contains(Point point)
    {
        return
            point.Row >= FirstPoint.Row && point.Row <= LastPoint.Row &&
            point.Column >= FirstPoint.Column && point.Column <= LastPoint.Column;
    }

    /// <summary>
    /// Create a new range from this one by taking a number of rows from the bottom row up.
    /// </summary>
    /// <param name="rows">How many rows to take, must be at least one.</param>
    public Area SliceFromBottom(int rows)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(rows, 1);

        return new Area(new Point(BottomRow - rows + 1, FirstPoint.Column), LastPoint);
    }

    /// <summary>
    /// Create a new range from this one by taking a number of rows from the top row down.
    /// </summary>
    /// <param name="rows">How many rows to take, must be at least one.</param>
    public Area SliceFromTop(int rows)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(rows, 1);

        return new Area(FirstPoint, new Point(TopRow + rows - 1, LastPoint.Column));
    }

    /// <summary>
    /// Create a new range from this one by taking a number of rows from the left column to the right.
    /// </summary>
    /// <param name="columns">How many columns to take, must be at least one.</param>
    public Area SliceFromLeft(int columns)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(columns, 1);

        return new Area(FirstPoint, new Point(FirstPoint.Row, LeftColumn + columns - 1));
    }

    /// <summary>
    /// Create a new range from this one by taking a number of rows from the bottom row up.
    /// </summary>
    /// <param name="columns">How many columns to take, must be at least one.</param>
    public Area SliceFromRight(int columns)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(columns, 1);

        return new Area(new Point(FirstPoint.Row, RightColumn - columns + 1), LastPoint);
    }

    /// <summary>
    /// Create a new sheet range that is a result of range operator (<c>:</c>)
    /// of this sheet range and <paramref name="otherRange"/>
    /// </summary>
    /// <param name="otherRange">The other range.</param>
    /// <returns>A range that contains both this range and <paramref name="otherRange"/>.</returns>
    public Area Range(Area otherRange)
    {
        var topRow = Math.Min(TopRow, otherRange.TopRow);
        var leftColumn = Math.Min(LeftColumn, otherRange.LeftColumn);
        var bottomRow = Math.Max(BottomRow, otherRange.BottomRow);
        var rightColumn = Math.Max(RightColumn, otherRange.RightColumn);
        return new Area(topRow, leftColumn, bottomRow, rightColumn);
    }

    /// <summary>
    /// Does this range intersects with <paramref name="other"/>.
    /// </summary>
    /// <returns><c>true</c> if intersects, <c>false</c> otherwise.</returns>
    internal bool Intersects(Area other)
    {
        return Intersect(other) is not null;
    }

    /// <summary>
    /// Do an intersection between this range and other range.
    /// </summary>
    /// <param name="other">Other range.</param>
    /// <returns>The intersection range if it exists and is non-empty or null, if intersection doesn't exist.</returns>
    internal Area? Intersect(Area other)
    {
        var leftColumn = Math.Max(LeftColumn, other.LeftColumn);
        var rightColumn = Math.Min(RightColumn, other.RightColumn);
        var topRow = Math.Max(TopRow, other.TopRow);
        var bottomRow = Math.Min(BottomRow, other.BottomRow);

        if (bottomRow < topRow || rightColumn < leftColumn)
            return null;

        return new Area(topRow, leftColumn, bottomRow, rightColumn);
    }

    /// <summary>
    /// Does this range overlaps the <paramref name="otherRange"/>?
    /// </summary>
    internal bool Overlaps(Area otherRange)
    {
        return TopRow <= otherRange.TopRow &&
               RightColumn >= otherRange.RightColumn &&
               BottomRow >= otherRange.BottomRow &&
               LeftColumn <= otherRange.LeftColumn;
    }

    /// <summary>
    /// Does range cover all rows, from top row to bottom row of a sheet.
    /// </summary>
    internal bool IsEntireColumn()
    {
        return TopRow == 1 && BottomRow == XLHelper.MaxRowNumber;
    }

    /// <summary>
    /// Does range cover all columns, from first to last column of a sheet.
    /// </summary>
    public bool IsEntireRow()
    {
        return LeftColumn == 1 && RightColumn == XLHelper.MaxColumnNumber;
    }

    /// <summary>
    /// Return a new range that has the same size as the current one,
    /// </summary>
    /// <param name="topLeftCorner">New top left coordinate of returned range.</param>
    /// <returns>New range.</returns>
    internal Area At(Point topLeftCorner)
    {
        var bottomRightCorner = topLeftCorner.ShiftColumn(Width - 1).ShiftRow(Height - 1);
        return new Area(topLeftCorner, bottomRightCorner);
    }

    /// <summary>
    /// Return a new range that has been shifted in vertical direction by <paramref name="rowShift"/>.
    /// </summary>
    /// <param name="rowShift">By how much to shift the range, positive - downwards, negative - upwards.</param>
    /// <returns>Newly created area.</returns>
    internal Area ShiftRows(int rowShift)
    {
        var topLeftCorner = FirstPoint.ShiftRow(rowShift);
        var bottomRightCorner = LastPoint.ShiftRow(rowShift);
        return new Area(topLeftCorner, bottomRightCorner);
    }

    /// <summary>
    /// Return a new range that has been shifted in horizontal direction by <paramref name="columnShift"/>.
    /// </summary>
    /// <param name="columnShift">By how much to shift the range, positive - rightward, negative - leftward.</param>
    /// <returns>Newly created area.</returns>
    internal Area ShiftColumns(int columnShift)
    {
        var topLeftCorner = FirstPoint.ShiftColumn(columnShift);
        var bottomRightCorner = LastPoint.ShiftColumn(columnShift);
        return new Area(topLeftCorner, bottomRightCorner);
    }

    public IEnumerator<Point> GetEnumerator()
    {
        for (var row = TopRow; row <= BottomRow; ++row)
        {
            for (var col = LeftColumn; col <= RightColumn; ++col)
            {
                yield return new Point(row, col);
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// Calculate size and position of the area when another area is inserted into a sheet.
    /// </summary>
    /// <param name="insertedArea">Inserted area.</param>
    /// <param name="result">The result, might be <c>null</c> as a valid result if area is pushed out.</param>
    /// <returns><c>true</c> if results wasn't partially shifted.</returns>
    internal bool TryInsertAreaAndShiftRight(Area insertedArea, out Area? result)
    {
        // Inserted fully upward, downward or to the right
        if (insertedArea.BottomRow < TopRow ||
            insertedArea.TopRow > BottomRow ||
            insertedArea.LeftColumn > RightColumn)
        {
            result = this;
            return true;
        }

        var fullyOverlaps = insertedArea.TopRow <= TopRow &&
                            insertedArea.BottomRow >= BottomRow;
        if (!fullyOverlaps)
        {
            result = null;
            return false;
        }

        // Are is effectively inserted into a seam at the left column of the insertedArea
        if (insertedArea.LeftColumn <= LeftColumn)
        {
            // Area is completely pushed out
            if (LeftColumn + insertedArea.Width > XLHelper.MaxColumnNumber)
            {
                result = null;
                return true;
            }

            // Area is partially pushed out
            if (RightColumn + insertedArea.Width > XLHelper.MaxColumnNumber)
            {
                var pushedOutColsCount = RightColumn + insertedArea.Width - XLHelper.MaxColumnNumber;
                var keepCols = Width - pushedOutColsCount;
                var resized = SliceFromLeft(keepCols);
                result = resized.ShiftColumns(insertedArea.Width);
                return true;
            }

            // Not pushed out = only shift
            result = ShiftColumns(insertedArea.Width);
            return true;
        }

        result = ExtendRight(insertedArea.Width);
        return true;
    }

    /// <summary>
    /// Calculate size and position of the area when another area is inserted into a sheet.
    /// </summary>
    /// <param name="insertedArea">Inserted area.</param>
    /// <param name="result">The result, might be <c>null</c> as a valid result if area is pushed out.</param>
    /// <returns><c>true</c> if results wasn't partially shifted.</returns>
    internal bool TryInsertAreaAndShiftDown(Area insertedArea, out Area? result)
    {
        // Inserted fully to the left, to the right or below
        if (insertedArea.RightColumn < LeftColumn ||
            insertedArea.LeftColumn > RightColumn ||
            insertedArea.TopRow > BottomRow)
        {
            result = this;
            return true;
        }

        var fullyOverlaps = insertedArea.LeftColumn <= LeftColumn &&
                            insertedArea.RightColumn >= RightColumn;
        if (!fullyOverlaps)
        {
            result = null;
            return false;
        }

        // Are is effectively inserted into a seam at the top row of the insertedArea
        if (insertedArea.TopRow <= TopRow)
        {
            // Area is completely pushed out
            if (TopRow + insertedArea.Height > XLHelper.MaxRowNumber)
            {
                result = null;
                return true;
            }

            // Area is partially pushed out
            if (BottomRow + insertedArea.Height > XLHelper.MaxRowNumber)
            {
                var pushedOutRowsCount = BottomRow + insertedArea.Height - XLHelper.MaxRowNumber;
                var keepRows = Height - pushedOutRowsCount;
                var resized = SliceFromTop(keepRows);
                result = resized.ShiftRows(insertedArea.Height);
                return true;
            }

            // Not pushed out = only shift
            result = ShiftRows(insertedArea.Height);
            return true;
        }

        result = ExtendBelow(insertedArea.Height);
        return true;
    }

    /// <summary>
    /// Take the area and reposition it as if the <paramref name="deletedArea"/> was removed
    /// from sheet. If cells the left of the area are deleted, the area shifts to the left.
    /// If <paramref name="deletedArea"/> is within the area, the width of the area decreases.
    /// </summary>
    /// <remarks>
    /// If the method returns <c>false</c>, there is a partial cover and it's up to you to
    /// decide what to do.
    /// </remarks>
    /// <returns>
    /// The <paramref name="result"/> has a value <c>null</c> if the range was completely
    /// removed by <paramref name="deletedArea"/>.
    /// </returns>
    internal bool TryDeleteAreaAndShiftLeft(Area deletedArea, out Area? result)
    {
        // Deleted area is fully upwards, downwards or to the right of this area.
        if (deletedArea.BottomRow < TopRow ||
            deletedArea.TopRow > BottomRow ||
            deletedArea.LeftColumn > RightColumn)
        {
            result = this;
            return true;
        }

        var doesntOverlapHeight = deletedArea.TopRow > TopRow ||
                                  deletedArea.BottomRow < BottomRow;
        var deletesColumnsToLeft = deletedArea.LeftColumn < LeftColumn;
        var deletesColumnsOfArea = deletedArea.LeftColumn <= RightColumn &&
                                   deletedArea.RightColumn >= LeftColumn;
        if (doesntOverlapHeight && (deletesColumnsToLeft || deletesColumnsOfArea))
        {
            result = null;
            return false;
        }

        var repositioned = this;
        if (deletesColumnsOfArea)
        {
            // Decrease width of repositioned area
            var left = Math.Max(deletedArea.LeftColumn, repositioned.LeftColumn);
            var right = Math.Min(deletedArea.RightColumn, repositioned.RightColumn);

            var columnsToDelete = right - left + 1;
            var newWidth = repositioned.Width - columnsToDelete;
            if (newWidth == 0)
            {
                result = null;
                return true;
            }

            repositioned = repositioned.SliceFromLeft(newWidth);
        }

        if (deletesColumnsToLeft)
        {
            // There are some deleted columns to the left of the area -> shift left
            var deletedLastColumnsOutwards = Math.Min(repositioned.LeftColumn - 1, deletedArea.RightColumn);

            var shiftLeft = deletedLastColumnsOutwards - deletedArea.LeftColumn + 1;
            repositioned = repositioned.ShiftColumns(-shiftLeft);
        }

        result = repositioned;
        return true;
    }

    /// <summary>
    /// Take the area and reposition it as if the <paramref name="deletedArea"/> was removed
    /// from sheet. If cells upward of the area are deleted, the area shifts to the upward.
    /// If <paramref name="deletedArea"/> is within the area, the height of the area decreases.
    /// </summary>
    /// <remarks>
    /// If the method returns <c>false</c>, there is a partial cover and it's up to you to
    /// decide what to do.
    /// </remarks>
    /// <returns>
    /// The <paramref name="result"/> has a value <c>null</c> if the range was completely
    /// removed by <paramref name="deletedArea"/>.
    /// </returns>
    internal bool TryDeleteAreaAndShiftUp(Area deletedArea, out Area? result)
    {
        // Deleted area is fully on left, right or bottom side of this area.
        if (deletedArea.RightColumn < LeftColumn ||
            deletedArea.LeftColumn > RightColumn ||
            deletedArea.TopRow > BottomRow)
        {
            result = this;
            return true;
        }

        var doesntOverlapWidth = deletedArea.LeftColumn > LeftColumn ||
                                 deletedArea.RightColumn < RightColumn;
        var deletesRowsAboveArea = deletedArea.TopRow < TopRow;
        var deletesRowsOfArea = deletedArea.TopRow <= BottomRow &&
                                deletedArea.BottomRow >= TopRow;
        if (doesntOverlapWidth && (deletesRowsAboveArea || deletesRowsOfArea))
        {
            result = null;
            return false;
        }

        var repositioned = this;
        if (deletesRowsOfArea)
        {
            // Decrease height of repositioned area
            var top = Math.Max(deletedArea.TopRow, repositioned.TopRow);
            var bottom = Math.Min(deletedArea.BottomRow, repositioned.BottomRow);

            var rowsToDelete = bottom - top + 1;
            var newHeight = repositioned.Height - rowsToDelete;
            if (newHeight == 0)
            {
                result = null;
                return true;
            }

            repositioned = repositioned.SliceFromTop(newHeight);
        }

        if (deletesRowsAboveArea)
        {
            // There are some deleted rows above the area -> shift up
            var deletedLastRowAboveArea = Math.Min(repositioned.TopRow - 1, deletedArea.BottomRow);

            var shiftUp = deletedLastRowAboveArea - deletedArea.TopRow + 1;
            repositioned = repositioned.ShiftRows(-shiftUp);
        }

        result = repositioned;
        return true;
    }

    /// <summary>
    /// Shift the range by rows and then columns, clipping any part pushed outside the sheet.
    /// </summary>
    /// <returns>The shifted, clipped range, or <c>null</c> if it was pushed entirely off the sheet.</returns>
    internal Area? ShiftAndClip(int rowShift, int columnShift)
    {
        if (ShiftRowsAndClip(rowShift) is not { } rowShifted)
            return null;

        if (rowShifted.ShiftColumnsAndClip(columnShift) is not { } rowAndColumnShifted)
            return null;

        return rowAndColumnShifted;
    }

    /// <summary>
    /// Shift the range vertically by <paramref name="rowShift"/>. If the shifted range is out of
    /// sheet bounds, clip the part that is out.
    /// </summary>
    /// <returns>Shifted clipped range or <c>null</c> if shifted completely off the sheet.</returns>
    internal Area? ShiftRowsAndClip(int rowShift)
    {
        var shiftedTop = TopRow + rowShift;
        if (shiftedTop > XLHelper.MaxRowNumber)
            return null;

        var shiftedBottom = BottomRow + rowShift;
        if (shiftedBottom < XLHelper.MinRowNumber)
            return null;

        var clippedTop = Math.Max(shiftedTop, XLHelper.MinRowNumber);
        var clippedBottom = Math.Min(shiftedBottom, XLHelper.MaxRowNumber);

        return new Area(clippedTop, LeftColumn, clippedBottom, RightColumn);
    }

    /// <summary>
    /// Shift the range horizontally by <paramref name="columnShift"/>. If the shifted range is out
    /// of sheet bounds, clip the part that is out.
    /// </summary>
    /// <returns>Shifted clipped range or <c>null</c> if shifted completely off the sheet.</returns>
    internal Area? ShiftColumnsAndClip(int columnShift)
    {
        var shiftedLeft = LeftColumn + columnShift;
        if (shiftedLeft > XLHelper.MaxColumnNumber)
            return null;

        var shiftedRight = RightColumn + columnShift;
        if (shiftedRight < XLHelper.MinColumnNumber)
            return null;

        var clippedLeft = Math.Max(shiftedLeft, XLHelper.MinColumnNumber);
        var clippedRight = Math.Min(shiftedRight, XLHelper.MaxColumnNumber);

        return new Area(TopRow, clippedLeft, BottomRow, clippedRight);
    }

    /// <summary>
    /// Remove <paramref name="range"/> from this range, appending the remaining (up to four)
    /// rectangular pieces to <paramref name="nonExcludedAreas"/>.
    /// </summary>
    /// <returns>The intersection that was excluded, or <c>null</c> if the ranges don't intersect.</returns>
    internal Area? Exclude(Area range, List<Area> nonExcludedAreas)
    {
        if (Intersect(range) is not { } intersection)
        {
            nonExcludedAreas.Add(this);
            return null;
        }

        // top
        if (TopRow < intersection.TopRow)
            nonExcludedAreas.Add(new Area(TopRow, LeftColumn, intersection.TopRow - 1, RightColumn));

        // bottom
        if (BottomRow > intersection.BottomRow)
            nonExcludedAreas.Add(new Area(intersection.BottomRow + 1, LeftColumn, BottomRow, RightColumn));

        // left
        if (LeftColumn < intersection.LeftColumn)
            nonExcludedAreas.Add(new Area(intersection.TopRow, LeftColumn, intersection.BottomRow, intersection.LeftColumn - 1));

        // right
        if (RightColumn > intersection.RightColumn)
            nonExcludedAreas.Add(new Area(intersection.TopRow, intersection.RightColumn + 1, intersection.BottomRow, RightColumn));

        return intersection;
    }

    /// <summary>
    /// Reposition the range as if columns were inserted at <paramref name="insertedLeftColumn"/>,
    /// mimicking Excel behavior (shift when the insert is to the left, extend when inside).
    /// </summary>
    internal Area? ShiftOrExtendRight(int insertedLeftColumn, int insertedWidth)
    {
        Debug.Assert(insertedWidth >= 0);

        // Range inserted at the right edge extends - hence the - 1
        if (RightColumn < insertedLeftColumn - 1)
            return this; // inserted is to the right of range -> no shift

        if (LeftColumn >= insertedLeftColumn)
            return ShiftColumnsAndClip(insertedWidth); // inserted is to the left -> shift

        // inserted is in the middle of the range: LeftColumn < insertedLeftColumn <= RightColumn
        return ExtendRight(insertedWidth);
    }

    /// <summary>
    /// Reposition the range as if rows were inserted at <paramref name="insertedTopRow"/>,
    /// mimicking Excel behavior (shift when the insert is above, extend when inside).
    /// </summary>
    internal Area? ShiftOrExtendDown(int insertedTopRow, int insertedHeight)
    {
        Debug.Assert(insertedHeight >= 0);

        // Range inserted at the bottom edge extends - hence the - 1
        if (BottomRow < insertedTopRow - 1)
            return this; // inserted is below the range -> no shift

        if (TopRow >= insertedTopRow)
            return ShiftRowsAndClip(insertedHeight); // inserted is above -> shift

        // inserted is in the middle of the range: TopRow < insertedTopRow <= BottomRow
        return ExtendBelow(insertedHeight);
    }

    /// <summary>
    /// Reposition the range as if rows were deleted from <paramref name="deletedTopRow"/>,
    /// mimicking Excel behavior (shift when the delete is above, shrink when overlapping).
    /// </summary>
    internal Area? ShiftOrShrinkUp(int deletedTopRow, int deletedHeight)
    {
        Debug.Assert(deletedHeight >= 0);
        if (BottomRow < deletedTopRow || deletedHeight == 0)
            return this; // deleted is below the range -> no shift or shrink

        var deletedBottomRow = deletedTopRow + deletedHeight - 1;
        if (deletedBottomRow < TopRow)
            return ShiftRows(-deletedHeight); // deleted completely above -> only shift

        // Shrink by how much the deleted range and this range overlap
        var shrink = Math.Min(BottomRow, deletedBottomRow) - Math.Max(TopRow, deletedTopRow) + 1;
        if (shrink == Height)
            return null;

        var shift = Math.Max(TopRow - deletedTopRow, 0);
        var shifted = ShiftRows(-shift);
        return new Area(shifted.TopRow, shifted.LeftColumn, shifted.BottomRow - shrink, shifted.RightColumn);
    }

    /// <summary>
    /// Reposition the range as if columns were deleted from <paramref name="deletedLeftColumn"/>,
    /// mimicking Excel behavior (shift when the delete is to the left, shrink when overlapping).
    /// </summary>
    internal Area? ShiftOrShrinkLeft(int deletedLeftColumn, int deletedWidth)
    {
        Debug.Assert(deletedWidth >= 0);
        if (RightColumn < deletedLeftColumn || deletedWidth == 0)
            return this; // deleted is to the right of the range -> no shift or shrink

        var deletedRightColumn = deletedLeftColumn + deletedWidth - 1;
        if (deletedRightColumn < LeftColumn)
            return ShiftColumns(-deletedWidth); // deleted completely to the left -> only shift

        // Shrink by how much the deleted range and this range overlap
        var shrink = Math.Min(RightColumn, deletedRightColumn) - Math.Max(LeftColumn, deletedLeftColumn) + 1;
        if (shrink == Width)
            return null;

        var shift = Math.Max(LeftColumn - deletedLeftColumn, 0);
        var shifted = ShiftColumns(-shift);
        return new Area(shifted.TopRow, shifted.LeftColumn, shifted.BottomRow, shifted.RightColumn - shrink);
    }

    /// <summary>
    /// Split the range above <paramref name="row"/>, into <paramref name="above"/> (rows before
    /// <paramref name="row"/>) and <paramref name="below"/> (from <paramref name="row"/> down).
    /// </summary>
    /// <returns><c>true</c> if <paramref name="above"/> is not null.</returns>
    internal bool SplitAbove(int row, [NotNullWhen(true)] out Area? above, out Area? below)
    {
        if (row is < XLHelper.MinRowNumber or > XLHelper.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(row));

        if (BottomRow < row)
        {
            above = this;
            below = null;
            return true;
        }

        if (TopRow >= row)
        {
            above = null;
            below = this;
            return false;
        }

        above = new Area(TopRow, LeftColumn, row - 1, RightColumn);
        below = new Area(row, LeftColumn, BottomRow, RightColumn);
        return true;
    }

    /// <summary>
    /// Split the range below <paramref name="row"/>, into <paramref name="below"/> (rows after
    /// <paramref name="row"/>) and <paramref name="above"/> (up to and including <paramref name="row"/>).
    /// </summary>
    /// <returns><c>true</c> if <paramref name="below"/> is not null.</returns>
    internal bool SplitBelow(int row, [NotNullWhen(true)] out Area? below, out Area? above)
    {
        if (row is < XLHelper.MinRowNumber or > XLHelper.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(row));

        if (TopRow > row)
        {
            below = this;
            above = null;
            return true;
        }

        if (BottomRow <= row)
        {
            below = null;
            above = this;
            return false;
        }

        below = new Area(row + 1, LeftColumn, BottomRow, RightColumn);
        above = new Area(TopRow, LeftColumn, row, RightColumn);
        return true;
    }

    /// <summary>
    /// Split the range before <paramref name="column"/>, into <paramref name="left"/> (columns
    /// before <paramref name="column"/>) and <paramref name="right"/> (from <paramref name="column"/> right).
    /// </summary>
    /// <returns><c>true</c> if <paramref name="left"/> is not null.</returns>
    internal bool SplitBefore(int column, [NotNullWhen(true)] out Area? left, out Area? right)
    {
        if (column is < XLHelper.MinColumnNumber or > XLHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(column));

        if (RightColumn < column)
        {
            left = this;
            right = null;
            return true;
        }

        if (LeftColumn >= column)
        {
            left = null;
            right = this;
            return false;
        }

        left = new Area(TopRow, LeftColumn, BottomRow, column - 1);
        right = new Area(TopRow, column, BottomRow, RightColumn);
        return true;
    }

    /// <summary>
    /// Split the range after <paramref name="column"/>, into <paramref name="right"/> (columns
    /// after <paramref name="column"/>) and <paramref name="left"/> (up to and including <paramref name="column"/>).
    /// </summary>
    /// <returns><c>true</c> if <paramref name="right"/> is not null.</returns>
    internal bool SplitAfter(int column, [NotNullWhen(true)] out Area? right, out Area? left)
    {
        if (column is < XLHelper.MinColumnNumber or > XLHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(column));

        if (LeftColumn > column)
        {
            right = this;
            left = null;
            return true;
        }

        if (RightColumn <= column)
        {
            right = null;
            left = this;
            return false;
        }

        right = new Area(TopRow, column + 1, BottomRow, RightColumn);
        left = new Area(TopRow, LeftColumn, BottomRow, column);
        return true;
    }

    /// <summary>
    /// Wrap this single range in a one-element <see cref="XLAreaList"/>.
    /// </summary>
    internal XLAreaList ToAreaList()
    {
        return new XLAreaList(this);
    }
}
