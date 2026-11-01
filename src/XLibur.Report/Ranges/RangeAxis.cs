using System.Globalization;
using XLibur.Excel;

namespace XLibur.Report.Ranges;

/// <summary>
/// Which way a bound range repeats, and every sheet operation that depends on the answer.
/// </summary>
/// <remarks>
/// <para>
/// A range repeating downwards and one repeating across do the same things to different axes:
/// insert slots, copy a block into them, put the options slot last, total along a line. Expressed
/// against this abstraction the expander is written once and reads the same either way, which is
/// worth more than the indirection costs — the alternative is an <c>isHorizontal</c> flag tested in
/// a dozen places, where every place is a chance to test it wrongly.
/// </para>
/// <para>
/// The vocabulary throughout: a <strong>slot</strong> is what repeats — a row when the range runs
/// downwards, a column when it runs across. A <strong>line</strong> is the other axis, the one a slot
/// cuts across. So a vertical range's summary tag sits in a line (a column) and totals down the
/// slots; a horizontal range's sits in a line (a row) and totals across them.
/// </para>
/// </remarks>
internal abstract class RangeAxis
{
    /// <summary>The range repeats downwards: one row, or block of rows, per item.</summary>
    public static readonly RangeAxis Vertical = new VerticalAxis();

    /// <summary>The range repeats across: one column, or block of columns, per item.</summary>
    public static readonly RangeAxis Horizontal = new HorizontalAxis();

    /// <summary>Whether the range repeats across rather than downwards.</summary>
    public abstract bool IsHorizontal { get; }

    /// <summary>What the axis calls a slot, for an error message a template author has to act on.</summary>
    public abstract string SlotName { get; }

    /// <summary>The first slot of <paramref name="area"/>.</summary>
    public abstract int FirstSlot(RangeArea area);

    /// <summary>The last slot of <paramref name="area"/> — its options slot, when it has one.</summary>
    public abstract int LastSlot(RangeArea area);

    /// <summary>How many slots <paramref name="area"/> spans.</summary>
    public abstract int SlotCount(RangeArea area);

    /// <summary>The first line of <paramref name="area"/>.</summary>
    public abstract int FirstLine(RangeArea area);

    /// <summary>The last line of <paramref name="area"/>.</summary>
    public abstract int LastLine(RangeArea area);

    /// <summary>The slot a cell sits in.</summary>
    public abstract int SlotOf(IXLAddress address);

    /// <summary>The line a cell sits in.</summary>
    public abstract int LineOf(IXLAddress address);

    /// <summary>The cell at <paramref name="slot"/> and <paramref name="line"/>.</summary>
    public abstract IXLCell Cell(IXLWorksheet sheet, int slot, int line);

    /// <summary>
    /// The rectangle covering slots <paramref name="firstSlot"/> to <paramref name="lastSlot"/>
    /// across the whole of <paramref name="area"/>'s lines.
    /// </summary>
    public abstract RangeArea Slots(RangeArea area, int firstSlot, int lastSlot);

    /// <summary>Inserts <paramref name="count"/> slots immediately after <paramref name="slot"/>.</summary>
    public abstract void InsertSlotsAfter(IXLWorksheet sheet, int slot, int count);

    /// <summary>Deletes slots <paramref name="firstSlot"/> to <paramref name="lastSlot"/>.</summary>
    public abstract void DeleteSlots(IXLWorksheet sheet, int firstSlot, int lastSlot);

    /// <summary>Hides the line <paramref name="line"/>.</summary>
    public abstract void HideLine(IXLWorksheet sheet, int line);

    /// <summary>Deletes the line <paramref name="line"/>.</summary>
    public abstract void DeleteLine(IXLWorksheet sheet, int line);

    /// <summary>
    /// An A1 reference to the run of cells in <paramref name="line"/> from
    /// <paramref name="firstSlot"/> to <paramref name="lastSlot"/> — what a summary formula totals.
    /// </summary>
    public abstract string LineReference(int line, int firstSlot, int lastSlot);

    /// <summary>
    /// Reads what a template author wrote for a line: a column letter or number for a vertical
    /// range, a row number for a horizontal one. Returns <c>null</c> if it names nothing.
    /// </summary>
    public abstract int? ParseLine(string text);

    /// <summary>The sheet range covering <paramref name="area"/>.</summary>
    public static IXLRange Range(IXLWorksheet sheet, RangeArea area) =>
        sheet.Range(area.FirstRow, area.FirstColumn, area.LastRow, area.LastColumn);

    private sealed class VerticalAxis : RangeAxis
    {
        public override bool IsHorizontal => false;

        public override string SlotName => "row";

        public override int FirstSlot(RangeArea area) => area.FirstRow;

        public override int LastSlot(RangeArea area) => area.LastRow;

        public override int SlotCount(RangeArea area) => area.RowCount;

        public override int FirstLine(RangeArea area) => area.FirstColumn;

        public override int LastLine(RangeArea area) => area.LastColumn;

        public override int SlotOf(IXLAddress address) => address.RowNumber;

        public override int LineOf(IXLAddress address) => address.ColumnNumber;

        public override IXLCell Cell(IXLWorksheet sheet, int slot, int line) => sheet.Cell(slot, line);

        public override RangeArea Slots(RangeArea area, int firstSlot, int lastSlot) =>
            new(firstSlot, area.FirstColumn, lastSlot, area.LastColumn);

        public override void InsertSlotsAfter(IXLWorksheet sheet, int slot, int count) =>
            sheet.Row(slot).InsertRowsBelow(count);

        public override void DeleteSlots(IXLWorksheet sheet, int firstSlot, int lastSlot) =>
            sheet.Rows(firstSlot, lastSlot).Delete();

        public override void HideLine(IXLWorksheet sheet, int line) => sheet.Column(line).Hide();

        public override void DeleteLine(IXLWorksheet sheet, int line) => sheet.Column(line).Delete();

        public override string LineReference(int line, int firstSlot, int lastSlot)
        {
            var letter = XLHelper.GetColumnLetterFromNumber(line);
            return string.Create(CultureInfo.InvariantCulture, $"{letter}{firstSlot}:{letter}{lastSlot}");
        }

        public override int? ParseLine(string text)
        {
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number))
            {
                return number;
            }

            try
            {
                return XLHelper.GetColumnNumberFromLetter(text);
            }
            catch (System.ArgumentException)
            {
                return null;
            }
        }
    }

    private sealed class HorizontalAxis : RangeAxis
    {
        public override bool IsHorizontal => true;

        public override string SlotName => "column";

        public override int FirstSlot(RangeArea area) => area.FirstColumn;

        public override int LastSlot(RangeArea area) => area.LastColumn;

        public override int SlotCount(RangeArea area) => area.ColumnCount;

        public override int FirstLine(RangeArea area) => area.FirstRow;

        public override int LastLine(RangeArea area) => area.LastRow;

        public override int SlotOf(IXLAddress address) => address.ColumnNumber;

        public override int LineOf(IXLAddress address) => address.RowNumber;

        public override IXLCell Cell(IXLWorksheet sheet, int slot, int line) => sheet.Cell(line, slot);

        public override RangeArea Slots(RangeArea area, int firstSlot, int lastSlot) =>
            new(area.FirstRow, firstSlot, area.LastRow, lastSlot);

        public override void InsertSlotsAfter(IXLWorksheet sheet, int slot, int count) =>
            sheet.Column(slot).InsertColumnsAfter(count);

        public override void DeleteSlots(IXLWorksheet sheet, int firstSlot, int lastSlot) =>
            sheet.Columns(firstSlot, lastSlot).Delete();

        public override void HideLine(IXLWorksheet sheet, int line) => sheet.Row(line).Hide();

        public override void DeleteLine(IXLWorksheet sheet, int line) => sheet.Row(line).Delete();

        public override string LineReference(int line, int firstSlot, int lastSlot)
        {
            var first = XLHelper.GetColumnLetterFromNumber(firstSlot);
            var last = XLHelper.GetColumnLetterFromNumber(lastSlot);
            return string.Create(CultureInfo.InvariantCulture, $"{first}{line}:{last}{line}");
        }

        public override int? ParseLine(string text) =>
            int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) ? number : null;
    }
}
