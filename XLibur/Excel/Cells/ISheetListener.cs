namespace XLibur.Excel;

/// <summary>
/// An interface for components reacting on changes in a worksheet.
/// </summary>
/// <remarks>
/// <para>
/// Implementing this and adding a <c>yield return</c> to
/// <see cref="XLWorksheet.GetSheetListeners"/> is the <em>only</em> thing that makes a sheet
/// feature survive a structural edit. Nothing else in the library reaches a listener by name.
/// </para>
/// <para>
/// The order listeners run in is part of the contract and is pinned by
/// <c>SheetListenerOrderTests</c>. A listener that is yielded for every sheet — defined names and
/// data-validation criteria formulas are workbook-scoped — guards on <c>edit.Sheet</c> only if it
/// should not act on an edit elsewhere; see <see cref="XLHyperlinks"/> for the guarded idiom and
/// <c>XLDefinedNames</c> for the deliberately unguarded one.
/// </para>
/// </remarks>
internal interface ISheetListener
{
    /// <summary>
    /// A handler called after an area was put into the sheet and cells shifted down.
    /// </summary>
    /// <param name="edit">The edit. <c>edit.Area</c> has been inserted; the original cells were shifted down.</param>
    void OnInsertAreaAndShiftDown(in SheetEdit edit);

    /// <summary>
    /// A handler called after an area was put into the sheet and cells shifted right.
    /// </summary>
    /// <param name="edit">The edit. <c>edit.Area</c> has been inserted; the original cells were shifted right.</param>
    void OnInsertAreaAndShiftRight(in SheetEdit edit);

    /// <summary>
    /// A handler called after an area was deleted from the sheet and cells shifted left.
    /// </summary>
    /// <param name="edit">The edit. <c>edit.Area</c> has been deleted; cells to the right were shifted left.</param>
    void OnDeleteAreaAndShiftLeft(in SheetEdit edit);

    /// <summary>
    /// A handler called after an area was deleted from the sheet and cells shifted up.
    /// </summary>
    /// <param name="edit">The edit. <c>edit.Area</c> has been deleted; cells below were shifted up.</param>
    void OnDeleteAreaAndShiftUp(in SheetEdit edit);
}
