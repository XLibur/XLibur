namespace XLibur.Excel;

/// <summary>
/// Listener for components that need to be notified about structural changes of a workbook
/// (removing or renaming a sheet). See <see cref="ISheetListener"/> for similar listener about
/// structural changes of a sheet.
/// </summary>
/// <remarks>
/// <see cref="XLWorksheets"/> is the one door a sheet is renamed or deleted through. It raises each
/// event on every listener <see cref="XLWorksheets.GetWorkbookListeners"/> yields.
/// <para>
/// <b>A listener must not throw.</b> The door does not catch, and there is no prepare step to roll
/// back, so a listener that throws leaves the workbook part-way through the rename or the delete:
/// some holders changed and others not. Input a listener cannot work with, such as a formula the
/// parser refuses, is left as it is (ADR 0002) rather than reported by throwing.
/// </para>
/// </remarks>
internal interface IWorkbookListener
{
    /// <summary>
    /// Method is called when sheet has already been renamed. Each component is responsible only
    /// for changing data in itself, not other components. The goal is to separate concerns so
    /// each component is not too dependent on others and can achieve the goal in efficient manner.
    /// </summary>
    /// <param name="oldSheetName">Old sheet name.</param>
    /// <param name="newSheetName">New sheet name, different from old one.</param>
    void OnSheetRenamed(string oldSheetName, string newSheetName);

    /// <summary>
    /// Raised before the sheet is removed, while it can still be resolved. Each component changes
    /// only its own data, as for <see cref="OnSheetRenamed"/>.
    /// </summary>
    /// <remarks>
    /// Must not throw: a failure here would leave the workbook part-way through the delete.
    /// </remarks>
    /// <param name="sheetName">Name of the sheet about to be deleted.</param>
    void OnSheetDeleting(string sheetName);
}
