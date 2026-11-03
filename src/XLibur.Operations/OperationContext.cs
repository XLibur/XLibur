using XLibur.Excel;
using XLibur.Operations.Paths;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations;

public sealed class OperationContext(PathPolicy paths, Limits? limits = null)
{
    public PathPolicy Paths { get; } = paths;

    public Limits Limits { get; } = limits ?? Limits.Default;

    internal (string FullPath, byte[] Bytes) ReadFile(string path)
    {
        using var file = Paths.OpenRead(path);
        return (file.FullPath, file.ReadAllBytes(Limits.MaxInputBytes));
    }

    /// <summary>
    /// Loads a workbook from the bytes read through the policy's open handle, so the file parsed is
    /// the file checked, and no handle to it stays open while the operation runs.
    /// </summary>
    internal LoadedWorkbook LoadWorkbook(string path, bool recalculate = false)
    {
        var (fullPath, bytes) = ReadFile(path);
        if (!IsWorkbookExtension(fullPath))
            throw OperationException.BadInput("not_a_workbook", $"'{path}' is not an .xlsx or .xlsm file.");
        return new LoadedWorkbook(fullPath, WorkbookLoader.Load(bytes, path, Limits, recalculate));
    }

    /// <summary>
    /// Dirty formulas are calculated so the file carries their values. Per ADR 0001, one that fails
    /// with a cycle, an unsupported feature or a refused formula is saved with no cached value, and a
    /// defect throws.
    /// </summary>
    internal void SaveWorkbook(XLWorkbook workbook, WriteTarget target)
        => Paths.WriteAtomically(target,
            stream => workbook.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true }));

    internal static bool IsWorkbookExtension(string path)
        => Path.GetExtension(path).ToLowerInvariant() is ".xlsx" or ".xlsm";
}

internal sealed class LoadedWorkbook(string fullPath, XLWorkbook workbook) : IDisposable
{
    public string FullPath { get; } = fullPath;

    public XLWorkbook Workbook { get; } = workbook;

    public void Dispose() => Workbook.Dispose();
}
