using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;

namespace XLibur.Operations.Workbooks;

/// <summary>Why reading a formula cell produced no value, in the words of CONTEXT.md.</summary>
public enum CellFailure
{
    Circular,
    Refused,
    Unsupported,
    Defect,
}

public readonly record struct CellRead(XLCellValue Value, CellFailure? Failure, string? Message)
{
    public string ToText() => Failure is { } failure ? "!" + Rendering.Name(failure) : Rendering.ToText(Value);
}

internal static class CellReader
{
    /// <summary>
    /// Reads a cell's value, which calculates a dirty formula. XLibur leaves the cells of a cycle, a
    /// refused formula and an unsupported feature dirty, and reading one throws, so each failure is
    /// classified here instead of ending the operation. Anything else thrown is a defect, recorded for
    /// that cell so one bug does not hide the rest of the report.
    /// </summary>
    internal static CellRead Read(IXLCell cell) => Read(cell, static c => c.Value);

    internal static CellRead Read<TState>(TState state, Func<TState, XLCellValue> evaluate)
    {
        try
        {
            return new CellRead(evaluate(state), null, null);
        }
        catch (XLCircularReferenceException e)
        {
            return new CellRead(Blank.Value, CellFailure.Circular, e.Message);
        }
        catch (ExpressionParseException e)
        {
            return new CellRead(Blank.Value, CellFailure.Refused, e.Message);
        }
        catch (Exception e) when (e is NotImplementedException or NotSupportedException or XLNoWorksheetContextException)
        {
            return new CellRead(Blank.Value, CellFailure.Unsupported, e.Message);
        }
        catch (Exception e)
        {
            return new CellRead(Blank.Value, CellFailure.Defect, $"{e.GetType().FullName}: {e.Message}");
        }
    }
}
