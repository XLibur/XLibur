using System;

namespace XLibur.Excel.CalcEngine.Exceptions;

/// <summary>
/// A formula depends, directly or through other cells, on its own value.
/// </summary>
/// <remarks>
/// It derives from <see cref="InvalidOperationException"/>, the type a cycle has always been
/// reported as, so a caller catching that is unaffected. The distinct type is what lets
/// <see cref="EvaluationFailure"/> tell a cycle - a property of the workbook - from an internal
/// invariant failing with the same base type.
/// </remarks>
internal sealed class CircularReferenceException(string message) : InvalidOperationException(message);
