using System;

namespace XLibur.Excel.CalcEngine.Exceptions;

/// <summary>
/// An LU decomposition found a column with nothing to pivot on: the matrix is singular.
/// </summary>
/// <remarks>
/// It derives from <see cref="InvalidOperationException"/>, which the regression functions already
/// catch for this case. The distinct type is what lets <c>MDETERM</c> and <c>MINVERSE</c> answer a
/// singular matrix without also swallowing every other invalid operation.
/// </remarks>
internal sealed class SingularMatrixException() : InvalidOperationException("The matrix is singular!");
