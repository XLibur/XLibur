using System;

namespace XLibur.Excel.CalcEngine.Exceptions;

/// <summary>
/// The formula is a valid Excel construct that XLibur does not evaluate: an unsupported feature.
/// </summary>
/// <remarks>
/// <para>
/// It derives from <see cref="NotImplementedException"/>, the type an unsupported feature has always
/// reached a caller as, so a caller's <c>catch (NotImplementedException)</c> keeps working and the
/// public edge needs no translation.
/// </para>
/// <para>
/// The distinct type is what lets <see cref="EvaluationFailure"/> tell an unsupported feature from a
/// <see cref="NotImplementedException"/> thrown anywhere else in the library, which is a defect.
/// Every gap the calc engine knows it has must be raised as this type: save writes no cached value
/// for an unsupported feature, but throws on a defect (ADR 0001).
/// </para>
/// </remarks>
internal sealed class UnsupportedFeatureException(string message) : NotImplementedException(message);
