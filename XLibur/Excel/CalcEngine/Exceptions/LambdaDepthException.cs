using System;

namespace XLibur.Excel.CalcEngine.Exceptions;

/// <summary>
/// Calls of a <c>LAMBDA</c>, one inside another, went deeper than the engine allows, as a function that
/// calls itself for ever does. The formula of the cell is <c>#NUM!</c>, as in Excel, where
/// <c>IFERROR</c> around the call does not catch it either, so it ends the whole formula rather than
/// being an error value.
/// </summary>
internal sealed class LambdaDepthException() : InvalidOperationException("LAMBDA calls went too deep.");
