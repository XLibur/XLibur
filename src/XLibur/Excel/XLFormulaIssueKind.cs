namespace XLibur.Excel;

/// <summary>
/// Why a formula could not be checked or evaluated. These are the failures a caller can expect from
/// the content of a workbook, as opposed to a defect in XLibur, which is always thrown.
/// </summary>
public enum XLFormulaIssueKind
{
    /// <summary>
    /// The formula text was refused: the parser cannot read it, or it has a part the calc engine
    /// cannot take, such as a function with the wrong number of arguments.
    /// </summary>
    RefusedFormula,

    /// <summary>The formula depends on its own value, directly or through other formulas.</summary>
    CircularReference,

    /// <summary>The formula uses a valid Excel construct that XLibur does not evaluate.</summary>
    UnsupportedFeature,
}
