namespace XLibur.Excel;

public class SaveOptions
{
    public SaveOptions()
    {
        ValidatePackage = false;
    }

    public bool ConsolidateConditionalFormatRanges { get; set; } = true;

    public bool ConsolidateDataValidationRanges { get; set; } = true;

    /// <summary>
    /// Evaluate a cell with a formula and save the calculated value along with the formula.
    /// <list type="bullet">
    /// <item>
    ///   True - formulas are evaluated and the calculated values are saved to the file.
    ///   If evaluation of a formula throws an exception, the value is not saved, but a file is still saved.
    /// </item>
    /// <item>
    ///   False (default) - formulas are not evaluated, and the formula cells don't have their values saved to the file.
    /// </item>
    /// </list>
    /// </summary>
    public bool EvaluateFormulasBeforeSaving { get; set; }

    /// <summary>
    /// Gets or sets the filter privacy flag. Set to null to leave the current property in a saved workbook unchanged
    /// </summary>
    public bool? FilterPrivacy { get; set; }

    public bool GenerateCalculationChain { get; set; } = true;

    public bool ValidatePackage { get; set; }
}
