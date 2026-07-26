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

    /// <summary>
    /// Password used to encrypt the saved workbook. Default value is <c>null</c>, which saves an
    /// ordinary unencrypted file. When set, the file is written with agile encryption
    /// (AES-256-CBC, SHA-512), the profile Excel itself writes.
    /// </summary>
    /// <remarks>
    /// A workbook opened with <see cref="LoadOptions.Password"/> is <em>not</em> re-encrypted
    /// automatically on save; the password has to be given here again. Carrying it over implicitly
    /// would mean a plain <c>SaveAs</c> silently produced a file the caller could not open without
    /// a password they never mentioned.
    /// </remarks>
    public string? Password { get; set; }

    public bool ValidatePackage { get; set; }
}
