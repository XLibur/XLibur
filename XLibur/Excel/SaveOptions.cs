using System.IO.Compression;

namespace XLibur.Excel;

public class SaveOptions
{
    public SaveOptions()
    {
        ValidatePackage = false;
    }

    /// <summary>
    /// How hard to compress the parts written to the package. Defaults to
    /// <see cref="CompressionLevel.Optimal"/>; <see cref="CompressionLevel.Fastest"/> trades a
    /// larger file for a quicker save.
    /// </summary>
    /// <remarks>
    /// <b>Only applies to parts this save creates</b>, which in practice means the first save of a
    /// workbook built from scratch. A part's compression is fixed when the zip entry is created, so
    /// re-saving a workbook that was loaded from a file or stream — or one that has been saved once
    /// already, since it adopts its destination as its origin — leaves every existing part at the
    /// level it was first written with. That includes the parts XLibur rewrites itself: a reopened
    /// <c>sheet1.xml</c> keeps its original entry, so new sheet data goes in at the old level.
    /// <para>
    /// Measured on a 50,000 × 3 workbook: a first save honours the setting (10,585 KB at
    /// <see cref="CompressionLevel.NoCompression"/> against 1,369 KB at
    /// <see cref="CompressionLevel.Optimal"/>), while a second save of the same workbook and a save
    /// of one loaded from a stream both produce identical output at either level. Callers who want
    /// a different level for a template-driven export cannot get one this way.
    /// </para>
    /// </remarks>
    public CompressionLevel CompressionLevel { get; set; } = CompressionLevel.Optimal;

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
    /// Password used to encrypt the saved workbook. When set, the file is written with agile
    /// encryption (AES-256-CBC, SHA-512), the profile Excel itself writes.
    /// </summary>
    /// <remarks>
    /// The default of <c>null</c> means different things to the two save methods, because they
    /// mean different things themselves:
    /// <list type="bullet">
    /// <item>
    ///   <c>SaveAs</c> states the encryption of the file it writes, so <c>null</c> is <em>no
    ///   encryption</em>. A plain <c>SaveAs</c> therefore cannot silently produce a file the caller
    ///   could not open without a password they never mentioned, and is how encryption is removed
    ///   from a workbook that was loaded with it.
    /// </item>
    /// <item>
    ///   <c>Save</c> puts a workbook back where it came from as it was, so <c>null</c> is
    ///   <em>unchanged</em>: one opened with <see cref="LoadOptions.Password"/> is written back
    ///   encrypted under that password. Setting a password rotates it, or encrypts an origin that
    ///   was not encrypted before. <c>Save</c> cannot remove encryption.
    /// </item>
    /// </list>
    /// </remarks>
    public string? Password { get; set; }

    public bool ValidatePackage { get; set; }
}
