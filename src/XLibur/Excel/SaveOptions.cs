using System.IO.Compression;

namespace XLibur.Excel;

public class SaveOptions
{
    public SaveOptions()
    {
        ValidatePackage = false;
    }

    /// <summary>
    /// How hard to compress the parts the save writes. Defaults to
    /// <see cref="CompressionLevel.Optimal"/>; <see cref="CompressionLevel.Fastest"/> trades a
    /// larger file for a quicker save.
    /// </summary>
    /// <remarks>
    /// The level applies to every part the save writes, whether the workbook is new or was loaded
    /// from a file or stream. A save of a loaded workbook leaves the parts it does not write as the
    /// file had them, compressed as they were: a sheet whose part it keeps whole (see
    /// <see cref="RewriteUnchangedSheets"/>), the theme, a VBA project, custom XML, and charts it
    /// did not change. The package relationships (<c>_rels/.rels</c>) and the core properties are
    /// always stored uncompressed, as every save of a new workbook stores them. Set
    /// <see cref="RewriteUnchangedSheets"/> to <c>true</c> to have every sheet written at this level.
    /// <para>
    /// Use <see cref="CompressionLevel.Fastest"/> when save time matters more than file size. On a
    /// workbook of 8,000 numbers, loading it, changing one cell and saving it takes about a quarter
    /// less time than at <see cref="CompressionLevel.Optimal"/>; a part it writes comes out 65–90%
    /// larger in XLibur's measurements. <see cref="CompressionLevel.SmallestSize"/> rarely pays: it
    /// takes about twice the time for a few percent less, and some content comes out larger than at
    /// <see cref="CompressionLevel.Optimal"/>.
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
    /// Write the cells of every sheet from the workbook's model, including sheets whose cells have
    /// not changed since the workbook was loaded. Defaults to <c>false</c>.
    /// </summary>
    /// <remarks>
    /// By default, the first save of a loaded workbook copies the cells of an unchanged sheet as the
    /// file had them, instead of writing them again. That is faster, and keeps markup XLibur would
    /// otherwise drop or respell, such as row <c>spans</c> and the exact text of numbers. The rest of
    /// the sheet is still written from the model. When that would say the same as the file did, the
    /// whole sheet is kept as the file had it, down to the bytes. Set this to <c>true</c> to have every
    /// sheet written as XLibur writes it, for example to normalise a file.
    /// <para>
    /// A sheet's cells are copied only when nothing a save writes into them has changed, including
    /// the text and cached value of each formula, every cell in a table's totals row holds its field's
    /// label or a formula, and every shared string the file had is written back at its index. A sheet
    /// with formulas is written from the model when <see cref="WriteSharedFormulas"/> is
    /// off, since its kept cells would keep the file's shared formulas.
    /// </para>
    /// </remarks>
    public bool RewriteUnchangedSheets { get; set; }

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

    /// <summary>
    /// Write back the shared formulas that a load brought in, limited to the cells that still hold
    /// them. Defaults to <c>true</c>.
    /// </summary>
    /// <remarks>
    /// Excel saves a formula that was filled or copied as a shared formula: the first cell holds the
    /// formula text, and each later cell holds only a reference to it. By default a save keeps that
    /// form for every cell whose formula has not changed or moved since the load. The file is about
    /// 10% smaller, and Excel opens it faster. XLibur loads it 6–15% more slowly, but a load followed
    /// by a recalculation allocates less and takes the same time or less, because each cell of a shared
    /// formula is evaluated through one parsed formula. Formulas that a caller writes are never saved
    /// as shared.
    /// <para>
    /// Set this to <c>false</c> to write the formula text in every cell, as XLibur did before. Do
    /// that for a reader that reads the text of each cell's formula and does not expand shared
    /// formulas.
    /// </para>
    /// </remarks>
    public bool WriteSharedFormulas { get; set; } = true;
}
