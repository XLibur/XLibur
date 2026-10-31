namespace XLibur.Excel;

/// <summary>
/// What a data validation rule is: its thirteen settings, apart from the cells it covers.
/// </summary>
/// <remarks>
/// <para>
/// This is the one list of the settings. A rule holds its settings as one of these; copying a rule
/// copies the value; the consolidation equality test is the value's equality; resetting a rule is
/// <see cref="Default"/>; and both file spellings (see <c>DataValidationMapping</c>) read into and
/// write from it. Adding a fourteenth setting is a component here, its public property on
/// <see cref="IXLDataValidation"/>, and its attribute in the two adapters.
/// </para>
/// <para>
/// A component read from a file that leaves it out holds the value the schema gives it, so the three
/// flags the schema defaults to false are false in a value read from a file, and true in <see cref="Default"/>.
/// </para>
/// </remarks>
internal readonly record struct DataValidationSettings(
    XLAllowedValues AllowedValues,
    XLOperator Operator,
    string MinValue,
    string MaxValue,
    bool IgnoreBlanks,
    bool InCellDropdown,
    bool ShowInputMessage,
    bool ShowErrorMessage,
    string InputTitle,
    string InputMessage,
    string ErrorTitle,
    string ErrorMessage,
    XLErrorStyle ErrorStyle)
{
    /// <summary>The settings of a rule that has just been created.</summary>
    public static readonly DataValidationSettings Default = new(
        AllowedValues: XLAllowedValues.AnyValue,
        Operator: XLOperator.Between,
        MinValue: string.Empty,
        MaxValue: string.Empty,
        IgnoreBlanks: true,
        InCellDropdown: true,
        ShowInputMessage: true,
        ShowErrorMessage: true,
        InputTitle: string.Empty,
        InputMessage: string.Empty,
        ErrorTitle: string.Empty,
        ErrorMessage: string.Empty,
        ErrorStyle: XLErrorStyle.Stop);

    /// <summary>The settings of <paramref name="dataValidation"/>, which can be any implementation.</summary>
    public static DataValidationSettings From(IXLDataValidation dataValidation) =>
        dataValidation is XLDataValidation xl
            ? xl.Settings
            : new DataValidationSettings(
                dataValidation.AllowedValues,
                dataValidation.Operator,
                dataValidation.MinValue ?? string.Empty,
                dataValidation.MaxValue ?? string.Empty,
                dataValidation.IgnoreBlanks,
                dataValidation.InCellDropdown,
                dataValidation.ShowInputMessage,
                dataValidation.ShowErrorMessage,
                dataValidation.InputTitle ?? string.Empty,
                dataValidation.InputMessage ?? string.Empty,
                dataValidation.ErrorTitle ?? string.Empty,
                dataValidation.ErrorMessage ?? string.Empty,
                dataValidation.ErrorStyle);

    /// <summary>
    /// Only the types that compare values use the operator. List, custom and any value do not.
    /// </summary>
    public bool UsesOperator() => AllowedValues is
        XLAllowedValues.WholeNumber or
        XLAllowedValues.Decimal or
        XLAllowedValues.Date or
        XLAllowedValues.Time or
        XLAllowedValues.TextLength;

    /// <summary>
    /// Does the rule restrict the values, or show a message the user can see? A rule that does neither
    /// has not been set up.
    /// </summary>
    public bool IsDirty() =>
        AllowedValues != XLAllowedValues.AnyValue
        || (ShowInputMessage && (!string.IsNullOrWhiteSpace(InputTitle) || !string.IsNullOrWhiteSpace(InputMessage)))
        || (ShowErrorMessage && (!string.IsNullOrWhiteSpace(ErrorTitle) || !string.IsNullOrWhiteSpace(ErrorMessage)));

    /// <summary>
    /// Does the rule say anything? <see cref="IsDirty"/>, or a title or message, shown or not.
    /// </summary>
    /// <remarks>
    /// Excel writes the text of an input or error message whose display is turned off, and leaves
    /// out <c>showInputMessage</c> or <c>showErrorMessage</c>. Such a rule loads with the message
    /// hidden, and <see cref="IsDirty"/> does not count hidden text, so an "Any value" rule that has
    /// only that text would be dropped (#709).
    /// </remarks>
    public bool HasContent() =>
        IsDirty()
        || !string.IsNullOrWhiteSpace(InputTitle) || !string.IsNullOrWhiteSpace(InputMessage)
        || !string.IsNullOrWhiteSpace(ErrorTitle) || !string.IsNullOrWhiteSpace(ErrorMessage);
}
