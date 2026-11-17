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

    /// <summary>
    /// The list a file gave in <c>x12ac:list</c>, or null. It is not a setting of its own: it is how
    /// <see cref="MinValue"/> is spelled when an item holds a comma, so it is internal where the settings
    /// are public.
    /// </summary>
    internal QuotedItemList? QuotedItems { get; init; }

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

/// <summary>
/// A list rule's items as Excel 2010 and later spell them in <c>x12ac:list</c>, where an item that holds
/// a comma is in quotes, as in <c>1,"2,3",4</c>. Excel writes this form only for such a list, with
/// <c>formula1</c> as the fallback for older readers, which splits every comma: <c>"1,2,3,4"</c>.
/// </summary>
/// <param name="Items">The text of <c>x12ac:list</c>.</param>
/// <param name="Fallback">The <c>formula1</c> read with it, which a rule's <c>MinValue</c> holds.</param>
internal readonly record struct QuotedItemList(string Items, string Fallback)
{
    /// <summary>
    /// Whether <paramref name="settings"/> still has the list this was read with. A rule whose list was
    /// changed after the load is written as its new <c>formula1</c> alone.
    /// </summary>
    public bool Spells(in DataValidationSettings settings) =>
        settings.AllowedValues == XLAllowedValues.List && settings.MinValue == Fallback;
}
