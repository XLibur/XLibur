namespace XLibur.Operations;

/// <summary>Every operation, in the order a host lists them. A host adds no operation of its own.</summary>
public static class OperationRegistry
{
    public static IReadOnlyList<IOperation> All { get; } =
    [
        new InspectOperation(),
        new ReadRangeOperation(),
        new FindOperation(),
        new CheckFormulaOperation(),
        new RecalcOperation(),
        new ValidateOperation(),
        new DiffOperation(),
        new EditOperation(),
        new ConvertOperation(),
        new RenderReportOperation(),
    ];

    public static IOperation? Find(string name)
        => All.FirstOrDefault(o => string.Equals(o.Name, name.Replace('-', '_'), StringComparison.OrdinalIgnoreCase));
}
