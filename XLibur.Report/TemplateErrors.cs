using System.Collections;
using System.Collections.Generic;

namespace XLibur.Report;

/// <summary>
/// The errors collected during one generation run, in the order they were found.
/// </summary>
public sealed class TemplateErrors : IReadOnlyList<TemplateError>
{
    private readonly List<TemplateError> _errors = new();

    /// <inheritdoc />
    public int Count => _errors.Count;

    /// <inheritdoc />
    public TemplateError this[int index] => _errors[index];

    /// <inheritdoc />
    public IEnumerator<TemplateError> GetEnumerator() => _errors.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    internal void Add(TemplateError error) => _errors.Add(error);

    internal void Clear() => _errors.Clear();
}
