using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Packaging;

namespace XLibur.Extensions;

internal static class OpenXmlPartContainerExtensions
{
    /// <summary>
    /// Whether the container has a child part with relationship id <paramref name="relId"/>.
    /// A null or empty id has no part.
    /// </summary>
    /// <remarks>
    /// One dictionary lookup through <see cref="OpenXmlPartContainer.TryGetPartById"/>, where the
    /// old check scanned every child relationship. The guard keeps the old answer of false for a
    /// null or empty id, which <c>TryGetPartById</c> is not documented to accept.
    /// </remarks>
    public static bool HasPartWithId(this OpenXmlPartContainer container, [NotNullWhen(true)] string? relId)
    {
        return !string.IsNullOrEmpty(relId) && container.TryGetPartById(relId, out _);
    }

    /// <summary>
    /// The child part with relationship id <paramref name="relId"/>, or null when there is no such
    /// relationship or the part is not a <typeparamref name="T"/>.
    /// </summary>
    /// <remarks>
    /// <see cref="OpenXmlPartContainer.GetPartById"/> throws for an unknown id, and a relationship
    /// can be missing when the model object was created in a workbook that has not been saved
    /// through this package yet. This is one dictionary lookup, where the check it replaces
    /// scanned every child relationship first.
    /// </remarks>
    public static T? GetPartOrNull<T>(this OpenXmlPartContainer container, string? relId)
        where T : OpenXmlPart
    {
        if (string.IsNullOrEmpty(relId))
            return null;

        return container.TryGetPartById(relId, out var part) ? part as T : null;
    }
}
