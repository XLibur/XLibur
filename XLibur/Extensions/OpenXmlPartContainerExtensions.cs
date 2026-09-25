using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace XLibur.Extensions;

internal static class OpenXmlPartContainerExtensions
{
    public static bool HasPartWithId(this OpenXmlPartContainer container, string relId)
    {
        return container.Parts.Any(p => p.RelationshipId.Equals(relId));
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
