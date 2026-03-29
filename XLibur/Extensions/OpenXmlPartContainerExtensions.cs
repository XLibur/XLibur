using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace XLibur.Extensions;

internal static class OpenXmlPartContainerExtensions
{
    public static bool HasPartWithId(this OpenXmlPartContainer container, string relId)
    {
        return container.Parts.Any(p => p.RelationshipId.Equals(relId));
    }
}
