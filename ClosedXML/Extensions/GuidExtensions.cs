#nullable disable


using System;

namespace ClosedXML.Extensions;

internal static class GuidExtensions
{
    internal static string WrapInBraces(this Guid guid)
    {
        return string.Concat('{', guid.ToString(), '}');
    }
}
