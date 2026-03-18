using System;
using System.Linq;
using System.Reflection;

namespace XLibur.Attributes;

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public class XLColumnAttribute : Attribute
{
    public string? Header { get; set; }

    public bool Ignore { get; set; }

    public int Order { get; set; }

    private static XLColumnAttribute? GetXLColumnAttribute(MemberInfo mi)
    {
        return !mi.HasAttribute<XLColumnAttribute>() ? null : mi.GetAttributes<XLColumnAttribute>()[0];
    }

    internal static string? GetHeader(MemberInfo mi)
    {
        var attribute = GetXLColumnAttribute(mi);
        if (attribute == null) return null;
        return string.IsNullOrWhiteSpace(attribute.Header) ? null : attribute.Header;
    }

    internal static int GetOrder(MemberInfo mi)
    {
        var attribute = GetXLColumnAttribute(mi);
        return attribute?.Order ?? int.MaxValue;
    }

    internal static bool IgnoreMember(MemberInfo mi)
    {
        var attribute = GetXLColumnAttribute(mi);
        return attribute is { Ignore: true };
    }
}
