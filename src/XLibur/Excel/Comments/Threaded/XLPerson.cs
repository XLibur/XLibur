using System;

namespace XLibur.Excel;

internal sealed class XLPerson : IXLPerson
{
    internal XLPerson(Guid id, string displayName, string? userId, string? providerId)
    {
        Id = id;
        DisplayName = displayName;
        UserId = userId;
        ProviderId = providerId;
    }

    public string DisplayName { get; }

    public string? UserId { get; }

    public string? ProviderId { get; }

    public Guid Id { get; }

    public override string ToString() => DisplayName;
}
