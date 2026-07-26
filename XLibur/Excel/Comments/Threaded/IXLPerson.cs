using System;

namespace XLibur.Excel;

/// <summary>
/// An identity that can author a <see cref="IXLThreadedComment"/>. Persons are stored once per
/// workbook in <c>xl/persons/person.xml</c> and referenced by threaded comments through
/// <see cref="Id"/>.
/// </summary>
public interface IXLPerson
{
    /// <summary>
    /// The name Excel displays next to the person's comments.
    /// </summary>
    string DisplayName { get; }

    /// <summary>
    /// The identifier of the person within <see cref="ProviderId"/>, e.g. a Windows SID for the
    /// <c>AD</c> provider. Null for persons created through the API without an identity provider.
    /// </summary>
    string? UserId { get; }

    /// <summary>
    /// The identity provider that issued <see cref="UserId"/>, e.g. <c>AD</c>, <c>PeoplePicker</c>
    /// or <c>Windows Live</c>. Null for persons created through the API without an identity provider.
    /// </summary>
    string? ProviderId { get; }

    /// <summary>
    /// The workbook-unique identifier of the person. Preserved across a load/save round trip so
    /// that threaded comments keep pointing at the same person.
    /// </summary>
    Guid Id { get; }
}
