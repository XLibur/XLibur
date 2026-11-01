using System;
using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// The workbook-level set of identities that author threaded comments.
/// </summary>
public interface IXLPersons : IEnumerable<IXLPerson>
{
    /// <summary>
    /// The number of persons in the workbook.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Adds a person with a freshly generated <see cref="IXLPerson.Id"/> and no identity provider.
    /// </summary>
    /// <param name="displayName">The name Excel shows next to the person's comments.</param>
    IXLPerson Add(string displayName);

    /// <summary>
    /// Adds a person with a freshly generated <see cref="IXLPerson.Id"/>.
    /// </summary>
    /// <param name="displayName">The name Excel shows next to the person's comments.</param>
    /// <param name="userId">The identifier of the person within <paramref name="providerId"/>.</param>
    /// <param name="providerId">The identity provider, e.g. <c>AD</c> or <c>PeoplePicker</c>.</param>
    IXLPerson Add(string displayName, string? userId, string? providerId);

    /// <summary>
    /// Returns the person with the specified <paramref name="id"/>, or null when the workbook has none.
    /// </summary>
    IXLPerson? Get(Guid id);

    /// <summary>
    /// Returns the first person whose <see cref="IXLPerson.DisplayName"/> matches
    /// <paramref name="displayName"/> ordinally, or null when the workbook has none.
    /// </summary>
    IXLPerson? GetByDisplayName(string displayName);

    /// <summary>
    /// Removes a person. Threaded comments already authored by the person keep referencing it, so
    /// the person is written back out on save; this only removes it from the workbook-level list.
    /// </summary>
    /// <returns><c>true</c> when the person was present.</returns>
    bool Remove(Guid id);
}
