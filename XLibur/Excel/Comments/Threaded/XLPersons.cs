using System;
using System.Collections;
using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLPersons : IXLPersons
{
    /// <summary>
    /// Insertion ordered, because <c>person.xml</c> is written in this order and keeping it stable
    /// keeps a load/save round trip free of gratuitous diffs.
    /// </summary>
    private readonly Dictionary<Guid, XLPerson> _persons = new();

    private readonly List<Guid> _order = new();

    public int Count => _order.Count;

    public IXLPerson Add(string displayName) => Add(displayName, userId: null, providerId: null);

    public IXLPerson Add(string displayName, string? userId, string? providerId)
    {
        ArgumentNullException.ThrowIfNull(displayName);

        return AddCore(Guid.NewGuid(), displayName, userId, providerId);
    }

    public IXLPerson? Get(Guid id) => _persons.TryGetValue(id, out var person) ? person : null;

    public IXLPerson? GetByDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName);

        foreach (var id in _order)
        {
            var person = _persons[id];
            if (string.Equals(person.DisplayName, displayName, StringComparison.Ordinal))
                return person;
        }

        return null;
    }

    public bool Remove(Guid id)
    {
        if (!_persons.Remove(id))
            return false;

        _order.Remove(id);
        return true;
    }

    public IEnumerator<IXLPerson> GetEnumerator()
    {
        foreach (var id in _order)
            yield return _persons[id];
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>
    /// Adds a person with an id read from a file. Existing ids win, so that a person referenced by
    /// several sheets is added once and threaded comments keep resolving to the same instance.
    /// </summary>
    internal XLPerson AddOrGet(Guid id, string displayName, string? userId, string? providerId)
    {
        return _persons.TryGetValue(id, out var existing)
            ? existing
            : AddCore(id, displayName, userId, providerId);
    }

    /// <summary>
    /// Returns a person with the same display name and provider identity, adding one when the
    /// workbook has none. Used when a threaded comment is copied into another workbook, where the
    /// source person's id may already be taken by a different identity.
    /// </summary>
    internal XLPerson Map(IXLPerson source)
    {
        if (_persons.TryGetValue(source.Id, out var byId) && IsSameIdentity(byId, source))
            return byId;

        // Only match on identity when there is an identity provider to match on. Two persons with no
        // userId and no providerId who happen to share a display name are indistinguishable in the
        // file yet may well be different people, and merging them would reattribute one person's
        // comments to the other.
        if (HasProviderIdentity(source))
        {
            foreach (var id in _order)
            {
                var candidate = _persons[id];
                if (IsSameIdentity(candidate, source))
                    return candidate;
            }
        }

        var newId = _persons.ContainsKey(source.Id) ? Guid.NewGuid() : source.Id;
        return AddCore(newId, source.DisplayName, source.UserId, source.ProviderId);
    }

    private static bool HasProviderIdentity(IXLPerson person)
    {
        return !string.IsNullOrEmpty(person.UserId) || !string.IsNullOrEmpty(person.ProviderId);
    }

    private static bool IsSameIdentity(XLPerson left, IXLPerson right)
    {
        return string.Equals(left.DisplayName, right.DisplayName, StringComparison.Ordinal)
               && string.Equals(left.UserId, right.UserId, StringComparison.Ordinal)
               && string.Equals(left.ProviderId, right.ProviderId, StringComparison.Ordinal);
    }

    private XLPerson AddCore(Guid id, string displayName, string? userId, string? providerId)
    {
        var person = new XLPerson(id, displayName, userId, providerId);
        _persons.Add(id, person);
        _order.Add(id);
        return person;
    }
}
