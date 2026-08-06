using System;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;

namespace XLibur.Excel.ContentManagers;

/// <summary>
/// Tracks, for one OpenXML container, which element currently occupies each schema-ordered slot,
/// so a writer inserting a new element can find the element it must follow.
/// </summary>
/// <remarks>
/// Backed by a dense array indexed by the enum value rather than a dictionary. The slots are a
/// small contiguous range known at construction, so the array costs one allocation, needs no
/// hashing, and lets <see cref="GetPreviousElementFor"/> walk backwards without allocating —
/// it previously ran a LINQ <c>Where</c>/<c>DefaultIfEmpty</c>/<c>MaxBy</c> chain over the whole
/// dictionary on every call, and writers call it once per element they emit.
/// </remarks>
internal abstract class XLBaseContentManager<T>
    where T : struct, Enum
{
    /// <summary>
    /// Enforces the precondition of <see cref="Index"/> once per closed generic type, in every
    /// build configuration. A <c>Debug.Assert</c> would not do: it compiles out of Release,
    /// so declaring a future slot enum with a non-<see cref="int"/> underlying type would silently
    /// index the wrong slot — or past the end of the array — in exactly the builds that ship.
    /// </summary>
    static XLBaseContentManager()
    {
        var underlying = Enum.GetUnderlyingType(typeof(T));
        if (underlying != typeof(int))
        {
            throw new InvalidOperationException(
                $"{typeof(T).Name} is backed by {underlying.Name}, but {nameof(XLBaseContentManager<T>)} " +
                "reinterprets its slot enum as int. Declare the enum with the default int underlying type.");
        }
    }

    private readonly OpenXmlElement?[] _contents;

    /// <param name="highestSlot">
    /// The largest value of <typeparamref name="T"/>. Slots are addressed by their numeric value,
    /// so gaps in the enum simply stay null.
    /// </param>
    protected XLBaseContentManager(T highestSlot)
    {
        _contents = new OpenXmlElement?[Index(highestSlot) + 1];
    }

    /// <summary>
    /// The element occupying the highest-numbered slot below <paramref name="content"/>, or null
    /// when nothing precedes it.
    /// </summary>
    public OpenXmlElement? GetPreviousElementFor(T content)
    {
        for (var i = Index(content) - 1; i >= 0; i--)
        {
            if (_contents[i] is { } element)
                return element;
        }

        return null;
    }

    public void SetElement(T content, OpenXmlElement? element) => _contents[Index(content)] = element;

    /// <summary>
    /// Reinterprets the enum as its underlying <see cref="int"/>. Casting through
    /// <see cref="ValueType"/> would box on every call, which matters because writers call
    /// <see cref="GetPreviousElementFor"/> once per emitted element. The static constructor
    /// guarantees the reinterpret is valid.
    /// </summary>
    private static int Index(T content) => Unsafe.As<T, int>(ref content);
}
