namespace XLibur.Fuzz;

/// <summary>
/// Turns a fuzzer's raw byte string into a sequence of bounded choices.
///
/// Structure-aware targets need to ask questions like "how many sheets?" or "which of these
/// five shapes of relationship id?", and the fuzzer only offers bytes. This type is the single
/// place that conversion happens, so a target reads as a description of the thing being built
/// rather than as arithmetic on an index.
///
/// Running out of input is not an error: every accessor returns a defined value once the buffer
/// is exhausted, so a short input produces a small, valid artefact rather than an exception. That
/// matters because the fuzzer will supply an empty input on the very first run.
/// </summary>
internal sealed class FuzzBytes
{
    private readonly byte[] _data;
    private int _position;

    public FuzzBytes(ReadOnlySpan<byte> data)
    {
        _data = data.ToArray();
    }

    /// <summary>True once every byte has been consumed. Accessors keep working after this point.</summary>
    public bool Exhausted => _position >= _data.Length;

    /// <summary>Take one byte, or zero if the input is spent.</summary>
    public byte Byte()
    {
        return _position < _data.Length ? _data[_position++] : (byte)0;
    }

    /// <summary>Take an integer in <paramref name="min"/>..<paramref name="max"/> inclusive.</summary>
    public int Int(int min, int max)
    {
        if (max <= min)
            return min;

        var span = max - min + 1;

        // Two bytes so ranges wider than 256 are reachable; the fuzzer mutates each independently.
        var value = Byte() | (Byte() << 8);
        return min + (value % span);
    }

    /// <summary>Take a boolean. Biased so that "the ordinary shape" is not rare.</summary>
    public bool Bool()
    {
        return (Byte() & 1) == 1;
    }

    /// <summary>Choose one of <paramref name="options"/>.</summary>
    public T Pick<T>(params T[] options)
    {
        return options[Int(0, options.Length - 1)];
    }

    /// <summary>
    /// Take a short string built from a character set wide enough to reach XML escaping,
    /// name-validation and culture-sensitive comparison paths, but never invalid XML — the
    /// generated package must always parse, or the target degenerates into the blind one.
    /// </summary>
    public string Text(int maxLength)
    {
        const string alphabet = "abcXYZ019 _-'&<>\"éя漢";

        var length = Int(0, maxLength);
        if (length == 0)
            return string.Empty;

        var builder = new System.Text.StringBuilder(length);
        for (var i = 0; i < length; i++)
            builder.Append(alphabet[Int(0, alphabet.Length - 1)]);

        return builder.ToString();
    }
}
