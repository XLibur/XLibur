using System;
using System.Text;

namespace XLibur.Excel.IO;

/// <summary>
/// Finds the byte range of the <c>&lt;sheetData&gt;</c> element in a UTF-8 worksheet part without
/// tokenising it.
/// </summary>
/// <remarks>
/// The rule this relies on is lexical and exact for the inputs it accepts. In XML, a literal
/// <c>&lt;</c> can appear only as the start of markup: never in text content and never in an
/// attribute value. Only comments, CDATA sections and processing instructions can carry one as
/// data, so a byte search that finds none of those in the region it crosses can only have found
/// real tags. Any input outside that shape — another encoding, a comment ahead of the cells, a
/// CDATA formula — is refused rather than handled approximately, and the caller reads the part the
/// slow way instead.
/// <para>
/// The loader also confirms the result against an <see cref="System.Xml.XmlReader"/> before using
/// it: the first element named <c>sheetData</c> has to be in the main namespace, a child of the
/// root and written with the prefix found here.
/// </para>
/// </remarks>
internal static class SheetDataLocator
{
    private const int MaxQualifiedNameLength = 64;

    /// <summary>Where <c>&lt;sheetData&gt;</c> sits in the part.</summary>
    /// <param name="Start">Index of the <c>&lt;</c> opening the start tag.</param>
    /// <param name="End">Index just past the <c>&gt;</c> closing the element.</param>
    /// <param name="NameStart">Index of the element's qualified name inside the start tag.</param>
    /// <param name="NameLength">Byte length of the qualified name, prefix included.</param>
    /// <param name="Prefix">The namespace prefix, or an empty string for none.</param>
    internal readonly record struct Location(int Start, int End, int NameStart, int NameLength, string Prefix);

    /// <summary>
    /// Locates <c>&lt;sheetData&gt;</c>, or returns false when the part is not in a shape this can
    /// read exactly.
    /// </summary>
    internal static bool TryLocate(ReadOnlySpan<byte> xml, out Location location)
    {
        location = default;

        var bodyStart = SkipPreamble(xml);
        if (bodyStart < 0)
            return false;

        if (!TryFindStartTag(xml, bodyStart, out var start, out var nameStart, out var nameLength))
            return false;

        // Everything ahead of the element has to be plain elements, or the match above could sit
        // inside a comment or a CDATA section.
        if (ContainsNonElementMarkup(xml[bodyStart..start]))
            return false;

        var name = xml.Slice(nameStart, nameLength);
        var tagEnd = FindTagEnd(xml, nameStart + nameLength);
        if (tagEnd < 0)
            return false;

        var colon = name.IndexOf((byte)':');
        var prefix = colon < 0 ? string.Empty : Encoding.UTF8.GetString(name[..colon]);

        if (xml[tagEnd - 1] == (byte)'/')
        {
            location = new Location(start, tagEnd + 1, nameStart, nameLength, prefix);
            return true;
        }

        if (!TryFindEndTag(xml, tagEnd + 1, name, out var end))
            return false;

        location = new Location(start, end, nameStart, nameLength, prefix);
        return true;
    }

    /// <summary>
    /// Returns the index the markup starts at, past a UTF-8 byte order mark and the XML declaration,
    /// or -1 for anything other than UTF-8.
    /// </summary>
    private static int SkipPreamble(ReadOnlySpan<byte> xml)
    {
        if (xml.Length < 4)
            return -1;

        // UTF-16 and UTF-32, with or without a byte order mark: in either, one of the first four
        // bytes is 0x00 or the mark itself.
        if (xml[0] is 0xFE or 0xFF || xml[..4].Contains((byte)0))
            return -1;

        var pos = xml.StartsWith("﻿"u8) ? 3 : 0;
        var rest = xml[pos..];
        if (!rest.StartsWith("<?xml"u8) || rest.Length < 6 || !IsWhitespace(rest[5]))
            return pos;

        var declarationLength = rest.IndexOf("?>"u8);
        if (declarationLength < 0)
            return -1;

        return IsUtf8Declaration(rest[..declarationLength]) ? pos + declarationLength + 2 : -1;
    }

    private static bool IsUtf8Declaration(ReadOnlySpan<byte> declaration)
    {
        var at = declaration.IndexOf("encoding"u8);
        if (at < 0)
            return true; // No encoding declared means UTF-8.

        var rest = declaration[(at + "encoding".Length)..].TrimStart(" \t\r\n"u8);
        if (rest.IsEmpty || rest[0] != (byte)'=')
            return false;

        rest = rest[1..].TrimStart(" \t\r\n"u8);
        if (rest.IsEmpty || rest[0] is not ((byte)'"' or (byte)'\''))
            return false;

        var quote = rest[0];
        rest = rest[1..];
        var close = rest.IndexOf(quote);
        return close >= 0 && Ascii.EqualsIgnoreCase(rest[..close], "utf-8"u8);
    }

    /// <summary>
    /// Finds the first start tag whose local name is <c>sheetData</c>, with or without a prefix.
    /// </summary>
    private static bool TryFindStartTag(ReadOnlySpan<byte> xml, int from, out int start, out int nameStart,
        out int nameLength)
    {
        ReadOnlySpan<byte> localName = "sheetData"u8;
        start = nameStart = nameLength = 0;

        while (true)
        {
            var found = xml[from..].IndexOf(localName);
            if (found < 0)
                return false;

            var local = from + found;
            var afterName = local + localName.Length;
            from = local + 1;

            if (afterName >= xml.Length || !IsNameTerminator(xml[afterName]))
                continue;

            if (local > 0 && xml[local - 1] == (byte)'<')
            {
                start = local - 1;
                nameStart = local;
            }
            else if (local > 1 && xml[local - 1] == (byte)':')
            {
                var i = local - 2;
                while (i >= 0 && IsNameByte(xml[i]))
                    i--;

                // A prefix of at least one character, opened by '<'. An attribute value such as
                // codeName="x:sheetData" is preceded by a quote, not by '<', and so falls through.
                if (i < 0 || i == local - 2 || xml[i] != (byte)'<')
                    continue;

                start = i;
                nameStart = i + 1;
            }
            else
            {
                continue;
            }

            nameLength = afterName - nameStart;
            return nameLength <= MaxQualifiedNameLength;
        }
    }

    /// <summary>
    /// Finds the end of <c>&lt;/name&gt;</c> for the element whose content starts at
    /// <paramref name="from"/>, refusing a region that holds anything but plain tags and text.
    /// </summary>
    private static bool TryFindEndTag(ReadOnlySpan<byte> xml, int from, ReadOnlySpan<byte> name, out int end)
    {
        end = 0;
        Span<byte> endTag = stackalloc byte[name.Length + 2];
        endTag[0] = (byte)'<';
        endTag[1] = (byte)'/';
        name.CopyTo(endTag[2..]);

        var searchFrom = from;
        while (true)
        {
            var found = xml[searchFrom..].IndexOf(endTag);
            if (found < 0)
                return false;

            var close = searchFrom + found;
            var afterName = close + endTag.Length;
            var gt = afterName;
            while (gt < xml.Length && IsWhitespace(xml[gt]))
                gt++;

            if (gt < xml.Length && xml[gt] == (byte)'>')
            {
                var content = xml[from..close];

                // A comment, CDATA section or processing instruction could hide an end tag; a
                // nested element of the same name would make the first end tag the wrong one.
                // Neither occurs in a file Excel writes, so both are refused, not handled.
                if (ContainsNonElementMarkup(content) || ContainsStartTag(content, name))
                    return false;

                end = gt + 1;
                return true;
            }

            searchFrom = close + 1;
        }
    }

    private static bool ContainsStartTag(ReadOnlySpan<byte> content, ReadOnlySpan<byte> name)
    {
        var from = 0;
        while (true)
        {
            var found = content[from..].IndexOf(name);
            if (found < 0)
                return false;

            var at = from + found;
            if (at > 0 && content[at - 1] == (byte)'<')
                return true;

            from = at + 1;
        }
    }

    /// <summary>
    /// Returns the index of the <c>&gt;</c> that closes the tag, skipping any inside a quoted
    /// attribute value, where the character is legal.
    /// </summary>
    private static int FindTagEnd(ReadOnlySpan<byte> xml, int from)
    {
        byte quote = 0;
        for (var i = from; i < xml.Length; i++)
        {
            var b = xml[i];
            if (quote != 0)
            {
                if (b == quote)
                    quote = 0;
            }
            else if (b is (byte)'"' or (byte)'\'')
            {
                quote = b;
            }
            else if (b == (byte)'>')
            {
                return i;
            }
            else if (b == (byte)'<')
            {
                return -1;
            }
        }

        return -1;
    }

    private static bool ContainsNonElementMarkup(ReadOnlySpan<byte> region) =>
        region.IndexOf("<!"u8) >= 0 || region.IndexOf("<?"u8) >= 0;

    private static bool IsNameTerminator(byte b) => IsWhitespace(b) || b is (byte)'>' or (byte)'/';

    private static bool IsWhitespace(byte b) => b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n';

    /// <summary>
    /// A byte that can belong to an XML name. Every byte of a multi-byte UTF-8 sequence is at or
    /// above 0x80, so a non-ASCII name character is accepted without decoding it.
    /// </summary>
    private static bool IsNameByte(byte b) =>
        b >= 0x80 || char.IsAsciiLetterOrDigit((char)b) || b is (byte)'-' or (byte)'.' or (byte)'_';
}
