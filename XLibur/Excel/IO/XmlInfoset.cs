using System;
using System.IO;
using System.Text;
using System.Xml;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

/// <summary>
/// Compares two XML documents by what they say rather than how they spell it (#702, tier 1).
/// </summary>
/// <remarks>
/// Two documents are equivalent when they hold the same elements, in the same order, with the same
/// attributes and text. Names are compared by namespace and local name, so prefixes do not matter.
/// Ignored: the XML declaration and byte order mark, namespace declarations, attribute order,
/// comments, processing instructions, text that is only whitespace, and whether an element with no
/// content is written as <c>&lt;a/&gt;</c> or <c>&lt;a&gt;&lt;/a&gt;</c>.
/// <para>
/// Attribute values are compared as they are, which errs towards reporting a difference. The
/// markup-compatibility attributes are the exception that needs more: <c>mc:Ignorable</c> and its
/// kin list prefixes, so the same value can name other namespaces in each document. Each prefix
/// they list must resolve to the same namespace in both.
/// </para>
/// </remarks>
internal static class XmlInfoset
{
    /// <summary>
    /// Do <paramref name="left"/> and <paramref name="right"/> hold the same document? Both are
    /// read from their current position and left open.
    /// </summary>
    internal static bool AreEquivalent(Stream left, Stream right)
    {
        using var leftReader = PartXmlReader.Create(left);
        using var rightReader = PartXmlReader.Create(right);
        var l = new Tokens(leftReader);
        var r = new Tokens(rightReader);

        while (true)
        {
            var leftKind = l.Next();
            var rightKind = r.Next();
            if (leftKind != rightKind)
                return false;

            switch (leftKind)
            {
                case TokenKind.End:
                    return true;
                case TokenKind.StartElement:
                    if (!ElementsMatch(leftReader, rightReader))
                        return false;
                    break;
                case TokenKind.Text:
                    if (l.Text != r.Text)
                        return false;
                    break;
            }
        }
    }

    private static bool ElementsMatch(XmlReader left, XmlReader right)
    {
        if (left.LocalName != right.LocalName || left.NamespaceURI != right.NamespaceURI)
            return false;

        if (CountAttributes(left) != CountAttributes(right))
            return false;

        if (!left.MoveToFirstAttribute())
            return true;

        try
        {
            do
            {
                if (IsNamespaceDeclaration(left))
                    continue;

                var value = right.GetAttribute(left.LocalName, left.NamespaceURI);
                if (value != left.Value)
                    return false;

                if (ListsPrefixes(left) && !PrefixesMatch(left, right, value))
                    return false;
            } while (left.MoveToNextAttribute());

            return true;
        }
        finally
        {
            left.MoveToElement();
        }
    }

    /// <summary>
    /// The attributes of the element the reader is on, not counting namespace declarations.
    /// </summary>
    private static int CountAttributes(XmlReader reader)
    {
        var count = 0;
        if (!reader.MoveToFirstAttribute())
            return count;

        do
        {
            if (!IsNamespaceDeclaration(reader))
                count++;
        } while (reader.MoveToNextAttribute());

        reader.MoveToElement();
        return count;
    }

    private static bool IsNamespaceDeclaration(XmlReader attribute) => attribute.NamespaceURI == XmlnsNs;

    /// <summary>
    /// Is the attribute the reader is on one whose value lists prefixes? The qualified attributes of
    /// markup compatibility, and <c>Requires</c> on an <c>mc:Choice</c>.
    /// </summary>
    private static bool ListsPrefixes(XmlReader attribute) =>
        attribute.NamespaceURI == MarkupCompatibilityNs
        || (attribute.NamespaceURI.Length == 0 && attribute.LocalName == "Requires");

    /// <summary>
    /// Does each prefix listed in <paramref name="value"/> name the same namespace for both readers?
    /// Called on an attribute of the left reader; the right reader is on the matching element. A
    /// token such as <c>w:*</c> or <c>w:p</c> names the prefix before its colon.
    /// </summary>
    private static bool PrefixesMatch(XmlReader left, XmlReader right, string value)
    {
        foreach (var token in value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
        {
            var colon = token.IndexOf(':');
            var prefix = colon < 0 ? token : token[..colon];
            if (left.LookupNamespace(prefix) != right.LookupNamespace(prefix))
                return false;
        }

        return true;
    }

    private const string XmlnsNs = "http://www.w3.org/2000/xmlns/";

    private enum TokenKind
    {
        End,
        StartElement,
        EndElement,
        Text,
    }

    /// <summary>
    /// Reads a document as a sequence of start tags, end tags and text. An empty element gives an
    /// end tag of its own, and adjacent text and CDATA are joined into one text.
    /// </summary>
    private sealed class Tokens(XmlReader reader)
    {
        /// <summary>
        /// Must the reader move on before the next token? False at a node not yet returned: the one
        /// that ended a text.
        /// </summary>
        private bool _advance = true;

        private bool _pendingEnd;
        private readonly StringBuilder _text = new();

        internal string Text { get; private set; } = string.Empty;

        internal TokenKind Next()
        {
            if (_pendingEnd)
            {
                _pendingEnd = false;
                return TokenKind.EndElement;
            }

            if (_advance && !reader.Read())
                return TokenKind.End;

            _advance = true;
            while (!reader.EOF)
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        _pendingEnd = reader.IsEmptyElement;
                        return TokenKind.StartElement;
                    case XmlNodeType.EndElement:
                        return TokenKind.EndElement;
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.SignificantWhitespace:
                        ReadText();
                        _advance = false;
                        return TokenKind.Text;
                    default:
                        if (!reader.Read())
                            return TokenKind.End;
                        break;
                }
            }

            return TokenKind.End;
        }

        /// <summary>
        /// Joins the text from here to the next node that is not text, and stops on that node.
        /// </summary>
        private void ReadText()
        {
            _text.Clear();
            while (reader.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace)
            {
                _text.Append(reader.Value);
                if (!reader.Read())
                    break;
            }

            Text = _text.ToString();
        }
    }
}
