using System;

namespace XLibur.Excel.IO;

/// <summary>
/// An exception thrown when a workbook's structure is not what loading requires. That covers
/// both a problem with the data inside an XML part and a problem with the parts of the package
/// themselves — a required part absent, or a relationship pointing at a part that is not there.
/// The exception messages are rather generic and not very helpful, but they
/// aren't supposed to be. If this exception is thrown, there is either
/// a problem with the producer of a workbook or XLibur. Both should do
/// investigation based on the file causing an error.
/// </summary>
public sealed class PartStructureException : Exception
{
    private PartStructureException(string message, string? detail = null)
        : base(detail is null ? message : message[..^1] + " (" + detail + ").")
    {
    }

    private PartStructureException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    /// <summary>
    /// Create a new exception with info that some element that should be present in a workbook
    /// is missing.
    /// </summary>
    /// <param name="missingElementDesc">optional info about what element is missing.</param>
    internal static Exception ExpectedElementNotFound(string? missingElementDesc = null)
    {
        return new PartStructureException("The structure of XML expected a certain kind of element, but it isn't there.", missingElementDesc);
    }

    internal static Exception IncorrectElementsCount()
    {
        return new PartStructureException("There is a problem with element structure in XML, the number of elements found is not what was expected.");
    }

    internal static Exception MissingAttribute()
    {
        return new PartStructureException("XML doesn't contain a required attribute.");
    }

    internal static Exception MissingAttribute(string attributeName)
    {
        return new PartStructureException($"XML doesn't contain a required attribute '{attributeName}'.");
    }

    internal static Exception IncorrectAttributeFormat()
    {
        return new PartStructureException("The attribute has a value in an incorrect format.");
    }

    public static PartStructureException IncorrectElementFormat(string elementName)
    {
        return new PartStructureException($"The element '{elementName}' is missing required child elements or attributes required by the workbook constraints.");
    }

    internal static Exception IncorrectAttributeValue()
    {
        return new PartStructureException("The value of attribute doesn't make sense with the rest of data of a workbook (e.g. reference that doesn't exist).");
    }

    internal static Exception InvalidAttributeValue(string attributeValue)
    {
        return new PartStructureException($"The value of attribute '{attributeValue}' is not valid value for the attribute.");
    }

    public static PartStructureException RequiredElementIsMissing()
    {
        return new PartStructureException("The XML schema requires an element, but it is not present.");
    }

    /// <summary>
    /// Create a new exception for a package that does not contain a part that loading requires.
    /// Unlike the element-level factories, this one names the part: a caller handed a package
    /// with no workbook in it can do nothing useful with a generic message.
    /// </summary>
    /// <param name="partName">The package-relative name of the absent part, e.g. <c>/xl/workbook.xml</c>.</param>
    internal static PartStructureException MissingPart(string partName)
    {
        return new PartStructureException($"The package does not contain the required part '{partName}'.");
    }

    /// <summary>
    /// Create a new exception for a package whose relationships name a part that is not present.
    /// The underlying package reader signals this with an exception type of its own; this factory
    /// exists so that type does not escape XLibur's public surface.
    /// </summary>
    /// <param name="innerException">The exception raised by the package reader.</param>
    internal static PartStructureException ReferencedPartIsMissing(Exception innerException)
    {
        return new PartStructureException(
            "The package declares a relationship to a part that is not present in the package.",
            innerException);
    }

    /// <summary>
    /// Create a new exception for a package that could not be opened as an OPC package at all —
    /// a malformed archive, or one whose content types are unusable.
    /// </summary>
    /// <param name="innerException">The exception raised by the package reader.</param>
    internal static PartStructureException PackageCannotBeOpened(Exception innerException)
    {
        return new PartStructureException(
            "The stream could not be opened as a spreadsheet package.",
            innerException);
    }
}
