namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Identifies one kind of named graphic frame: a slicer's, a timeline's.
/// </summary>
/// <remarks>
/// Excel draws slicers and timelines through the same construct — a <c>xdr:graphicFrame</c> holding
/// a single element that carries nothing but the control's name. What separates them is the
/// <c>a:graphicData/@uri</c> and the name of that element, which is exactly what this carries.
/// <para>
/// <see cref="ChildNamespace"/> currently equals <see cref="GraphicUri"/> for both kinds Excel
/// defines. They are kept apart because nothing in the format requires them to agree, and a future
/// control that separates them would otherwise be silently mis-serialised.
/// </para>
/// </remarks>
/// <param name="GraphicUri">The <c>a:graphicData/@uri</c> Excel resolves the control from.</param>
/// <param name="Prefix">The namespace prefix of the frame's single child, e.g. <c>sle</c>.</param>
/// <param name="LocalName">The local name of that child, e.g. <c>slicer</c>.</param>
/// <param name="ChildNamespace">That child's namespace URI.</param>
internal readonly record struct DrawingFrameSpec(
    string GraphicUri,
    string Prefix,
    string LocalName,
    string ChildNamespace);
