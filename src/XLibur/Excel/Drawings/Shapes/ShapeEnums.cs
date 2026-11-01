namespace XLibur.Excel.Drawings;

/// <summary>The preset shapes XLibur names. <see cref="IXLShapes.AddPreset"/> adds any other DrawingML geometry.</summary>
public enum XLShapeType
{
    /// <summary>A text box: <c>rect</c>, with <c>cNvSpPr/@txBox</c> set.</summary>
    TextBox,

    /// <summary><c>rect</c></summary>
    Rectangle,

    /// <summary><c>roundRect</c></summary>
    RoundedRectangle,

    /// <summary><c>ellipse</c></summary>
    Ellipse,

    /// <summary><c>triangle</c></summary>
    Triangle,

    /// <summary><c>rightArrow</c></summary>
    RightArrow,

    /// <summary><c>leftArrow</c></summary>
    LeftArrow,

    /// <summary><c>upArrow</c></summary>
    UpArrow,

    /// <summary><c>downArrow</c></summary>
    DownArrow,

    /// <summary><c>line</c></summary>
    Line,

    /// <summary><c>wedgeRoundRectCallout</c></summary>
    RoundedCallout,

    /// <summary><c>wedgeRectCallout</c></summary>
    WedgeRectCallout,

    /// <summary><c>diamond</c></summary>
    Diamond,

    /// <summary><c>pentagon</c></summary>
    Pentagon,

    /// <summary><c>chevron</c></summary>
    Chevron,

    /// <summary><c>star5</c></summary>
    Star5,

    /// <summary><c>can</c></summary>
    Can,

    /// <summary><c>cloud</c></summary>
    Cloud,

    /// <summary>A geometry the enum does not name; <see cref="IXLShape.PresetGeometry"/> holds it.</summary>
    Custom,
}

/// <summary>The horizontal alignment of a paragraph of shape text, <c>a:pPr/@algn</c>.</summary>
public enum XLTextAlignment
{
    /// <summary><c>l</c>, the default.</summary>
    Left,

    /// <summary><c>ctr</c></summary>
    Center,

    /// <summary><c>r</c></summary>
    Right,

    /// <summary><c>just</c></summary>
    Justified,

    /// <summary><c>dist</c></summary>
    Distributed,
}

/// <summary>Where shape text sits between the top and bottom of the shape, <c>a:bodyPr/@anchor</c>.</summary>
public enum XLTextVerticalAlignment
{
    /// <summary><c>t</c>, the default.</summary>
    Top,

    /// <summary><c>ctr</c></summary>
    Center,

    /// <summary><c>b</c></summary>
    Bottom,
}

/// <summary>The underline of a run of shape text, <c>a:rPr/@u</c>.</summary>
public enum XLShapeTextUnderline
{
    /// <summary><c>none</c></summary>
    None,

    /// <summary><c>sng</c></summary>
    Single,

    /// <summary><c>dbl</c></summary>
    Double,
}
