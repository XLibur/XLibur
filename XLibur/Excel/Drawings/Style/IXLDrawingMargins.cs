namespace XLibur.Excel;

public interface IXLDrawingMargins
{
    bool Automatic { get; set; }

    /// <summary>
    /// Left margin in inches.
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// Right margin in inches.
    /// </summary>
    double Right { get; set; }

    /// <summary>
    /// Top margin in inches.
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// Bottom margin in inches.
    /// </summary>
    double Bottom { get; set; }

    /// <summary>
    /// Set <see cref="Left"/>, <see cref="Top"/>, <see cref="Right"/>, <see cref="Bottom"/> margins at once.
    /// </summary>
    double All { set; }

    IXLDrawingStyle SetAutomatic(); IXLDrawingStyle SetAutomatic(bool value);
    IXLDrawingStyle SetLeft(double value);
    IXLDrawingStyle SetRight(double value);
    IXLDrawingStyle SetTop(double value);
    IXLDrawingStyle SetBottom(double value);
    IXLDrawingStyle SetAll(double value);

}
