namespace XLibur.Excel;

/// <summary>
/// How finely a timeline's band is divided.
/// </summary>
/// <remarks>
/// The numbers are the values Excel writes to <c>x15:timeline/@level</c>, not an XLibur invention.
/// A file may carry a value outside this set; <see cref="IXLTimeline.Level"/> is a projection over
/// the raw number, which is preserved through a save either way.
/// </remarks>
public enum XLTimelineLevel
{
    Years = 0,
    Quarters = 1,
    Months = 2,
    Days = 3,
}
