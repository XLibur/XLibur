

namespace XLibur.Excel;

internal sealed class XLFileSharing : IXLFileSharing
{
    public bool ReadOnlyRecommended { get; set; }
    public string? UserName { get; set; }
}
