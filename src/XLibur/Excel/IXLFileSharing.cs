
namespace XLibur.Excel;

public interface IXLFileSharing
{
    bool ReadOnlyRecommended { get; set; }

    string? UserName { get; set; }
}
