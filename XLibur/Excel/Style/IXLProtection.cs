using System;

namespace XLibur.Excel;

public interface IXLProtection : IEquatable<IXLProtection>
{
    bool Locked { get; set; }

    bool Hidden { get; set; }

    IXLStyle SetLocked(); IXLStyle SetLocked(bool value);

    IXLStyle SetHidden(); IXLStyle SetHidden(bool value);
}
