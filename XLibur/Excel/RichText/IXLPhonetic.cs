using System;

namespace XLibur.Excel;

public interface IXLPhonetic : IEquatable<IXLPhonetic>
{
    string Text { get; }
    int Start { get; }
    int End { get; }
}
