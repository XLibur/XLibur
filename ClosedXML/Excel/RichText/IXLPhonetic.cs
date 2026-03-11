using System;

namespace ClosedXML.Excel;

public interface IXLPhonetic : IEquatable<IXLPhonetic>
{
    string Text { get; }
    int Start { get; }
    int End { get; }
}
