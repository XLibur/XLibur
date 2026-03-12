namespace XLibur.Excel;

internal sealed class XLPhonetic : IXLPhonetic
{
    public XLPhonetic(string text, int start, int end)
    {
        Text = text;
        Start = start;
        End = end;
    }
    public string Text { get; }
    public int Start { get; }
    public int End { get; }

    public bool Equals(IXLPhonetic? other)
    {
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return Text == other.Text && Start == other.Start && End == other.End;
    }
}
