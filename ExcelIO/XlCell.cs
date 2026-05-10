namespace ExcelIO;

public class XlCell
{
    private readonly XlRow _row;
    private string _value = string.Empty;

    public XlRow Row => _row;

    public XlCell(XlRow row)
    {
        _row = row;
    }

    public string Value
    {
        get => IsNull ? string.Empty : _value;
        set
        {
            if (IsNull) return;
            _value = value;
        }
    }

    public XlStyle? Style { get; set; }

    public override string ToString()
    {
        return Value;
    }

    public bool IsNull => ReferenceEquals(this, Null) || _row is null;

    public static XlCell Null { get; } = new XlCell(null!);
}
