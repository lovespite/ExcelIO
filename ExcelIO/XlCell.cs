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
            Formula = null; // Clear formula when setting literal value
            OnValueChanged?.Invoke(this);
        }
    }

    public void SetCalculatedValue(string value)
    {
        _value = value;
    }

    /// <summary>
    /// Internal hook for formula engine dirty-tracking. Set by ExcelIO.Formula on init.
    /// </summary>
    public static Action<XlCell>? OnValueChanged { get; set; }

    public XlStyle? Style { get; set; }

    /// <summary>
    /// Formula text (e.g., "=SUM(A1:A2)"). Null if this is not a formula cell.
    /// </summary>
    public string? Formula { get; set; }

    /// <summary>
    /// True if this cell contains a formula.
    /// </summary>
    public bool HasFormula => !string.IsNullOrEmpty(Formula);

    /// <summary>
    /// Set formula and cached value.
    /// </summary>
    public void SetFormula(string formula, string cachedValue = "")
    {
        Formula = formula;
        _value = cachedValue;
    }

    public override string ToString()
    {
        return Value;
    }

    public bool IsNull => ReferenceEquals(this, Null) || _row is null;

    public static XlCell Null { get; } = new XlCell(null!);
}
