using System.Collections;

namespace ExcelIO;

public class XlRow : IReadOnlyList<string>
{
    private readonly XlWorksheet _ws;

    public XlWorksheet Worksheet => _ws;

    public XlRow(XlWorksheet ws)
    {
        _ws = ws;
    }

    public string this[int index]
    {
        get
        {
            if (index < 0 || index >= Cells.Count) return string.Empty;
            return Cells[index].Value;
        }
        set
        {
            Cells[index].Value = value;
        }
    }

    public string this[string columnName]
    {
        get
        {
            return Cell(columnName).Value;
        }
        set
        {
            Cell(columnName).Value = value;
        }
    }

    public XlCell Cell(string columnName)
    {
        var index = _ws.IndexOf(columnName);
        if (index < 0) throw new IndexOutOfRangeException();
        return Cells[index];
    }

    public List<XlCell> Cells { get; set; } = [];

    public int Count => Cells.Count;

    public IEnumerator<string> GetEnumerator()
    {
        return Cells.Select(c => c.Value).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public void Add(string value)
    {
        Cells.Add(new XlCell(this) { Value = value });
    }

    public void AddRange(params string[] values)
    {
        for (int i = 0; i < values.Length; i++)
        {
            var cell = new XlCell(this) { Value = values[i] };
            Cells.Add(cell);
        }
    }

    public void AddRange(IEnumerable<string> values)
    {
        foreach (var value in values)
        {
            var cell = new XlCell(this) { Value = value };
            Cells.Add(cell);
        }
    }

    public XlCell Insert(int index, string value)
    {
        var cell = new XlCell(this) { Value = value };
        Cells.Insert(index, cell);
        return cell;
    }

    public XlCell RemoveAt(int index)
    {
        var cell = Cells[index];
        Cells.RemoveAt(index);
        return cell;
    }

    public void RemoveRange(int index, int count)
    {
        Cells.RemoveRange(index, count);
    }

    public bool Remove(XlCell cell)
    {
        return Cells.Remove(cell);
    }

    public void RemoveAll()
    {
        Cells.Clear();
    }

    public override string ToString()
    {
        return this.ToString(',');
    }

    public string ToString(char separator)
    {
        return string.Join(separator, Cells.Select(c => c.Value));
    }
}
