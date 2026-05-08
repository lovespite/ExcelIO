using System.Collections;

namespace ExcelIO;

public class XlWorkbook : IReadOnlyList<XlWorksheet>
{
    public XlWorksheet this[int index] => ((IReadOnlyList<XlWorksheet>)Worksheets)[index];

    public List<XlWorksheet> Worksheets { get; set; } = [];
    public string[] WorksheetNames => [.. Worksheets.Select(w => w.Name)];

    public int Count => ((IReadOnlyCollection<XlWorksheet>)Worksheets).Count;

    public IEnumerator<XlWorksheet> GetEnumerator()
    {
        return ((IEnumerable<XlWorksheet>)Worksheets).GetEnumerator();
    }

    public XlWorksheet? GetWorkSheet(string sheetName)
    {
        return Worksheets.FirstOrDefault(w => w.Name == sheetName);
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable)Worksheets).GetEnumerator();
    }

    public XlWorksheet NewWorksheet(string sheetName)
    {
        if (Worksheets.Any(x => string.Equals(sheetName, x.Name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Worksheet with name '{sheetName}' already exists.");

        var ws = new XlWorksheet(this) { Name = sheetName };
        Worksheets.Add(ws);
        return ws;
    }
}
