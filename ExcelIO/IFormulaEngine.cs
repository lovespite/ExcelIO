namespace ExcelIO;

public interface IFormulaEngine
{
    string? Evaluate(XlCell cell, IFormulaContext context);
    void Calculate(XlWorksheet worksheet);
    void Calculate(XlWorkbook workbook);
    IReadOnlyList<CircularReference> CircularReferences { get; }
    IFunctionRegistry Functions { get; }
}

public interface IFormulaContext
{
    XlCell? GetCell(int row, int col);
    XlWorksheet Worksheet { get; }
    XlWorksheet? GetSheet(string name);
}

public sealed class CircularReference
{
    public IReadOnlyList<(int Row, int Col)> Path { get; }

    public CircularReference(IReadOnlyList<(int Row, int Col)> path)
    {
        Path = path;
    }
}
