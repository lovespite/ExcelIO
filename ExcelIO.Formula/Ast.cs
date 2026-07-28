namespace ExcelIO.Formula;

// ── Base ──

public abstract class Expr
{
    private protected Expr() { }
}

// ── Leaf values ──

public sealed class NumberExpr : Expr
{
    public double Value { get; }
    public NumberExpr(double value) => Value = value;
}

public sealed class TextExpr : Expr
{
    public string Value { get; }
    public TextExpr(string value) => Value = value;
}

public sealed class BoolExpr : Expr
{
    public bool Value { get; }
    public BoolExpr(bool value) => Value = value;
}

public sealed class ErrorExpr : Expr
{
    public string Error { get; }
    public ErrorExpr(string error) => Error = error;
}

// ── References ──

public sealed class CellRefExpr : Expr
{
    public string? Sheet { get; }
    public int Row { get; }       // zero-based
    public int Col { get; }       // zero-based
    public bool ColAbs { get; }
    public bool RowAbs { get; }

    public CellRefExpr(string? sheet, int row, int col, bool colAbs, bool rowAbs)
    {
        Sheet = sheet;
        Row = row;
        Col = col;
        ColAbs = colAbs;
        RowAbs = rowAbs;
    }
}

public sealed class RangeRefExpr : Expr
{
    public string? Sheet { get; }
    public int StartRow { get; }
    public int StartCol { get; }
    public int EndRow { get; }
    public int EndCol { get; }
    public bool ColAbsStart { get; }
    public bool RowAbsStart { get; }
    public bool ColAbsEnd { get; }
    public bool RowAbsEnd { get; }

    public RangeRefExpr(string? sheet,
        int startRow, int startCol, int endRow, int endCol,
        bool colAbsStart, bool rowAbsStart, bool colAbsEnd, bool rowAbsEnd)
    {
        Sheet = sheet;
        StartRow = startRow;
        StartCol = startCol;
        EndRow = endRow;
        EndCol = endCol;
        ColAbsStart = colAbsStart;
        RowAbsStart = rowAbsStart;
        ColAbsEnd = colAbsEnd;
        RowAbsEnd = rowAbsEnd;
    }
}

// ── Operators ──

public enum BinaryOp { Add, Sub, Mul, Div, Pow, Concat, Eq, Ne, Lt, Le, Gt, Ge }
public enum UnaryOp { Neg, Percent, Plus }

public sealed class BinaryExpr : Expr
{
    public BinaryOp Op { get; }
    public Expr Left { get; }
    public Expr Right { get; }

    public BinaryExpr(BinaryOp op, Expr left, Expr right)
    {
        Op = op;
        Left = left;
        Right = right;
    }
}

public sealed class UnaryExpr : Expr
{
    public UnaryOp Op { get; }
    public Expr Operand { get; }

    public UnaryExpr(UnaryOp op, Expr operand)
    {
        Op = op;
        Operand = operand;
    }
}

// ── Function call ──

public sealed class FunctionExpr : Expr
{
    public string Name { get; }
    public IReadOnlyList<Expr> Arguments { get; }

    public FunctionExpr(string name, IReadOnlyList<Expr> arguments)
    {
        Name = name;
        Arguments = arguments;
    }
}
