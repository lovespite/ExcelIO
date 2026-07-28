namespace ExcelIO.Formula;

public sealed class FormulaEvaluator
{
    private readonly IFunctionRegistry _functions;

    public FormulaEvaluator(IFunctionRegistry functions)
    {
        _functions = functions;
    }

    public object Evaluate(Expr expr, IFormulaContext ctx)
    {
        return Eval(expr, ctx);
    }

    public object EvaluateFormula(string formula, IFormulaContext ctx)
    {
        var ast = FormulaParser.Parse(formula);
        return Eval(ast, ctx);
    }

    private object Eval(Expr expr, IFormulaContext ctx)
    {
        switch (expr)
        {
            case NumberExpr n: return n.Value;
            case TextExpr t: return t.Value;
            case BoolExpr b: return b.Value;
            case ErrorExpr e: return e.Error;

            case CellRefExpr r:
                return EvalCellRef(r, ctx);

            case RangeRefExpr r:
                return EvalRangeRef(r, ctx);

            case UnaryExpr u:
                return EvalUnary(u, ctx);

            case BinaryExpr b:
                return EvalBinary(b, ctx);

            case FunctionExpr f:
                return EvalFunction(f, ctx);

            default:
                return "#VALUE!";
        }
    }

    private object EvalCellRef(CellRefExpr r, IFormulaContext ctx)
    {
        var sheet = r.Sheet != null ? ctx.GetSheet(r.Sheet) : ctx.Worksheet;
        if (sheet is null) return "#REF!";

        if (r.Row < 0 || r.Row >= sheet.Rows.Count) return 0d;
        var row = sheet.Rows[r.Row];
        if (r.Col < 0 || r.Col >= row.Cells.Count) return 0d;
        var cell = row.Cells[r.Col];
        return ParseCellValue(cell.Value);
    }

    private object EvalRangeRef(RangeRefExpr r, IFormulaContext ctx)
    {
        var sheet = r.Sheet != null ? ctx.GetSheet(r.Sheet) : ctx.Worksheet;
        if (sheet is null) return new List<object> { "#REF!" };

        var values = new List<object>();
        int startRow = r.StartRow == int.MinValue ? 0 : Math.Max(0, r.StartRow);
        int endRow = r.EndRow == int.MaxValue ? sheet.Rows.Count - 1 : Math.Min(sheet.Rows.Count - 1, r.EndRow);

        for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++)
        {
            if (rowIdx >= sheet.Rows.Count) break;
            var row = sheet.Rows[rowIdx];
            int startCol = Math.Max(0, r.StartCol);
            int endCol = Math.Min(row.Cells.Count - 1, r.EndCol);
            for (int colIdx = startCol; colIdx <= endCol; colIdx++)
            {
                // Return raw string values — functions handle their own parsing
                values.Add(row.Cells[colIdx].Value);
            }
        }
        return values;
    }

    private object EvalUnary(UnaryExpr u, IFormulaContext ctx)
    {
        var operand = Eval(u.Operand, ctx);
        switch (u.Op)
        {
            case UnaryOp.Neg:
                return -Convert.ToDouble(operand);
            case UnaryOp.Plus:
                return Convert.ToDouble(operand);
            case UnaryOp.Percent:
                return Convert.ToDouble(operand) / 100.0;
            default:
                return "#VALUE!";
        }
    }

    private object EvalBinary(BinaryExpr b, IFormulaContext ctx)
    {
        // Range operator — handled specially (not really a binary op in practice,
        // but if we get here it was parsed as such)
        if (b.Op == BinaryOp.Concat)
        {
            var left = Eval(b.Left, ctx);
            var right = Eval(b.Right, ctx);
            return CoerceToString(left) + CoerceToString(right);
        }

        // For comparison operators, coerce to same type
        var lv = Eval(b.Left, ctx);
        var rv = Eval(b.Right, ctx);

        if (IsComparison(b.Op))
        {
            return Compare(lv, b.Op, rv);
        }

        // Arithmetic
        double ln = Convert.ToDouble(lv);
        double rn = Convert.ToDouble(rv);

        return b.Op switch
        {
            BinaryOp.Add => ln + rn,
            BinaryOp.Sub => ln - rn,
            BinaryOp.Mul => ln * rn,
            BinaryOp.Div => rn == 0 ? "#DIV/0!" : ln / rn,
            BinaryOp.Pow => Math.Pow(ln, rn),
            _ => "#VALUE!",
        };
    }

    private object EvalFunction(FunctionExpr f, IFormulaContext ctx)
    {
        var func = _functions.Find(f.Name);
        if (func is null) return "#NAME?";

        var args = new List<object>();
        foreach (var arg in f.Arguments)
        {
            args.Add(Eval(arg, ctx));
        }

        try
        {
            return func.Delegate(args.AsReadOnly(), ctx);
        }
        catch
        {
            return "#VALUE!";
        }
    }

    // ── Helpers ──

    private static object ParseCellValue(string value)
    {
        if (string.IsNullOrEmpty(value)) return 0d;
        if (double.TryParse(value, out var num)) return num;
        if (bool.TryParse(value, out var b)) return b;
        return value;
    }

    private static string CoerceToString(object val) => val?.ToString() ?? "";

    private static bool IsComparison(BinaryOp op) => op is
        BinaryOp.Eq or BinaryOp.Ne or BinaryOp.Lt
        or BinaryOp.Le or BinaryOp.Gt or BinaryOp.Ge;

    private static object Compare(object left, BinaryOp op, object right)
    {
        int cmp;

        if (left is double ld && right is double rd)
            cmp = ld.CompareTo(rd);
        else
            cmp = string.Compare(CoerceToString(left), CoerceToString(right), StringComparison.OrdinalIgnoreCase);

        return op switch
        {
            BinaryOp.Eq => cmp == 0,
            BinaryOp.Ne => cmp != 0,
            BinaryOp.Lt => cmp < 0,
            BinaryOp.Le => cmp <= 0,
            BinaryOp.Gt => cmp > 0,
            BinaryOp.Ge => cmp >= 0,
            _ => false,
        };
    }
}
