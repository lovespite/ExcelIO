using ExcelIO.Formula;

namespace ExcelIO.Formula.Test;

public class FormulaParserTests
{
    [Theory]
    [InlineData("42", 42d)]
    [InlineData("3.14", 3.14)]
    [InlineData("0", 0d)]
    public void Parse_SimpleNumber(string input, double expected)
    {
        var expr = FormulaParser.Parse(input);
        Assert.IsType<NumberExpr>(expr);
        Assert.Equal(expected, ((NumberExpr)expr).Value);
    }

    [Fact]
    public void Parse_StringLiteral()
    {
        var expr = FormulaParser.Parse("\"hello\"");
        Assert.IsType<TextExpr>(expr);
        Assert.Equal("hello", ((TextExpr)expr).Value);
    }

    [Theory]
    [InlineData("=TRUE", true)]
    [InlineData("=FALSE", false)]
    [InlineData("true", true)]
    [InlineData("false", false)]
    public void Parse_BoolLiteral(string input, bool expected)
    {
        var expr = FormulaParser.Parse(input);
        Assert.IsType<BoolExpr>(expr);
        Assert.Equal(expected, ((BoolExpr)expr).Value);
    }

    [Theory]
    [InlineData("=1+2", BinaryOp.Add, 1d, 2d)]
    [InlineData("=3-1", BinaryOp.Sub, 3d, 1d)]
    [InlineData("=2*3", BinaryOp.Mul, 2d, 3d)]
    [InlineData("=6/2", BinaryOp.Div, 6d, 2d)]
    [InlineData("=2^3", BinaryOp.Pow, 2d, 3d)]
    public void Parse_BinaryArithmetic(string input, BinaryOp op, double left, double right)
    {
        var expr = FormulaParser.Parse(input);
        Assert.IsType<BinaryExpr>(expr);
        var bin = (BinaryExpr)expr;
        Assert.Equal(op, bin.Op);
        Assert.IsType<NumberExpr>(bin.Left);
        Assert.IsType<NumberExpr>(bin.Right);
        Assert.Equal(left, ((NumberExpr)bin.Left).Value);
        Assert.Equal(right, ((NumberExpr)bin.Right).Value);
    }

    [Fact]
    public void Parse_OperatorPrecedence()
    {
        // 1+2*3 should be 1+(2*3), not (1+2)*3
        var expr = FormulaParser.Parse("=1+2*3");
        Assert.IsType<BinaryExpr>(expr);
        var bin = (BinaryExpr)expr;
        Assert.Equal(BinaryOp.Add, bin.Op);
        Assert.IsType<NumberExpr>(bin.Left);
        Assert.IsType<BinaryExpr>(bin.Right); // right side is 2*3
        Assert.Equal(BinaryOp.Mul, ((BinaryExpr)bin.Right).Op);
    }

    [Fact]
    public void Parse_FunctionCall_NoArgs()
    {
        var expr = FormulaParser.Parse("=TODAY()");
        Assert.IsType<FunctionExpr>(expr);
        var func = (FunctionExpr)expr;
        Assert.Equal("TODAY", func.Name);
        Assert.Empty(func.Arguments);
    }

    [Fact]
    public void Parse_FunctionCall_WithArgs()
    {
        var expr = FormulaParser.Parse("=SUM(1,2,3)");
        Assert.IsType<FunctionExpr>(expr);
        var func = (FunctionExpr)expr;
        Assert.Equal("SUM", func.Name);
        Assert.Equal(3, func.Arguments.Count);
    }

    [Fact]
    public void Parse_FunctionCall_Nested()
    {
        var expr = FormulaParser.Parse("=SUM(A1,MAX(B1:B3))");
        Assert.IsType<FunctionExpr>(expr);
        var func = (FunctionExpr)expr;
        Assert.Equal(2, func.Arguments.Count);
        Assert.IsType<FunctionExpr>(func.Arguments[1]);
    }

    [Fact]
    public void Parse_CellRef()
    {
        var expr = FormulaParser.Parse("=A1");
        Assert.IsType<CellRefExpr>(expr);
        var cr = (CellRefExpr)expr;
        Assert.Equal(0, cr.Row);
        Assert.Equal(0, cr.Col);
        Assert.False(cr.ColAbs);
        Assert.False(cr.RowAbs);
    }

    [Fact]
    public void Parse_CellRef_Absolute()
    {
        var expr = FormulaParser.Parse("=$A$1");
        Assert.IsType<CellRefExpr>(expr);
        var cr = (CellRefExpr)expr;
        Assert.True(cr.ColAbs);
        Assert.True(cr.RowAbs);
    }

    [Fact]
    public void Parse_CellRef_MixedAbsolute()
    {
        var expr = FormulaParser.Parse("=A$1");
        Assert.IsType<CellRefExpr>(expr);
        var cr = (CellRefExpr)expr;
        Assert.False(cr.ColAbs);
        Assert.True(cr.RowAbs);
    }

    [Fact]
    public void Parse_RangeRef()
    {
        var expr = FormulaParser.Parse("=A1:B3");
        Assert.IsType<RangeRefExpr>(expr);
        var rr = (RangeRefExpr)expr;
        Assert.Equal(0, rr.StartRow);
        Assert.Equal(0, rr.StartCol);
        Assert.Equal(2, rr.EndRow);
        Assert.Equal(1, rr.EndCol);
    }

    [Fact]
    public void Parse_ColumnOnlyRef()
    {
        var expr = FormulaParser.Parse("=A");
        Assert.IsType<RangeRefExpr>(expr);
        var rr = (RangeRefExpr)expr;
        Assert.Equal(0, rr.StartCol);
        Assert.Equal(0, rr.EndCol);
    }

    [Fact]
    public void Parse_Comparison()
    {
        var expr = FormulaParser.Parse("=1>0");
        Assert.IsType<BinaryExpr>(expr);
        var bin = (BinaryExpr)expr;
        Assert.Equal(BinaryOp.Gt, bin.Op);
    }

    [Fact]
    public void Parse_Concatenation()
    {
        var expr = FormulaParser.Parse("=\"hello\"&\"world\"");
        Assert.IsType<BinaryExpr>(expr);
        var bin = (BinaryExpr)expr;
        Assert.Equal(BinaryOp.Concat, bin.Op);
    }

    [Fact]
    public void Parse_Malformed_ReturnsError()
    {
        var expr = FormulaParser.Parse("=SUM(");
        Assert.IsType<ErrorExpr>(expr);
    }

    [Fact]
    public void Parse_UnaryMinus()
    {
        var expr = FormulaParser.Parse("=-5");
        Assert.IsType<UnaryExpr>(expr);
        var un = (UnaryExpr)expr;
        Assert.Equal(UnaryOp.Neg, un.Op);
        Assert.IsType<NumberExpr>(un.Operand);
    }

    [Fact]
    public void Parse_EmptyString_ReturnsError()
    {
        var expr = FormulaParser.Parse("");
        Assert.IsType<ErrorExpr>(expr);
    }

    [Fact]
    public void Parse_IfStatement()
    {
        var expr = FormulaParser.Parse("=IF(A1>10,\"big\",\"small\")");
        Assert.IsType<FunctionExpr>(expr);
        var func = (FunctionExpr)expr;
        Assert.Equal("IF", func.Name);
        Assert.Equal(3, func.Arguments.Count);
    }

    [Fact]
    public void Parse_Percent()
    {
        var expr = FormulaParser.Parse("=50%");
        Assert.IsType<UnaryExpr>(expr);
        var un = (UnaryExpr)expr;
        Assert.Equal(UnaryOp.Percent, un.Op);
    }
}
