using ExcelIO.Formula;

namespace ExcelIO.Formula.Test;

public class FormulaEngineSmokeTests
{
    [Fact]
    public void Engine_CanBeCreated()
    {
        var engine = new FormulaEngine();
        Assert.NotNull(engine);
        Assert.NotNull(engine.Functions);
        Assert.True(engine.Functions.Count > 0);
    }

    [Fact]
    public void FunctionRegistry_HasBuiltinFunctions()
    {
        var engine = new FormulaEngine();
        Assert.True(engine.Functions.Exists("SUM"));
        Assert.True(engine.Functions.Exists("AVERAGE"));
        Assert.True(engine.Functions.Exists("IF"));
        Assert.True(engine.Functions.Exists("VLOOKUP"));
        Assert.True(engine.Functions.Exists("CONCATENATE"));
    }

    [Fact]
    public void FunctionRegistry_Categories()
    {
        var engine = new FormulaEngine();
        var categories = engine.Functions.GetCategories();
        Assert.Contains("Math", categories);
        Assert.Contains("Logic", categories);
        Assert.Contains("Text", categories);
        Assert.Contains("Lookup", categories);
        Assert.Contains("Date", categories);
    }

    [Fact]
    public void FunctionRegistry_Find()
    {
        var engine = new FormulaEngine();
        var sumFunc = engine.Functions.Find("sum");
        Assert.NotNull(sumFunc);
        Assert.Equal("SUM", sumFunc!.Name);
        Assert.Equal("Math", sumFunc.Category);
    }

    [Fact]
    public void FunctionRegistry_BuiltinSum_Works()
    {
        var engine = new FormulaEngine();
        var sumFunc = engine.Functions.Find("SUM")!;
        var result = sumFunc.Delegate(new object[] { 1d, 2d, 3d }, null!);
        Assert.Equal(6d, result);
    }

    [Fact]
    public void Register_CustomFunction()
    {
        var engine = new FormulaEngine();
        engine.Functions.Register(new ExcelFunction(
            "DOUBLE", "Custom", "Doubles a number", 1, 1,
            (args, ctx) => Convert.ToDouble(args[0]) * 2));

        var func = engine.Functions.Find("DOUBLE");
        Assert.NotNull(func);
        Assert.Equal(2d, func!.Delegate(new object[] { 1d }, null!));
    }

    [Fact]
    public void RegisterAll_Class()
    {
        var engine = new FormulaEngine();
        engine.Functions.RegisterAll<TestFunctions>();

        var pv = engine.Functions.Find("TEST.PV");
        Assert.NotNull(pv);
        Assert.Equal("Financial", pv!.Category);
    }

    [Fact]
    public void XlHelper_FormulaEngine_Integration()
    {
        var engine = new FormulaEngine();
        XlHelper.FormulaEngine = engine;

        Assert.Same(engine, XlHelper.FormulaEngine);

        XlHelper.FormulaEngine = null; // cleanup
    }
}

public class TestFunctions : IFormulaFunction
{
    [ExcelFunction(Name = "TEST.PV", Category = "Financial",
        Description = "Present value", MinArgs = 3, MaxArgs = 3)]
    public static object Pv(IReadOnlyList<object> args, IFormulaContext ctx)
        => Convert.ToDouble(args[0]) + Convert.ToDouble(args[1]) + Convert.ToDouble(args[2]);
}
