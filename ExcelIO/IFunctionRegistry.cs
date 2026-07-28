using System.Reflection;

namespace ExcelIO;

public delegate object ExcelFunctionDelegate(IReadOnlyList<object> args, IFormulaContext ctx);

public interface IFunctionRegistry
{
    void Register(ExcelFunction function);
    void RegisterAll<T>() where T : IFormulaFunction;
    void RegisterAll(Type type);
    void RegisterAssembly(Assembly assembly);

    ExcelFunction? Find(string name);
    bool Exists(string name);
    IReadOnlyList<ExcelFunction> GetAll();
    IReadOnlyList<ExcelFunction> GetByCategory(string category);
    IReadOnlyList<string> GetCategories();

    int Count { get; }
}

public sealed class ExcelFunction
{
    public string Name { get; }
    public string Category { get; }
    public string Description { get; }
    public int MinArgs { get; }
    public int MaxArgs { get; }
    public ExcelFunctionDelegate Delegate { get; }

    public ExcelFunction(string name, string category, string description,
                         int minArgs, int maxArgs, ExcelFunctionDelegate del)
    {
        Name = name;
        Category = category;
        Description = description;
        MinArgs = minArgs;
        MaxArgs = maxArgs;
        Delegate = del;
    }
}

[AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
public class ExcelFunctionAttribute : Attribute
{
    public string Name { get; init; } = string.Empty;
    public string Category { get; init; } = "Custom";
    public string Description { get; init; } = string.Empty;
    public int MinArgs { get; init; } = 0;
    public int MaxArgs { get; init; } = int.MaxValue;
}

public interface IFormulaFunction { }
