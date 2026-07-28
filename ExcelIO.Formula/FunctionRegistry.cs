namespace ExcelIO.Formula;

public sealed class FunctionRegistry : IFunctionRegistry
{
    private readonly Dictionary<string, ExcelFunction> _functions = new(StringComparer.OrdinalIgnoreCase);

    public int Count => _functions.Count;

    public void Register(ExcelFunction function)
    {
        if (string.IsNullOrWhiteSpace(function.Name))
            throw new ArgumentException("Function name cannot be empty.", nameof(function));
        _functions[function.Name] = function;
    }

    public void RegisterAll<T>() where T : IFormulaFunction
        => RegisterAll(typeof(T));

    public void RegisterAll(Type type)
    {
        if (!typeof(IFormulaFunction).IsAssignableFrom(type))
            throw new ArgumentException($"Type {type.Name} must implement IFormulaFunction.", nameof(type));

        var methods = type.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
        foreach (var method in methods)
        {
            var attr = method.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
                             .FirstOrDefault() as ExcelFunctionAttribute;
            if (attr is null) continue;

            var del = (ExcelFunctionDelegate)Delegate.CreateDelegate(
                typeof(ExcelFunctionDelegate), method);

            var func = new ExcelFunction(attr.Name, attr.Category, attr.Description,
                                         attr.MinArgs, attr.MaxArgs, del);
            Register(func);
        }
    }

    public void RegisterAssembly(System.Reflection.Assembly assembly)
    {
        foreach (var type in assembly.GetTypes())
        {
            if (typeof(IFormulaFunction).IsAssignableFrom(type) && !type.IsAbstract && !type.IsInterface)
            {
                RegisterAll(type);
            }
        }
    }

    public ExcelFunction? Find(string name)
        => _functions.TryGetValue(name, out var func) ? func : null;

    public bool Exists(string name)
        => _functions.ContainsKey(name);

    public IReadOnlyList<ExcelFunction> GetAll()
        => _functions.Values.ToList();

    public IReadOnlyList<ExcelFunction> GetByCategory(string category)
        => _functions.Values.Where(f => string.Equals(f.Category, category, StringComparison.OrdinalIgnoreCase)).ToList();

    public IReadOnlyList<string> GetCategories()
        => _functions.Values.Select(f => f.Category).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
}
