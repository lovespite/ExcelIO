namespace ExcelIO.Formula;

internal static class BuiltinFunctions
{
    internal static void RegisterAll(IFunctionRegistry registry)
    {
        // Math
        registry.Register(new ExcelFunction("SUM", "Math", "Adds all numbers in a range.", 1, int.MaxValue,
            (args, ctx) => Sum(args)));
        registry.Register(new ExcelFunction("AVERAGE", "Math", "Returns the average of numbers.", 1, int.MaxValue,
            (args, ctx) => Average(args)));
        registry.Register(new ExcelFunction("MIN", "Math", "Returns the smallest number.", 1, int.MaxValue,
            (args, ctx) => Min(args)));
        registry.Register(new ExcelFunction("MAX", "Math", "Returns the largest number.", 1, int.MaxValue,
            (args, ctx) => Max(args)));
        registry.Register(new ExcelFunction("COUNT", "Math", "Counts cells that contain numbers.", 1, int.MaxValue,
            (args, ctx) => Count(args)));
        registry.Register(new ExcelFunction("COUNTA", "Math", "Counts cells that are not empty.", 1, int.MaxValue,
            (args, ctx) => CountA(args)));
        registry.Register(new ExcelFunction("ROUND", "Math", "Rounds a number to a specified number of digits.", 2, 2,
            (args, ctx) => Round(args)));
        registry.Register(new ExcelFunction("ABS", "Math", "Returns the absolute value.", 1, 1,
            (args, ctx) => Math.Abs(Convert.ToDouble(args[0]))));
        registry.Register(new ExcelFunction("PRODUCT", "Math", "Multiplies all numbers.", 1, int.MaxValue,
            (args, ctx) => Product(args)));

        // Logic
        registry.Register(new ExcelFunction("IF", "Logic", "Returns one value if condition is true, another if false.", 3, 3,
            (args, ctx) => If(args)));
        registry.Register(new ExcelFunction("IFERROR", "Logic", "Returns value if no error, else fallback.", 2, 2,
            (args, ctx) => IfError(args)));
        registry.Register(new ExcelFunction("AND", "Logic", "Returns TRUE if all arguments are TRUE.", 1, int.MaxValue,
            (args, ctx) => And(args)));
        registry.Register(new ExcelFunction("OR", "Logic", "Returns TRUE if any argument is TRUE.", 1, int.MaxValue,
            (args, ctx) => Or(args)));
        registry.Register(new ExcelFunction("NOT", "Logic", "Reverses the logical value.", 1, 1,
            (args, ctx) => !ToBool(args[0])));

        // Text
        registry.Register(new ExcelFunction("CONCATENATE", "Text", "Joins text strings.", 1, int.MaxValue,
            (args, ctx) => string.Concat(args.Select(a => ToString(a)))));
        registry.Register(new ExcelFunction("LEFT", "Text", "Returns leftmost characters.", 1, 2,
            (args, ctx) => Left(args)));
        registry.Register(new ExcelFunction("RIGHT", "Text", "Returns rightmost characters.", 1, 2,
            (args, ctx) => Right(args)));
        registry.Register(new ExcelFunction("MID", "Text", "Returns characters from the middle.", 3, 3,
            (args, ctx) => Mid(args)));
        registry.Register(new ExcelFunction("LEN", "Text", "Returns the number of characters.", 1, 1,
            (args, ctx) => ToString(args[0]).Length));
        registry.Register(new ExcelFunction("TRIM", "Text", "Removes extra spaces.", 1, 1,
            (args, ctx) => ToString(args[0]).Trim()));
        registry.Register(new ExcelFunction("UPPER", "Text", "Converts text to uppercase.", 1, 1,
            (args, ctx) => ToString(args[0]).ToUpperInvariant()));
        registry.Register(new ExcelFunction("LOWER", "Text", "Converts text to lowercase.", 1, 1,
            (args, ctx) => ToString(args[0]).ToLowerInvariant()));

        // Lookup
        registry.Register(new ExcelFunction("VLOOKUP", "Lookup", "Looks up a value in the first column and returns a value from another column.", 3, 4,
            (args, ctx) => VLookup(args, ctx)));
        registry.Register(new ExcelFunction("MATCH", "Lookup", "Returns the position of a value in a range.", 2, 3,
            (args, ctx) => Match(args, ctx)));
        registry.Register(new ExcelFunction("INDEX", "Lookup", "Returns a value at the given position in a range.", 2, 3,
            (args, ctx) => Index(args, ctx)));

        // Date
        registry.Register(new ExcelFunction("TODAY", "Date", "Returns the current date.", 0, 0,
            (args, ctx) => DateTime.Today.ToShortDateString()));
        registry.Register(new ExcelFunction("NOW", "Date", "Returns the current date and time.", 0, 0,
            (args, ctx) => DateTime.Now.ToString()));
    }

    // ── Math implementations ──

    private static object Sum(IReadOnlyList<object> args)
    {
        var numbers = FlattenNumbers(args);
        return numbers.Sum();
    }

    private static object Average(IReadOnlyList<object> args)
    {
        var numbers = FlattenNumbers(args);
        return numbers.Count == 0 ? 0d : numbers.Average();
    }

    private static object Min(IReadOnlyList<object> args)
    {
        var numbers = FlattenNumbers(args);
        return numbers.Count == 0 ? 0d : numbers.Min();
    }

    private static object Max(IReadOnlyList<object> args)
    {
        var numbers = FlattenNumbers(args);
        return numbers.Count == 0 ? 0d : numbers.Max();
    }

    private static object Count(IReadOnlyList<object> args)
    {
        var numbers = FlattenNumbers(args);
        return numbers.Count;
    }

    private static object CountA(IReadOnlyList<object> args)
    {
        return FlattenValues(args).Count(v => v is not null && v.ToString()?.Length > 0);
    }

    private static object Round(IReadOnlyList<object> args)
    {
        var value = Convert.ToDouble(args[0]);
        var digits = Convert.ToInt32(args[1]);
        return Math.Round(value, digits);
    }

    private static object Product(IReadOnlyList<object> args)
    {
        var numbers = FlattenNumbers(args);
        return numbers.Count == 0 ? 0d : numbers.Aggregate(1.0, (a, b) => a * b);
    }

    // ── Logic implementations ──

    private static object If(IReadOnlyList<object> args) => ToBool(args[0]) ? args[1] : args[2];

    private static object IfError(IReadOnlyList<object> args)
    {
        var val = ToString(args[0]);
        return val.StartsWith("#") ? args[1] : args[0];
    }

    private static object And(IReadOnlyList<object> args) => args.All(ToBool);

    private static object Or(IReadOnlyList<object> args) => args.Any(ToBool);

    // ── Text implementations ──

    private static object Left(IReadOnlyList<object> args)
    {
        var text = ToString(args[0]);
        var count = args.Count > 1 ? Convert.ToInt32(args[1]) : 1;
        if (count >= text.Length) return text;
        return text.Substring(0, count);
    }

    private static object Right(IReadOnlyList<object> args)
    {
        var text = ToString(args[0]);
        var count = args.Count > 1 ? Convert.ToInt32(args[1]) : 1;
        if (count >= text.Length) return text;
        return text.Substring(text.Length - count);
    }

    private static object Mid(IReadOnlyList<object> args)
    {
        var text = ToString(args[0]);
        var start = Convert.ToInt32(args[1]) - 1; // 1-based
        var count = Convert.ToInt32(args[2]);
        if (start >= text.Length || start < 0) return "";
        if (start + count > text.Length) count = text.Length - start;
        return text.Substring(start, count);
    }

    // ── Lookup implementations ──

    private static object VLookup(IReadOnlyList<object> args, IFormulaContext ctx)
    {
        var lookupValue = ToString(args[0]);
        var colIndex = Convert.ToInt32(args[2]) - 1; // 1-based to 0-based
        var exactMatch = args.Count > 3 ? ToBool(args[3]) : false;
        var tableRange = args[1] as List<object>;

        if (tableRange is null) return "#VALUE!";

        // tableRange is a flat list of cell value strings from the range.
        // Guess column count from the range's total size by looking at the first row cells.
        var sheet = ctx.Worksheet;
        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            if (row.Cells.Count == 0) continue;
            var firstCellValue = row.Cells[0].Value;
            var match = string.Equals(firstCellValue, lookupValue, StringComparison.OrdinalIgnoreCase);
            if (match || (!exactMatch && string.Compare(firstCellValue, lookupValue, StringComparison.OrdinalIgnoreCase) > 0))
            {
                if (colIndex < row.Cells.Count)
                    return row.Cells[colIndex].Value;
                return "";
            }
        }
        return "#N/A";
    }

    private static object Match(IReadOnlyList<object> args, IFormulaContext ctx)
    {
        var lookupValue = ToString(args[0]);
        var sheet = ctx.Worksheet;
        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            for (int c = 0; c < row.Cells.Count; c++)
            {
                if (string.Equals(row.Cells[c].Value, lookupValue, StringComparison.OrdinalIgnoreCase))
                    return r + 1; // 1-based
            }
        }
        return "#N/A";
    }

    private static object Index(IReadOnlyList<object> args, IFormulaContext ctx)
    {
        var rowNum = Convert.ToInt32(args[0]) - 1;
        var colNum = args.Count > 1 ? Convert.ToInt32(args[1]) - 1 : 0;
        var sheet = ctx.Worksheet;
        if (rowNum >= 0 && rowNum < sheet.Rows.Count)
        {
            var row = sheet.Rows[rowNum];
            if (colNum >= 0 && colNum < row.Cells.Count)
                return row.Cells[colNum].Value;
        }
        return "#REF!";
    }

    // ── Helpers ──

    private static List<double> FlattenNumbers(IReadOnlyList<object> args)
    {
        var result = new List<double>();
        foreach (var arg in args)
        {
            if (arg is List<object> list)
            {
                result.AddRange(FlattenNumbers(list));
            }
            else if (arg is double d)
            {
                result.Add(d);
            }
            else if (arg is int i)
            {
                result.Add(i);
            }
            else if (double.TryParse(ToString(arg), out var num))
            {
                result.Add(num);
            }
        }
        return result;
    }

    private static List<object?> FlattenValues(IReadOnlyList<object> args)
    {
        var result = new List<object?>();
        foreach (var arg in args)
        {
            if (arg is List<object> list)
                result.AddRange(FlattenValues(list));
            else
                result.Add(arg);
        }
        return result;
    }

    private static string ToString(object? val) => val?.ToString() ?? "";

    private static bool ToBool(object? val)
    {
        if (val is bool b) return b;
        if (val is null) return false;
        var s = val.ToString();
        if (string.IsNullOrEmpty(s)) return false;
        if (double.TryParse(s, out var num)) return num != 0;
        return string.Equals(s, "true", StringComparison.OrdinalIgnoreCase);
    }
}
