namespace ExcelIO.Formula;

/// <summary>
/// Builds a dependency graph for formula cells in a worksheet and provides
/// topological ordering for recalculation.
/// </summary>
public sealed class DependencyGraph
{
    // cell → cells that depend on it
    private readonly Dictionary<(int Row, int Col), HashSet<(int Row, int Col)>> _dependents = new();
    // formula cell → cells it depends on
    private readonly Dictionary<(int Row, int Col), List<(int Row, int Col)>> _precedents = new();
    // all formula cell positions
    private readonly List<(int Row, int Col)> _formulaCells = new();

    /// <summary>
    /// Cells that depend on (row, col). Empty if none.
    /// </summary>
    public IReadOnlySet<(int Row, int Col)> GetDependents(int row, int col)
    {
        if (_dependents.TryGetValue((row, col), out var set))
            return set;
        return new HashSet<(int, int)>();
    }

    /// <summary>
    /// Build the dependency graph from all formula cells in the worksheet.
    /// </summary>
    public void Build(XlWorksheet sheet)
    {
        _dependents.Clear();
        _precedents.Clear();
        _formulaCells.Clear();

        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            for (int c = 0; c < row.Cells.Count; c++)
            {
                var cell = row.Cells[c];
                if (!cell.HasFormula) continue;

                var pos = (r, c);
                _formulaCells.Add(pos);

                var ast = FormulaParser.Parse(cell.Formula!);
                var refs = CollectCellRefs(ast);

                if (!_precedents.ContainsKey(pos))
                    _precedents[pos] = new List<(int, int)>();

                foreach (var refPos in refs)
                {
                    _precedents[pos].Add(refPos);

                    if (!_dependents.ContainsKey(refPos))
                        _dependents[refPos] = new HashSet<(int, int)>();
                    _dependents[refPos].Add(pos);
                }
            }
        }
    }

    /// <summary>
    /// Topological sort of all formula cells. Returns cells in evaluation order,
    /// or null if a circular reference is detected.
    /// </summary>
    public List<(int Row, int Col)>? TopologicalSort()
    {
        // Build set of formula cell positions for O(1) lookup
        var formulaSet = new HashSet<(int, int)>(_formulaCells);

        // Kahn's algorithm
        // In-degree = number of unresolved formula-cell precedents.
        // Non-formula precedents (plain value cells) don't need evaluation — skip them.
        var inDegree = new Dictionary<(int, int), int>();
        foreach (var pos in _formulaCells)
        {
            var count = 0;
            if (_precedents.TryGetValue(pos, out var precs))
            {
                foreach (var p in precs)
                    if (formulaSet.Contains(p))
                        count++;
            }
            inDegree[pos] = count;
        }

        // Queue starts with cells that have no unresolved precedents
        var queue = new Queue<(int, int)>();
        foreach (var (pos, degree) in inDegree)
        {
            if (degree == 0)
                queue.Enqueue(pos);
        }

        var sorted = new List<(int Row, int Col)>();

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            sorted.Add(current);

            // For each cell that depends on current, reduce its in-degree
            if (_dependents.TryGetValue(current, out var deps))
            {
                foreach (var dep in deps)
                {
                    if (inDegree.ContainsKey(dep))
                    {
                        inDegree[dep]--;
                        if (inDegree[dep] == 0)
                            queue.Enqueue(dep);
                    }
                }
            }
        }

        // If sorted count < formula cell count, there's a cycle
        if (sorted.Count < _formulaCells.Count)
        {
            return null; // circular reference
        }

        return sorted;
    }

    /// <summary>
    /// Detect circular references by running Kahn's algorithm and collecting
    /// cells that remain in the graph (in-degree never reaches 0).
    /// </summary>
    public List<List<(int Row, int Col)>> DetectCircularReferences()
    {
        var formulaSet = new HashSet<(int, int)>(_formulaCells);

        var inDegree = new Dictionary<(int, int), int>();
        foreach (var pos in _formulaCells)
        {
            var count = 0;
            if (_precedents.TryGetValue(pos, out var precs))
            {
                foreach (var p in precs)
                    if (formulaSet.Contains(p))
                        count++;
            }
            inDegree[pos] = count;
        }

        var queue = new Queue<(int, int)>();
        foreach (var (pos, degree) in inDegree)
        {
            if (degree == 0)
                queue.Enqueue(pos);
        }

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            if (_dependents.TryGetValue(current, out var deps))
            {
                foreach (var dep in deps)
                {
                    if (inDegree.ContainsKey(dep))
                    {
                        inDegree[dep]--;
                        if (inDegree[dep] == 0)
                            queue.Enqueue(dep);
                    }
                }
            }
        }

        // Remaining cells with in-degree > 0 are part of cycles
        var remaining = new HashSet<(int, int)>();
        foreach (var (pos, degree) in inDegree)
        {
            if (degree > 0)
                remaining.Add(pos);
        }

        var cycles = new List<List<(int, int)>>();
        while (remaining.Count > 0)
        {
            var start = remaining.First();
            var path = new List<(int, int)>();
            var visited = new HashSet<(int, int)>();
            if (FindCycle(start, start, path, visited, remaining))
            {
                cycles.Add(path);
                foreach (var p in path)
                    remaining.Remove(p);
            }
            else
            {
                remaining.Remove(start);
            }
        }

        return cycles;
    }

    private bool FindCycle((int Row, int Col) current, (int Row, int Col) target,
                            List<(int, int)> path, HashSet<(int, int)> visited,
                            HashSet<(int, int)> remaining)
    {
        if (visited.Contains(current)) return false;
        visited.Add(current);
        path.Add(current);

        if (_dependents.TryGetValue(current, out var deps))
        {
            foreach (var dep in deps)
            {
                if (dep.Equals(target))
                {
                    path.Add(target);
                    return true;
                }
                if (remaining.Contains(dep) && FindCycle(dep, target, path, visited, remaining))
                    return true;
            }
        }

        path.RemoveAt(path.Count - 1);
        return false;
    }

    // ── Helpers ──

    private static HashSet<(int, int)> CollectCellRefs(Expr expr)
    {
        var refs = new HashSet<(int, int)>();
        CollectCellRefsRecursive(expr, refs);
        return refs;
    }

    private static void CollectCellRefsRecursive(Expr expr, HashSet<(int, int)> refs)
    {
        switch (expr)
        {
            case CellRefExpr c:
                // Only track same-sheet references for now (cross-sheet handled at workbook level)
                if (c.Sheet is null)
                    refs.Add((c.Row, c.Col));
                break;

            case RangeRefExpr r:
                if (r.Sheet is not null) break;
                int startRow = r.StartRow == int.MinValue ? 0 : Math.Max(0, r.StartRow);
                int endRow = r.EndRow == int.MaxValue ? int.MaxValue : r.EndRow;
                if (endRow == int.MaxValue) break; // unbounded range, skip
                int startCol = Math.Max(0, r.StartCol);
                int endCol = Math.Max(0, r.EndCol);
                // Limit to reasonable range to avoid memory issues from full-column refs
                int rowLimit = Math.Min(endRow, startRow + 5000);
                int colLimit = Math.Min(endCol, startCol + 5000);
                for (int row = startRow; row <= rowLimit; row++)
                    for (int col = startCol; col <= colLimit; col++)
                        refs.Add((row, col));
                break;

            case BinaryExpr b:
                CollectCellRefsRecursive(b.Left, refs);
                CollectCellRefsRecursive(b.Right, refs);
                break;

            case UnaryExpr u:
                CollectCellRefsRecursive(u.Operand, refs);
                break;

            case FunctionExpr f:
                foreach (var arg in f.Arguments)
                    CollectCellRefsRecursive(arg, refs);
                break;
        }
    }
}
