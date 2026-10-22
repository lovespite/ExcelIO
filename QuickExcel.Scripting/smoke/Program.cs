using System.Diagnostics;
using QuickExcel.Scripting;

static void Assert(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

static void Fails(Action action, string expected)
{
    try { action(); }
    catch (InvalidOperationException error) when (error.Message.Contains(expected, StringComparison.OrdinalIgnoreCase)) { return; }
    throw new Exception("Expected JavaScript error containing: " + expected);
}

using var session = new QuickJsSession("""
    console.log('QuickJS ready');
    quickExcel.functions.register('discount', (price, rate) => price * (1 - rate));
    quickExcel.functions.register('RANGESUM', values => values.reduce((sum, value) => sum + value, 0));
    quickExcel.functions.register('TEXT', value => '中文 ' + value);
    quickExcel.functions.register('BOOL', value => value > 0);
    quickExcel.functions.register('EMPTY', () => null);
    quickExcel.functions.register('BAD', () => ({}));
    quickExcel.functions.register('PROMISE', () => Promise.resolve(42));
    quickExcel.functions.register('LOOP', () => { while (true) {} });
    quickExcel.functions.register('RECURSE', () => { function recurse() { return recurse(); } return recurse(); });
    quickExcel.functions.register('ENV', () => typeof std + ',' + typeof os + ',' + typeof require + ',' + typeof fetch);
    """);

Assert(session.FunctionNames.Contains("DISCOUNT"), "Case normalization");
Assert(Equals(session.Invoke("DISCOUNT", [100d, 0.2d]), 80d), "Discount");
Assert(Equals(session.Invoke("RANGESUM", [new List<object> { 1d, 2d, 3d }]), 6d), "Ranges");
Assert(Equals(session.Invoke("TEXT", ["测试"]), "中文 测试"), "UTF-8");
Assert(Equals(session.Invoke("BOOL", [1d]), true), "Boolean");
Assert(session.Invoke("EMPTY", []) is null, "Null");
Assert(Equals(session.Invoke("ENV", []), "undefined,undefined,undefined,undefined"), "No host APIs");
Assert(session.GetLogs().Single() == "QuickJS ready", "Console logs");
Fails(() => session.Invoke("BAD", []), "objects and promises");
Fails(() => session.Invoke("PROMISE", []), "objects and promises");
Fails(() => session.Invoke("RECURSE", []), "call stack size");
var timer = Stopwatch.StartNew();
Fails(() => session.Invoke("LOOP", []), "time limit");
Assert(timer.Elapsed < TimeSpan.FromSeconds(3), "Bounded interrupt");
Assert(Equals(session.Invoke("DISCOUNT", [100d, 0.2d]), 80d), "Recovery after interrupt");
await Task.Run(() => Assert(Equals(session.Invoke("DISCOUNT", [100d, 0.2d]), 80d), "Thread switch"));
Fails(() => { using var invalid = new QuickJsSession("quickExcel.functions.register('BAD_NAME', () => 1)"); }, "A-Z");
Fails(() => { using var invalid = new QuickJsSession("while (true) {}"); }, "time limit");
Fails(() => { using var invalid = new QuickJsSession("const huge = new ArrayBuffer(80 * 1024 * 1024);"); }, "out of memory");

timer.Restart();
for (var index = 0; index < 10000; index++) session.Invoke("DISCOUNT", [100d, 0.2d]);
Console.WriteLine($"PASS: QuickJS runtime smoke checks; 10,000 calls: {timer.ElapsedMilliseconds} ms.");
