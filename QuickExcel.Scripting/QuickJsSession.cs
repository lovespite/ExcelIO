using System.Runtime.InteropServices;
using System.Text.Json;

namespace QuickExcel.Scripting;

/// <summary>A serialized, memory-limited QuickJS session for workbook functions.</summary>
public sealed class QuickJsSession : IDisposable
{
    private const string NativeLibrary = "quickexcel-js";
    private readonly object _gate = new();
    private IntPtr _session;

    // The bridge object is held by native code, not exposed on globalThis. Workbook
    // code receives only registration and console APIs, with no host CLR objects.
    private const string Bootstrap = """
        (() => {
            const registry = Object.create(null);
            const ownKeys = Object.keys;
            const freeze = Object.freeze;
            const makeFunction = Function;
            const upper = Function.prototype.call.bind(String.prototype.toUpperCase);
            const validName = Function.prototype.call.bind(RegExp.prototype.test, /^[A-Z]+$/);
            const isFiniteNumber = Number.isFinite;
            const TypeErrorCtor = TypeError;
            let loaded = false;
            const logs = [];
            const logger = (...values) => {
                if (logs.length < 100) logs.push(values.map(String).join(' ').slice(0, 4000));
            };
            const console = freeze({ log: logger, info: logger, warn: logger, error: logger });
            const quickExcel = freeze({ functions: freeze({
                register(name, fn) {
                    if (loaded) throw new TypeErrorCtor('Functions can only be registered while loading the script.');
                    if (typeof name !== 'string') throw new TypeErrorCtor('Function name must be a string.');
                    name = upper(name);
                    if (!validName(name)) throw new TypeErrorCtor('Function names must contain letters A-Z only.');
                    if (typeof fn !== 'function') throw new TypeErrorCtor('A registered formula must be a function.');
                    if (registry[name]) throw new TypeErrorCtor('Duplicate function: ' + name);
                    registry[name] = fn;
                }
            }) });
            return freeze({
                load(source) {
                    if (loaded) throw new TypeErrorCtor('Script has already been loaded.');
                    makeFunction('quickExcel', 'console', '"use strict";\n' + source)(quickExcel, console);
                    loaded = true;
                    return ownKeys(registry);
                },
                invoke(name, args) {
                    const fn = registry[name];
                    if (!fn) throw new TypeErrorCtor('Unknown function: ' + name);
                    const value = fn(...args);
                    if (value === null || value === undefined) return null;
                    if (typeof value === 'string' || typeof value === 'boolean') return value;
                    if (typeof value === 'number' && isFiniteNumber(value)) return value;
                    throw new TypeErrorCtor('Formula functions must return a finite number, string, boolean, or null; objects and promises are not supported.');
                },
                logs() { return logs.slice(); }
            });
        })()
        """;

    public IReadOnlyList<string> FunctionNames { get; }

    public QuickJsSession(string source)
    {
        ArgumentNullException.ThrowIfNull(source);
        if (!OperatingSystem.IsWindows() || !Environment.Is64BitProcess)
            throw new PlatformNotSupportedException("QuickExcel currently includes a Windows x64 QuickJS runtime.");
        _session = NativeCreate(Bootstrap, out var error);
        if (_session == IntPtr.Zero) throw ReadException(error);
        try
        {
            var result = Call("load", new object?[] { source });
            FunctionNames = Array.AsReadOnly(JsonSerializer.Deserialize<string[]>(result) ?? []);
        }
        catch
        {
            Dispose();
            throw;
        }
    }

    /// <summary>Calls a registered formula with scalar or array arguments.</summary>
    public object? Invoke(string functionName, IReadOnlyList<object> arguments)
    {
        ArgumentNullException.ThrowIfNull(functionName);
        ArgumentNullException.ThrowIfNull(arguments);
        lock (_gate)
        {
            string result = Call("invoke", new object?[] { functionName.ToUpperInvariant(), arguments });
            using var document = JsonDocument.Parse(result);
            var value = document.RootElement;
            return value.ValueKind switch
            {
                JsonValueKind.Null => null,
                JsonValueKind.String => value.GetString(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Number => value.GetDouble(),
                _ => throw new InvalidOperationException("JavaScript returned an unsupported formula value.")
            };
        }
    }

    public IReadOnlyList<string> GetLogs()
    {
        lock (_gate)
            return JsonSerializer.Deserialize<string[]>(Call("logs", [])) ?? [];
    }

    private string Call(string method, object?[] arguments)
    {
        ObjectDisposedException.ThrowIf(_session == IntPtr.Zero, this);
        var json = JsonSerializer.Serialize(arguments);
        if (json.Length > 4 * 1024 * 1024)
            throw new InvalidOperationException("JavaScript arguments exceed the 4 MiB input limit.");
        var result = NativeCall(_session, method, json, out var error);
        if (result == IntPtr.Zero) throw ReadException(error);
        try { return Marshal.PtrToStringUTF8(result) ?? "null"; }
        finally { NativeFreeString(result); }
    }

    private static InvalidOperationException ReadException(IntPtr error)
    {
        try { return new InvalidOperationException(Marshal.PtrToStringUTF8(error) ?? "QuickJS execution failed."); }
        finally { if (error != IntPtr.Zero) NativeFreeString(error); }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_session == IntPtr.Zero) return;
            NativeDestroy(_session);
            _session = IntPtr.Zero;
        }
        GC.SuppressFinalize(this);
    }

    ~QuickJsSession() => Dispose();

    [DllImport(NativeLibrary, EntryPoint = "qx_create", CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr NativeCreate([MarshalAs(UnmanagedType.LPUTF8Str)] string bootstrap, out IntPtr error);
    [DllImport(NativeLibrary, EntryPoint = "qx_call", CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr NativeCall(IntPtr session, [MarshalAs(UnmanagedType.LPUTF8Str)] string method, [MarshalAs(UnmanagedType.LPUTF8Str)] string argumentsJson, out IntPtr error);
    [DllImport(NativeLibrary, EntryPoint = "qx_destroy", CallingConvention = CallingConvention.Cdecl)]
    private static extern void NativeDestroy(IntPtr session);
    [DllImport(NativeLibrary, EntryPoint = "qx_free_string", CallingConvention = CallingConvention.Cdecl)]
    private static extern void NativeFreeString(IntPtr text);
}
