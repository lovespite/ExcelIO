# QuickJS runtime provenance

QuickExcel embeds **QuickJS-NG 0.17.0**, a maintained fork of Fabrice Bellard and
Charlie Gordon's QuickJS JavaScript engine. It does not use V8 or a managed JS
interpreter. The source is MIT licensed; the complete upstream license is in
`QUICKJS-LICENSE.txt` and is copied with the application.

- Upstream: https://github.com/quickjs-ng/quickjs
- Pinned release: https://github.com/quickjs-ng/quickjs/releases/tag/v0.17.0
- Source archive: https://github.com/quickjs-ng/quickjs/archive/refs/tags/v0.17.0.zip
- Archive SHA-256: `9DE1BF573606D2AEBBBAB6217E616A66E09E3613518489B37B9B3157CF398DCC`

`build.ps1` verifies the archive hash before extracting and compiles the engine
from source using CMake and the installed Visual Studio C++ compiler. Source
downloads and generated binaries are ignored by Git. No downloaded prebuilt
executable or binary dependency is used. A normal project build runs the script
only if the native DLL is missing; after editing native code run:

```powershell
.\QuickExcel.Scripting\native\build.ps1 -Rebuild
```

Requirements to build from a fresh checkout: Windows x64, .NET 10 SDK, CMake
supporting your Visual Studio version, Visual Studio 2022 or newer with C++ build
tools, and internet access for the first source download. The published desktop
app includes the built DLL and does not require these build tools.

The bridge compiles only the core engine, without `quickjs-libc`, OS modules,
module loaders, filesystem, network, or CLR bindings. Each session caps engine
allocation at 64 MiB and JS stack use at 256 KiB; each load and invocation has a
one second native interrupt deadline. Requests/results are limited to roughly
4 MiB. This is an in-process prototype runtime, not an operating system security
sandbox. Imported workbook scripts must still be explicitly enabled by the user.

Calls are serialized and `JS_UpdateStackTop` is called before execution, allowing
sequential access from different managed threads. Each document owns its own
runtime. The prototype accepts synchronous scalar formula results and array
arguments; promises and object results are rejected.

Run the runtime smoke checks with:

```powershell
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj
```
