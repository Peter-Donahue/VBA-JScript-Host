# JScript.bas

> A single-file VBA module that embeds the Windows **JScript9** engine directly inside any Office application — evaluate JavaScript expressions, call JS functions from VBA, and expose VBA objects to script, all without spawning a process or creating a `ScriptControl` reference.

> I was working on this several years ago when Wqweto published his version (https://github.com/wqweto/VBTixyLand) so I forgot about it for a while. If you are running this in VB6, rather than VBA, or if you will only need to run this on 32bit office, consider using his version instead.

```vba
Debug.Print Eval("1 + 2 + 3")                          ' → 6
Debug.Print Eval("Math.pow(2, 10)")                     ' → 1024
Debug.Print Eval("ThisWorkbook.Name")                   ' → MyBook.xlsm

Eval "function greet(name) { return 'Hello, ' + name + '!'; }"
Debug.Print Fn("greet")("World")                        ' → Hello, World!

Dim obj As Object
Set obj = JsonParse("[1, 2, 3, 4, 5]")
Debug.Print JsonStringify(obj)                          ' → [1,2,3,4,5]
```

---

## Why?

The venerable `MSScriptControl.ScriptControl` ActiveX works only in 32-bit hosts and has been removed from 64-bit Office. This module drives the same underlying **IActiveScript / IActiveScriptParse** COM interfaces directly via `DispCallFunc`, making it compatible with both 32-bit and 64-bit Office with no external dependencies and no registered type library.

---

## Features

| | |
|---|---|
| 🧩 **Zero dependencies** | No `References` to set, no registered type library, no COM interop assemblies |
| 🔢 **32-bit & 64-bit** | Conditional compilation handles pointer widths and API declarations automatically |
| 🚫 **Mac-safe** | The entire module is guarded by `#If Not Mac` — it compiles cleanly on macOS without doing anything |
| 📦 **JSON built-in** | Douglas Crockford's JSON2 polyfill is injected automatically at startup |
| 🔗 **Arrow predicates** | Lambda-style helpers: `Predicate("x => x * 2")(21)` → `42` |
| 📂 **Import from anywhere** | Load script from a string, a file path, or a URL |
| 🌍 **Named object bridge** | Expose any VBA object to the JS global scope with `AddNamedObject` |
| 📦 **`require()` shim** | Minimal CommonJS-style module system pre-installed in every session |

---

## Compatibility

| Environment | Supported |
|---|---|
| Excel / Word / Access (32-bit, Office 2007+) | ✅ |
| Excel / Word / Access (64-bit, Office 2010+) | ✅ |
| VBA6 (Office 2003 and earlier) | ✅ |
| VBA7 + Win32 build | ✅ |
| VBA7 + Win64 build | ✅ |
| macOS (any Office version) | 🚫 Windows-only engine — module is a no-op |

---

## Installation

1. Download [`JScript.bas`](./JScript.bas).
2. In the VBA Editor (**Alt + F11**), go to **File → Import File** and select `JScript.bas`.
3. That's it — no references, no setup.

---

## Quick Start

### Evaluate an expression

```vba
Dim result As Variant
result = Eval("2 ** 8")          ' 256
result = Eval("'hello'.toUpperCase()")  ' HELLO
```

### Execute statements

```vba
Eval "var counter = 0;"
Eval "function increment() { counter += 1; }"
Eval "increment(); increment();"
Debug.Print Eval("counter")     ' 2
```

### Call a named function

```vba
Eval "function add(a, b) { return a + b; }"

' Via Fn() — returns the JS function as a callable VBA Object
Debug.Print Fn("add")(3, 4)    ' 7
```

### Arrow-expression predicates

`Fn()` accepts an ES6-style arrow expression and wraps it in a function object:

```vba
Dim double As Object
Set double = Fn("x => x * 2")
Debug.Print double(21)          ' 42

' Useful for filtering / mapping with VBA collections
Dim isEven As Object
Set isEven = Fn("n => n % 2 === 0")
Debug.Print isEven(4)           ' True
Debug.Print isEven(7)           ' False
```

### JSON

```vba
' Parse JSON text into a live JS object
Dim arr As Object
Set arr = JsonParse("[10, 20, 30]")
Debug.Print arr(0)              ' 10

' Serialize any JS object back to a JSON string
Debug.Print JsonStringify(arr)  ' [10,20,30]
```

### Expose VBA objects to script

Application and ThisWorkbook are registered automatically. You can add any object:

```vba
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
dict("key") = "value"

AddNamedObject "MyDict", dict

Debug.Print Eval("MyDict('key')")   ' value
```

### Import from a file or URL

```vba
' From a file path
Import "C:\scripts\utils.js", ScriptSourceType.Path

' From a URL (requires ImplementRuntimeSourceResolveURL = True)
Import "https://example.com/lib.js", ScriptSourceType.URL
```

---

## Public API Reference

### Evaluation

| Function | Description |
|---|---|
| `Eval(code)` | Evaluate a JS expression or statement. Returns the result as a `Variant`. |
| `Import(source, [type])` | Load and evaluate script from a `String`, file `Path`, or `URL`. |
| `Fn(nameOrArrow)` | Return a named JS function object, or compile an arrow expression into one. |

### JSON

| Function | Description |
|---|---|
| `JsonParse(jsonText)` | Parse a JSON string into a JS object. |
| `JsonStringify(obj)` | Serialize a JS object to a JSON string. |

> Both are `Property Get` members so you can use them as function calls or assignments naturally in VBA.

### Engine Lifecycle

| Member | Description |
|---|---|
| `AddNamedObject(name, obj, [flags])` | Expose a VBA object as a global in the JS namespace. |
| `RunState` (get/let) | Get or set the engine's `ScriptState` (`Connected`, `Disconnected`, etc.). |
| `CloseScriptHost()` | Gracefully shut down the engine. Call from `Workbook_BeforeClose`. |

### Lower-Level IActiveScript Wrappers

These are exposed `Public` for advanced use but are not needed for typical scripting:

`IActiveScript_GetScriptDispatch`, `IActiveScript_AddTypeLib`, `IActiveScript_Clone`, `IActiveScript_GetCurrentScriptThreadID`, `IActiveScript_GetScriptThreadID`, `IActiveScript_GetScriptThreadState`, `IActiveScript_InterruptScriptThread`, `IActiveScriptParse_AddScriptlet`

---

## How It Works

VBA has no native way to implement a COM interface, but `DispCallFunc` (from `oleaut32`) can call any function at an arbitrary vtable offset on any pointer. This module:

1. **Allocates a block of `CoTaskMem`** and writes function pointers into it in the exact layout the COM ABI expects for three interfaces: `IActiveScriptSite`, `IActiveScriptSiteDebug`, and `IActiveScriptSiteWindow`.
2. **Passes `VarPtr(Host.Site)`** to `SetScriptSite` as the COM object pointer. Dereferencing it once gives the vtable address; dereferencing again gives individual method pointers — exactly what the engine expects.
3. **Persists the vtable block** address as a window property on `Application.HWnd` so it survives VBA recompiles without re-allocating.

This technique is explained in depth by Matt Curland in *Advanced Visual Basic 6*, and the vtable-dispatch pattern is adapted from LaVolpe's FauxInterface work on VBForums.

---

## Caveats & Known Limitations

- **Windows only.** JScript9 (`jscript9.dll`) is a Windows component. The module compiles on macOS but all procedures are dead code.
- **JScript, not V8.** This is Microsoft's JScript9 engine (ES5 with some ES6), not Node.js or a modern browser engine. ES6+ features like `class`, `let`, `const`, arrow functions in older IE-era syntax, and `Promise` may or may not be available depending on the Windows version.
- **Single-threaded.** The scripting engine runs on the VBA thread. Long-running scripts will block the UI.
- **Memory.** The vtable block is intentionally never freed — it is reclaimed when the host process exits. This is by design: the block is tiny (~108 bytes on x64) and reused across recompiles.

---

## License

MIT — see [`LICENSE`](./LICENSE) for the full text.

---

## Acknowledgements

| Author | Contribution |
|---|---|
| [Matt Curland](https://tinyurl.com/y2mghb93) | *Advanced Visual Basic 6* — COM internals and vtable techniques |
| [Olaf Schmidt](https://tinyurl.com/y5v4a2yr) | vbFriendly Lightweight COM Interfaces |
| [LaVolpe](https://tinyurl.com/yxdxpe4o) | FauxInterface — VBForums |
| [David Zimmer](http://sandsprite.com) | VB-ized IActiveScript Type Library |
| [Douglas Crockford](https://github.com/douglascrockford/JSON-js) | JSON2 polyfill |
| [Vladimir Vissoultchev](https://github.com/wqweto) | File-reading utilities (VbPeg) |
| [@Greedo](https://github.com/Greedquest) | `LongPtr` Enum polyfill for VBA6 |
| [Ion Cristian Buse](https://github.com/cristianbuse/VBA-MemoryTools) | Cross-platform conditional compilation patterns |
