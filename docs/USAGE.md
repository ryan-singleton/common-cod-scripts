# Common Fabrication Scripts — Usage Guide

This guide covers the key functions available after adding `Include MAPPATH_SCRIPTS + "Standard.cod"` to your script. It is organized by the kind of task you are trying to accomplish.

---

## Setup

### Install

[Git clone](https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository) from this repo, or download the zip.

Copy "src/Scripts/Common" into your Fabrication scripts folder.

### Include the scripts

Every script that uses this library starts with one line:

```vb
Include MAPPATH_SCRIPTS + "Standard.cod"
```

This single include pulls in the entire chain:

```
Standard.cod
  └── FabFunctions.cod
        └── CoreFunctions.cod
              └── CommonStrings.cod
  └── ContextFunctions.cod
        └── ShellFunctions.cod
              └── CommonPaths.cod
                    └── CoreFunctions.cod
```

The `Common*.cod` files (CutTypes, Dimensions, Options) are also available because `FabFunctions.cod` pulls in what it needs. If you want only a subset of the library, you can include a specific file lower in the chain instead.

> NOTE: If Script A includes the standard.cod file, Script A can call Script B and Script B will already be able to use the included functions. However, if Script B is also called by Script C, and Script C does *not* include the standard.cod, Script B will *not* have access to the standard functions. In the case of high potential reuse, it is a good idea to re-include the standard.cod for scripts like Script B.

---

## Comparing Item Properties

The primary tool for checking item properties without magic strings or nested `Wildcard()` calls.

### Checking a single value

```vb
REM Does the item's service name contain "weld"?
If Contains(item.Service, "weld") Then ...

REM Does the filename start with "square"?
If StartsWith(item.Filename, "square") Then ...

REM Is the CutType exactly "Decoiled Straight"?
If Equal(item.CutType, DecoiledStraight()) Then ...

REM Is the CutType NOT machine cut?
If NotEqual(item.CutType, MachineCut()) Then ...
```

### Checking against a list

```vb
REM Does the CID match any of several pattern numbers?
Dim patterns As ARRAY
patterns.Add(1, 866, 900)
If EqualsAny(item.CID, patterns) Then ...

REM Does the service contain any of several keywords?
Dim keywords As ARRAY
keywords.Add("weld", "plasma", "laser")
If ContainsAny(item.Service, keywords) Then ...
```

### Combining conditions

```vb
If EqualsAny(item.CID, patterns) And
    Not Contains(item.Service, "weld") And
    StartsWith(item.Filename, "square") And
    Contains(item.Material, "galv") Then
    ...
End If
```

---

## Working with Dimensions

Use `CommonDimensions.cod` functions instead of writing dimension name strings by hand.

```vb
REM Read a dimension value
Dim len = item.Dim[Length()].Value

REM Set a dimension to Auto
item.Dim[Length()].Value = Auto()

REM Check if a dimension exists on the item
If HasDimensionName(item, Width()) Then ...

REM Collect all dimensions whose name contains "Width"
Dim widthDims = MatchingDimensionNames(item, Width())
```

Available dimension name functions:

| Function | Returns |
|---|---|
| `Width()` | `"Width"` |
| `Depth()` | `"Depth"` |
| `Height()` | `"Height"` |
| `Length()` | `"Length"` |
| `Diameter()` | `"Diameter"` |
| `Angle()` | `"Angle"` |
| `TopWidth()` / `BottomWidth()` / `LeftWidth()` / `RightWidth()` | Side-specific widths |
| `WidthIn()` / `WidthOut()` | Inlet/outlet widths |
| `TopDepth()` / `BottomDepth()` / `LeftDepth()` / `RightDepth()` | Side-specific depths |
| `DepthIn()` / `DepthOut()` | Inlet/outlet depths |
| `TopExtension()` / `BottomExtension()` / `LeftExtension()` / `RightExtension()` | Extensions |
| `InnerRadius()` | `"Inner Radius"` |

---

## Changing Connectors, Seams, and Airturns

Use the `ChangeItem*` family of functions. They handle unlocking, updating, and re-locking for you.

```vb
REM Change connector at index 1
ChangeItemConnector(item, 1, "TDC")

REM Change all connectors in a loop
Dim i
For i=1 To item.Connectors
    If Not Contains(item.Connector[i].Value, "TDC") Then
        ChangeItemConnector(item, i, "TDC")
    End If
Next

REM Change a seam
ChangeItemSeam(item, 1, "Pittsburgh")

REM Change an airturn
ChangeItemAirturn(item, 1, "Standard")

REM Change a splitter
ChangeItemSplitter(item, 1, "Fixed")
```

If you have a direct reference rather than an index, the lower-level functions work too:

```vb
ChangeConnector(item.Connector[1], "TDC")
ChangeSeam(item.Seam[1], "Pittsburgh")
```

### Finding connectors or seams by value

```vb
REM Get all connectors whose value contains "TDC"
Dim tdcConnectors = MatchingConnectors(item, "TDC")

REM Get all seams whose value contains "Pittsburgh"
Dim pittSeams = MatchingSeams(item, "Pittsburgh")
```

---

## Working with Item Options

Reference options by name using `CommonOptions.cod` functions.

```vb
REM Check the throat type
If Equal(item.option[ThroatType()].Value, Radius()) Then ...
If Equal(item.option[ThroatType()].Value, Mitred()) Then ...

REM American English alias — same result
If Equal(item.option[ThroatType()].Value, Mitered()) Then ...

REM Set gore count on an elbow (also available as SetGores)
SetGores(item, 5)

REM Or directly via the option name
item.option[NumberOfSegments()].Value = 5
item.Update()
```

---

## Changing Gauge and Specification

```vb
REM Change sheet gauge
ChangeGauge(item, 26)

REM Change wire gauge
ChangeWireGauge(item, 18)

REM Change specification
ChangeSpecification(item, "HVAC: Galvanized 2in")
```

---

## Cut Type Comparisons

```vb
If Equal(item.CutType, DecoiledStraight()) Then ...
If Equal(item.CutType, MachineCut()) Then ...
If Equal(item.CutType, DrawOnly()) Then ...
If Equal(item.CutType, Equipment()) Then ...
If Equal(item.CutType, SpiralStraight()) Then ...
If Equal(item.CutType, OvalStraight()) Then ...
If Equal(item.CutType, Pipework()) Then ...
```

---

## Adding Notes to Items

```vb
REM Prepend a note only if not already present
MaybePrependNote(item, "EXPOSED", EmptyString())

REM Append a note only if not already present
MaybeAppendNote(item, "ADD EC C2", "-")
```

---

## String Utilities

```vb
REM Join an array of strings
Dim parts As ARRAY
parts.Add("A", "B", "C")
Dim result = StringJoin(parts, ", ")   REM "A, B, C"

REM Slice a string
TakeUntil("ABCDEF", "CD")    REM "AB"
TakeAfter("ABCDEF", "CD")    REM "EF"

REM Format a decimal inch as a readable fraction
FormatInchFraction(2.4489, 16)    REM "2 7/16"
FormatInchFraction(0.5, 2)        REM "1/2"

REM Safe special characters
NewLine()       REM line feed (ASCII 10)
Tab()           REM tab (ASCII 9)
Tabs(3)         REM three tabs
Quotes()        REM double-quote character
FalseQuotes()   REM two apostrophes, visual inch mark
Star()          REM asterisk without wildcard behavior
None()          REM the sentinel string "None"
```

---

## Context Object

The context object is an optional but recommended tool for any multi-step routine. Create it once at the very top of your process script.

```vb
Dim context = CreateContext("MyProcessName")
```

This captures the current username, a readable timestamp, and a filename-safe timestamp at the moment the routine starts.

### Reading from context

```vb
GetContextName(context)      REM "MyProcessName"
GetSubContext(context)       REM "No SubContext" until you set it
GetUserName(context)         REM Windows username of the active user
GetTimestamp(context)        REM e.g. "2024-03-15 09:42:01"
GetFileTimestamp(context)    REM e.g. "20240315-094201"
```

### Tracking progress through stages

Call `SetSubContext` as your routine moves through distinct phases. This makes any logs you write far easier to read.

```vb
Dim context = CreateContext("AutoConnector")

SetSubContext(context, "Filtering")
REM ... filter items ...

SetSubContext(context, "ApplyingConnectors")
REM ... change connectors ...

SetSubContext(context, "Complete")
```

### Using standalone runtime functions

If you do not need a context object, the underlying functions are available directly:

```vb
UserName()        REM current Windows username
Timestamp()       REM current readable timestamp
FileTimestamp()   REM current filename-safe timestamp
```

### Extending context with custom state

The context is just an array, so you can attach whatever you need. The README shows a logging pattern as an example — context[2] and beyond are yours to use.

```vb
Dim context = CreateContext("MyProcess")
Dim logs As ARRAY
context.Add(logs)
```

---

## Path Utilities

```vb
REM Build a path under the Scripts folder
ScriptPath("Custom\CommonConnectors.cod")

REM Build a path under the VBS scripts folder
VbScriptPath("EnsurePath.vbs")

REM Build a path under the temp folder
TempRoot()                          REM "C:\Windows\Temp\FabScripts\"
TempPath("DynamicFunctions.cod")    REM temp folder + filename

REM Strip a leading slash before concatenating
TrimStartSlashes("\MyFolder\File.cod")    REM "MyFolder\File.cod"
```

---

## Shell / File Operations

```vb
REM Open a file or folder with its default application
RunPath("C:\Windows\Temp\FabScripts\report.txt")

REM Delete a file
DeletePath("C:\Windows\Temp\FabScripts\old.txt")

REM Ensure a directory exists (creates it if absent)
EnsurePath(TempRoot())

REM Run a VBS file with arguments
RunVbScript(VbScriptPath("EnsurePath.vbs"), "C:\SomePath False")
```

---

## None / Null Sentinel

Throughout the library, `"None"` is used as a conventional absent-value marker.

```vb
If IsNone(someValue) Then ...
If NotNone(someValue) Then ...

REM Produce the sentinel directly
Dim val = None()
```

---

## Quick Reference: Include Chain

| You want... | Minimum include |
|---|---|
| Everything | `Standard.cod` |
| Fab + string utilities only | `Common\FabFunctions.cod` |
| String/array/math utilities only | `Common\CoreFunctions.cod` |
| Path helpers only | `Common\CommonPaths.cod` |
| Shell operations | `Common\ShellFunctions.cod` |
| Context object | `Common\ContextFunctions.cod` (also pulls in Shell + Paths) |