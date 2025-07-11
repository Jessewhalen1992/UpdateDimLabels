# UpdateDimLabels Plugin

## Quick Build

```bash
dotnet build -c Release            # builds net48 + net8.0-windows
# The ready-to-use plugin appears in: UpdateDimLabels.bundle
# Fody/Costura merges EPPlus so only `UpdateDimLabels.dll` is needed
# Copy that folder to C:\ProgramData\Autodesk\ApplicationPlugins
```

This AutoCAD Map 3D plug‑in adds the `UPDDIM` command which updates an aligned
 dimension using Object‑Data from a selected polyline. The command becomes
 available after loading the compiled DLL with `NETLOAD`.

## Building

1. Open `UpdateDimLabels.sln` in Visual Studio 2022 or newer.
2. Build the **Release|x64** configuration.
   Fody runs automatically and merges all required DLLs into
   `UpdateDimLabels.dll` in the output folder (`bin\x64\Release`).
3. Copy `CompanyLookup.xlsx` and `PurposeLookup.xlsx` to the output
   folder next to `UpdateDimLabels.dll`.

## Loading in AutoCAD Map 3D

1. Start AutoCAD Map 3D (the plug‑in relies on the Map APIs).
2. Run the `NETLOAD` command and browse to `UpdateDimLabels.dll` in the
   build output directory.
3. After loading, a message *"UpdateDimLabels loaded. Run UPDDIM to
   update dimensions."* appears on the command line.
4. Run the `UPDDIM` command to update dimension labels.
