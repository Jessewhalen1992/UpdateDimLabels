# UpdateDimLabels Plugin

This AutoCAD Map 3D plug‑in adds the `UPDDIM` command which updates an aligned
 dimension using Object‑Data from a selected polyline. The command becomes
 available after loading the compiled DLL with `NETLOAD`.

## Building

1. Open `UpdateDimLabels.sln` in Visual Studio 2022 or newer.
2. Build the **Release|x64** configuration.
   All required dependency DLLs will be copied to the output folder
   (`bin\x64\Release`).
3. Copy `CompanyLookup.xlsx` and `PurposeLookup.xlsx` to the output
   folder next to `UpdateDimLabels.dll`.

## Loading in AutoCAD Map 3D

1. Start AutoCAD Map 3D (the plug‑in relies on the Map APIs).
2. Run the `NETLOAD` command and browse to `UpdateDimLabels.dll` in the
   build output directory.
3. After loading, a message *"UpdateDimLabels loaded. Run UPDDIM to
   update dimensions."* appears on the command line.
4. Run the `UPDDIM` command to update dimension labels.
