# vb2net

The helper module `vb2net` for Visual Basic for Applications (VBA) 7.0 and support DLL, providing access to .NET (formerly .NET Core) assemblies and classes.

## Requirements

- Visual Basic for Application 7.0 (included in Microsoft Office 2010 or higher)
  - ***(not tested)*** To use on Visual Basic 6.0, rewrite `LongPtr` to `Long` and remove all `PtrSafe` specifiers.
- .NET 6.0 or higher
  - The file `vb2net.runtimeconfig.json` specifies .NET version 6, so if you use version 7 (or higher), please modify `vb2net.runtimeconfig.json`.

## Usage

1. Import [vb2net.bas](./vb2net.bas) and [ExitHandler.bas](./ExitHandler.bas) into your VB/VBA project
2. Download `vb2net.zip` from [Releases](https://github.com/jet2jet/vb2net/releases) and extract it (`vb2net.dll` and `vb2net.runtimeconfig.json` will be extracted)
3. Write your code with calling `InitializeVb2net` procedure
  - After initialization, `LoadAssembly` and `LoadAssemblyFromFile` can be used.
  - The sample is in vb2net.bas as `Sample` procedure.

## Build `vb2net.dll`

> Visual Studio 2022 is required to open the solution file [vb2net/vb2net.sln](./vb2net/vb2net.sln).

Build [vb2net/vb2net.csproj](./vb2net/vb2net.csproj) by `dotnet build vb2net.csproj -p Release`, or with Visual Studio or related build tools.
