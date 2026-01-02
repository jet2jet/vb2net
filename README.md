# vb2net

The helper module `vb2net` for Visual Basic for Applications (VBA) 7.0 and support DLL, providing access to .NET (dotnet; formerly .NET Core) assemblies and classes.

## Requirements

- Visual Basic for Application 7.0 (included in Microsoft Office 2010 or higher)
  - ***(not tested)*** To use on Visual Basic 6.0, rewrite `LongPtr` to `Long` and remove all `PtrSafe` specifiers.
- .NET 8.0 or higher
  - The runtimeconfig specifies .NET version 8, so if you use version 9 (or higher), please modify `MakeTempRuntimeConfig` function in [vb2net.bas](./vb2net.bas).
  - Starting from v0.2.0.0, .NET 8.0 or higher is required to use functions for loading an assembly from the binary.

## Usage

1. Import [vb2net.bas](./vb2net.bas) and [ExitHandler.bas](./ExitHandler.bas) into your VB/VBA project
2. Write your code with calling `InitializeVb2net` procedure
  - After initialization, `LoadAssembly` and `LoadAssemblyFromFile` can be used.
  - The sample is in vb2net.bas as `Sample` procedure.

**NOTE: After v0.2.0.0, you don't need to put `vb2net.dll` file since the binary is bundled in vb2net.bas.**

## Build `vb2net.dll`

> Visual Studio 2022 is required to open the solution file [vb2net/vb2net.sln](./vb2net/vb2net.sln).

Build [vb2net/vb2net.csproj](./vb2net/vb2net.csproj) by `dotnet build vb2net.csproj -p Release`, or with Visual Studio or related build tools.
