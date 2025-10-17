## Building

### Dependencies

- `InsertIcons.exe` ([source]((https://github.com/einaregilsson/InsertIcons)) is required in the project's `tools` directory for embedding icon files in the EXE. You can compile it using Visual Studio or use their provided EXE (for my own use I compiled it).

- Windows API Code Pack DLLs. Already included in project's `libs\bin` directory.

	- Reference assemblies that expose jumplist related features to .NET programs. Once hosted at MSDN but now only available via mirrors.

	- [This](https://archive.org/details/windows-api-code-pack-self-extractor) archived mirror contains the original v1.1 signed Microsoft installer that includes the source code and pre-compiled DLLs used by this project, added to the project files for convenience due to being hard to find proper provenance of other unpacked mirrors.

- (Optional) [DarkModeForms fork](https://github.com/chocmake/Dark-Mode-Forms) that contains the custom style adjustments used for Jumplist Creator and configured to compile to DLL. Compile, rename to `Theme.dll` and include in Jumplist Creator's `libs\bin` directory. Provides a dark theme but build doesn't require it be present.

### Steps

> Below uses the CLI only build tool but you can also build using the Visual Studio Blend GUI if you prefer.

1. Open the Visual Studio 2022 build tools installer ([official direct link](https://aka.ms/vs/17/release/vs_buildtools.exe)) and tick to install the .NET build tools. Can alternatively use VS 2019's CLI [build tools](https://aka.ms/vs/16/release/vs_buildtools.exe).

2. Unzip the repo source code and add the `InsertIcons.exe` to the `tools` directory.
	- Optionally add `Theme.dll` to `libs\bin`.

3. Search the start menu for *Developer Command Prompt for VS 2022* (or *2019*) to launch the Visual Studio CLI build environment script.

4. Within it `cd` to the Jumplist Creator directory and run:

```
msbuild JumplistCreator.sln /m /p:Configuration=Release /p:Platform="Any CPU"
```

5. Find the output in the newly generated `dist` directory.