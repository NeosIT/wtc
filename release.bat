REM Execute Visual Studio build / msbuild
REM Download ILMerge from https://www.nuget.org/packages/ilmerge and copy the file into this directory
ILMerge.exe /ndebug /targetplatform:4.0 /out:wtc.exe wtc\bin\Release\WTC.exe wtc\bin\Release\CommandLine.dll