@echo off

@where nuget >NUL 2>&1
if ERRORLEVEL 1 goto MissingNuGet

if not exist out md out

del out\*.nupkg

NuGet.exe pack WindowsAdminCenter.nuspec -NoPackageAnalysis -OutputDirectory out

exit /b 0

:MissingNuGet
echo NuGet.exe not found in your PATH.
echo Download: https://nuget.org/nuget.exe

exit /b 1


