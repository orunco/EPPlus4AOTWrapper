@echo off

set dotnet="C:\Program Files\dotnet\dotnet.exe" 

%dotnet% publish --framework net8.0 -r win-x64 -c Release -p:PublishAot=true -p:PublishTrimmed=true -p:NativeLib=Shared

copy /Y bin\Release\net8.0\win-x64\publish\EPPlus4.dll bin\Release\net8.0\win-x64\publish\EPPlus4.Native-win-x64.dll

copy /Y bin\Release\net8.0\win-x64\publish\EPPlus4.dll ..\EPPlusForNet\EPPlus4.Native-win-x64.dll
copy /Y bin\Release\net8.0\win-x64\publish\EPPlus4.pdb ..\EPPlusForNet\EPPlus4.Native-win-x64.pdb