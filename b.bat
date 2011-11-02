@echo off

%WINDIR%\Microsoft.NET\Framework\v3.5\MSBuild.exe %* /p:Platform=x86 /p:Configuration=Release
%WINDIR%\Microsoft.NET\Framework\v3.5\MSBuild.exe %* /p:Platform=x64 /p:Configuration=Release

RMDIR /S /Q ship
MKDIR ship
COPY x64\Release\*.msi ship
COPY x86\Release\*.msi ship

"C:\Program Files\Microsoft SDKs\Windows\v6.0A\Bin\signtool.exe" ^
sign ^
/d "Find Visio Command Addin" ^
/f "D:\Documents\code\bnk.pfx" ^
/p Ybrjkfq22 ^
/t http://timestamp.comodoca.com/authenticode ^
"ship\*.msi" "ship\*.exe"
