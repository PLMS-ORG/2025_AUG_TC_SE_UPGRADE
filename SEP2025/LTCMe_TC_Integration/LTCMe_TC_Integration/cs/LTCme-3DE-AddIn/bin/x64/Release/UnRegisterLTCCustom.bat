setlocal
@echo off

set Module_Home=%~dp0
echo "Current Directory" %Module_Home%
set REGASM=C:\Windows\Microsoft.NET\Framework64\v4.0.30319
echo "RegAsm Directory" %REGASM%

IF exist "%REGASM%\RegAsm.exe" ( 
call %REGASM%\RegAsm.exe /u %Module_Home%\LTCmeAddIn.dll
) else (
echo "RegAsm Directory Is Missing in Client Machine"
)
endlocal
pause