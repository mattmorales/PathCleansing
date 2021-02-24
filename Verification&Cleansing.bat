@ECHO OFF
SET  C:\Users\mmorales\PathCleansing=%~dp0
SET PowerShellScriptPath=% C:\Users\mmorales\PathCleansing%PathConfirmation.ps1
PowerShell  -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%C:\Users\mmorales\PathCleansing\PathConfirmation.ps1%""' -Verb RunAs}";