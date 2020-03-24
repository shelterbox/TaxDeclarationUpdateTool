@echo off
set "script=TaxDeclarationUpdateTool.ps1"
echo %script%
PowerShell -NoProfile -ExecutionPolicy RemoteSigned -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy RemoteSigned -File ""%~dp0%script%""' -Verb RunAs}";
