@echo off
color 0a
cls
@echo.
@echo.
@echo                 在右键菜单上添加插图功能键
@echo                 for EXCEL 03, 07 and 2010
@echo.
@echo.
pause

@copy .\personal.xlsb "%userprofile%\Application Data\Microsoft\Excel\XLSTART\" >nul 2>nul
if %errorlevel% == 0  (echo 07 addin Install finished.) else (echo 07 addin install failed.)

attrib "%userprofile%\Application Data\Microsoft\Excel\XLSTART\personal.xls" -H >nul 2>nul
@copy .\personal.xls "%userprofile%\Application Data\Microsoft\Excel\XLSTART\" >nul 2>nul
if %errorlevel% == 0  (
echo 03 addin Install finished.
attrib "%userprofile%\Application Data\Microsoft\Excel\XLSTART\personal.xls" +H
attrib "%userprofile%\Application Data\Microsoft\Excel\XLSTART\personal.xlsb" +H
) else (echo 03 addin install failed.)
pause