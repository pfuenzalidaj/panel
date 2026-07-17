@echo off
setlocal

set "BASE=%~dp0"
set "PY=C:\Users\pablo.vasquez\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "BUILDER=%BASE%build_excel_pilot.py"
set "XLSX=%BASE%Tickets_full__al 30_06_2026.xlsx"
set "HTML=%BASE%CEN_Dashboard_Ejecutivo_Jul2026_Piloto_Excel.html"

if not exist "%PY%" (
  where python >nul 2>&1
  if errorlevel 1 (
    echo No se encontro una instalacion de Python.
    goto :error
  )
  set "PY=python"
)

if not exist "%BUILDER%" (
  echo No se encontro: %BUILDER%
  goto :error
)

if not exist "%XLSX%" (
  echo No se encontro: %XLSX%
  goto :error
)

if not exist "%HTML%" (
  echo No se encontro: %HTML%
  goto :error
)

echo Refrescando dashboard piloto desde Tickets_full__al 30_06_2026.xlsx...
"%PY%" "%BUILDER%"
if errorlevel 1 (
  goto :error
)

echo.
echo Dashboard actualizado:
echo %HTML%
echo.
if /I not "%~1"=="--no-pause" pause
exit /b 0

:error
echo.
echo Error al regenerar el dashboard.
if /I not "%~1"=="--no-pause" pause
exit /b 1
pause
