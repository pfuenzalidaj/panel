@echo off
setlocal

set "BASE=%~dp0"
set "PY=C:\Users\pablo.vasquez\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "BUILDER=%BASE%aplicacion\generador_dashboard.py"
set "XLSX=%BASE%fuentes\Tickets_Incidentes.xlsx"
set "PROBLEMS=%BASE%fuentes\Informe_Problemas.xlsx"
set "HTML=%BASE%dashboard\Dashboard_Ejecutivo_Gestion_TI.html"

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

if not exist "%PROBLEMS%" (
  echo No se encontro: %PROBLEMS%
  goto :error
)

if not exist "%HTML%" (
  echo No se encontro: %HTML%
  goto :error
)

echo Actualizando Dashboard Ejecutivo de Gestion TI...
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
