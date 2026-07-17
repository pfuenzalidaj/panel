@echo off
setlocal

set "BASE=%~dp0"
set "PY=C:\Users\pablo.vasquez\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "APP=%BASE%aplicacion\servidor_actualizador.py"

if not exist "%PY%" (
  where python >nul 2>&1
  if errorlevel 1 (
    echo No se encontro una instalacion de Python.
    goto :error
  )
  set "PY=python"
)

if not exist "%APP%" (
  echo No se encontro: %APP%
  goto :error
)

"%PY%" "%APP%" %*
if errorlevel 1 goto :error
exit /b 0

:error
echo.
echo No fue posible iniciar el actualizador.
pause
exit /b 1
