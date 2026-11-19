@echo off
:: Desactivar el eco de los comandos en la consola

:: Habilitar el entorno local para las variables
setlocal

:: Cambiar el directorio de trabajo a la carpeta donde se encuentra este archivo .bat
cd /d "%~dp0"

:: Comprobar si la sesion actual cuenta con permisos de administrador
net session >nul 2>&1

:: Evaluar si no se cuenta con permisos de administrador para solicitar elevacion por UAC
if %errorLevel% neq 0 (
    :: Notificar en pantalla la solicitud de permisos administrativos
    echo Solicitando permisos de administrador...
    :: Reejecutar el script elevando privilegios mediante PowerShell con el verbo RunAs
    powershell -NoProfile -Command "Start-Process cmd.exe -ArgumentList '/c \"\"%~f0\" %*\"' -Verb RunAs"
    :: Finalizar el proceso actual sin privilegios
    exit /b
)

:: Definir la ruta del script de PowerShell principal
set "PS_SCRIPT=%~dp0bloquearWPS.ps1"

:: Verificar si el archivo no existe con 'bloquearWPS.ps1' para verificar 'bloquerWPS.ps1'
if not exist "%PS_SCRIPT%" (
    :: Asignar la ruta alternativa con el nombre bloquerWPS.ps1 si existe
    if exist "%~dp0bloquerWPS.ps1" (
        set "PS_SCRIPT=%~dp0bloquerWPS.ps1"
    )
)

:: Ejecutar el script PowerShell con permisos de administrador y bypass temporal de ExecutionPolicy
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*

:: Pausar la ventana para que el usuario pueda revisar los mensajes antes de que se cierre
pause

