# BloquearWPS

Script de PowerShell para bloquear la instalación de WPS Office mediante directivas de AppLocker en Windows.

## Ejecución rápida (Recomendado)

Para ejecutar el script sin necesidad de abrir una consola previa ni preocuparse por la directiva de ejecución (`ExecutionPolicy`) o permisos de administrador:

1. Haz doble clic en el archivo [ejecutar.bat](file:///home/secinfra/Proyectos/BloquearWPS-1/ejecutar.bat).
2. Acepta la ventana emergente de permisos de administrador (**UAC**).
3. Selecciona la plantilla de reglas deseada en el menú interactivo.

### Ejecución con parámetros

También puedes pasar el nombre del conjunto de reglas directamente desde la línea de comandos:

```cmd
ejecutar.bat bloquearWPS
```
o
```cmd
ejecutar.bat sinReglas
```

