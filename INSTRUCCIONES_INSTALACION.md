# Instrucciones de instalación — Generador de Testimonios

Este documento explica cómo instalar y poner en marcha el programa **Generador de Testimonios** en una computadora con Windows, a partir de los archivos publicados en **GitHub Releases**.

---

## 1. Requisitos previos

Antes de instalar, verifique que el equipo cumpla con lo siguiente:

- **Windows 10 u 11 (64 bits)**.
- **Microsoft Word instalado** (el programa utiliza Word para generar los documentos).
- **Conexión a internet** (necesaria al menos la primera vez, para descargar la aplicación).
- Permisos para ejecutar programas descargados de internet.

---

## 2. Descargar el programa

1. Abra su navegador y vaya a la página de releases del proyecto:

   **link de repositorio con releases**

2. En el release más reciente (el primero de la lista, marcado como *Latest*), busque la sección **Assets**.

3. Descargue el archivo:

   **`GeneradorTestimoniosLauncher-win64.zip`**

   > No es necesario descargar los demás archivos (`GeneradorTestimonios-win64.zip` y `checksums.txt`): el Launcher los descarga y verifica automáticamente.

---

## 3. Descomprimir el archivo ZIP

1. Localice el archivo descargado (normalmente en la carpeta **Descargas**).
2. Haga **clic derecho** sobre `GeneradorTestimoniosLauncher-win64.zip` y seleccione **Extraer todo...**
3. Elija una ubicación permanente y fácil de recordar, por ejemplo:
   - `C:\GeneradorTestimonios` o
   - `Documentos\GeneradorTestimonios`

   > **Importante:** no ejecute el programa directamente desde dentro del ZIP ni desde la carpeta Descargas. Extráigalo primero en una carpeta fija, porque esa será la ubicación desde la que se abrirá el programa todos los días.

4. Al terminar tendrá una carpeta llamada **`GeneradorTestimoniosLauncher`** que contiene, entre otros archivos, el programa **`GeneradorTestimoniosLauncher.exe`**.

---

## 4. Ejecutar el programa por primera vez

1. Abra la carpeta donde extrajo los archivos.
2. Haga **doble clic** en **`GeneradorTestimoniosLauncher.exe`**.
3. Si Windows muestra el aviso azul **"Windows protegió su PC"** (SmartScreen):
   - Haga clic en **Más información**.
   - Luego en **Ejecutar de todas formas**.

   > Este aviso es normal en programas que no tienen firma digital empresarial; no indica un problema.
4. El Launcher se conectará a GitHub Releases y **descargará automáticamente la versión más reciente** del Generador de Testimonios. Este proceso puede tardar unos minutos la primera vez, según la velocidad de internet.
5. Al finalizar, el programa principal se abrirá de forma automática.

La aplicación se instala en:

```
%LOCALAPPDATA%\GeneradorTestimonios
```

No necesita permisos de administrador.

---

## 5. Crear un acceso directo (recomendado)

Para abrir el programa fácilmente:

1. Haga **clic derecho** sobre `GeneradorTestimoniosLauncher.exe`.
2. Seleccione **Mostrar más opciones → Enviar a → Escritorio (crear acceso directo)**.
3. Use ese acceso directo para abrir el programa a diario.

> Abra **siempre** el programa mediante `GeneradorTestimoniosLauncher.exe` (o su acceso directo). El Launcher es el encargado de verificar e instalar las actualizaciones antes de abrir la aplicación.

---

## 6. Actualizaciones

- Las actualizaciones son **automáticas**: cada vez que abre el Launcher, este verifica si hay una versión nueva publicada en GitHub Releases.
- Si existe una versión más reciente, la descarga, verifica su integridad (SHA-256) y la instala sin que usted tenga que hacer nada.
- Si no hay internet en ese momento, el programa se abre con la última versión que ya esté instalada en el equipo.

---

## 7. Solución de problemas

| Problema | Solución |
|---|---|
| **"Windows protegió su PC"** | Clic en *Más información → Ejecutar de todas formas*. |
| **El antivirus bloquea la descarga o el ejecutable** | Agregue una exclusión para la carpeta donde extrajo el Launcher y para `%LOCALAPPDATA%\GeneradorTestimonios`. |
| **"No hay versión local y no se pudo consultar GitHub Releases"** | Verifique la conexión a internet y vuelva a intentar. La primera ejecución requiere internet sí o sí. |
| **"No se pudo actualizar porque la aplicación está abierta o bloqueada"** | Cierre completamente el Generador de Testimonios y vuelva a abrir el Launcher. |
| **"La verificación del paquete falló (SHA256 inválido)"** | La descarga se dañó. Revise su conexión y vuelva a intentar; el Launcher repetirá la descarga. |
| **El programa genera errores al crear documentos** | Confirme que **Microsoft Word** esté instalado y que pueda abrirse normalmente. |
| **El programa no abre y no muestra mensaje** | Ejecute el Launcher de nuevo; si persiste, borre la carpeta `%LOCALAPPDATA%\GeneradorTestimonios` y vuelva a intentar (se descargará de nuevo). |

---

## 8. Desinstalación

El programa no deja registros del sistema. Para desinstalarlo por completo:

1. Elimine la carpeta donde extrajo el **Launcher**.
2. Elimine la carpeta **`%LOCALAPPDATA%\GeneradorTestimonios`** (pegue esa ruta en la barra de direcciones del Explorador de archivos y presione Enter para llegar a ella).
3. Elimine el acceso directo del Escritorio, si lo creó.

---

## Resumen rápido

1. Descargar `GeneradorTestimoniosLauncher-win64.zip` desde **GitHub Releases**.
2. Extraer el ZIP en una carpeta fija.
3. Ejecutar `GeneradorTestimoniosLauncher.exe` (autorizar en SmartScreen si aparece el aviso).
4. Esperar la descarga automática de la aplicación. ¡Listo!
