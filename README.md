# Sync-Vba

PowerShell scripts for exporting and synchronizing VBA modules in an open PowerPoint presentation.

## What it does

- **Export** (`-Export`): exports standard modules (`.bas`), class modules (`.cls`), and forms (`.frm`) from the presentation's VBA project into a `vba-src` folder next to the script.
- **Synchronize** (no parameters): imports and updates those components from `vba-src`. Components of these types that do not have a matching source file in that folder are removed from the VBA project.

> **Warning:** Synchronization can remove modules from your presentation. Back up the presentation before running it.

## Requirements

- Windows with Microsoft PowerPoint and Windows PowerShell.
- The target presentation must already be open in PowerPoint.
- If the VBA project is password-protected, enter its password in the VBA editor to unlock the project before running the script. **Do not remove or disable the password.** The project only needs to be unlocked while the script runs.
- In PowerPoint's Trust Center, enable **Trust access to the VBA project object model**.

## Setup

1. Clone or download this repository into your project folder. This repository includes `.vscode/tasks.json` to configure the **Ctrl+Shift+B** shortcut. If you only copy the script, also copy the `.vscode` folder.
2. Open `Sync-Vba-Template.ps1` and set `$presentationPath` to the full path of your PowerPoint presentation.
3. Keep the `.bas`, `.cls`, and `.frm` files you want in the project in the `vba-src` folder next to the script. The folder is created automatically if it does not exist.

## Usage

To export the VBA components for the first time, open PowerShell in the folder containing `Sync-Vba-Template.ps1` and run:

```powershell
# Export modules from the open presentation into vba-src
.\Sync-Vba-Template.ps1 -Export

# Synchronize the open presentation from vba-src
.\Sync-Vba-Template.ps1
```

After opening the project folder in Visual Studio Code, press **Ctrl+Shift+B** to synchronize. The included task runs synchronization only; use `-Export` in PowerShell to export modules.

When exporting, the script will not overwrite existing source files in `vba-src`.

## Credits and use

Created by EamForexon tm with the assistance of Copilot and GPT Luna 6. Use, modification, and distribution of this code are free and open source.

---

# Sync-Vba (español)

Scripts de PowerShell para exportar y sincronizar módulos VBA en una presentación de PowerPoint abierta.

## Qué hace

- **Exportar** (`-Export`): exporta módulos estándar (`.bas`), módulos de clase (`.cls`) y formularios (`.frm`) del proyecto VBA de la presentación a una carpeta `vba-src` junto al script.
- **Sincronizar** (sin parámetros): importa y actualiza esos componentes desde `vba-src`. Los componentes de esos tipos que no tengan un archivo fuente correspondiente en esa carpeta se quitan del proyecto VBA.

> **Aviso:** La sincronización puede quitar módulos de la presentación. Haz una copia de seguridad antes de ejecutarla.

## Requisitos

- Windows con Microsoft PowerPoint y Windows PowerShell.
- La presentación de destino debe estar abierta en PowerPoint.
- Si el proyecto VBA está protegido con contraseña, introduce la contraseña en el editor de VBA para desbloquearlo antes de ejecutar el script. **No hace falta quitar ni desactivar la contraseña.** El proyecto solo debe estar desbloqueado mientras se ejecuta el script.
- En el Centro de confianza de PowerPoint, habilita **Confiar en el acceso al modelo de objetos de proyectos de VBA**.

## Configuración

1. Clona o descarga este repositorio en la carpeta de tu proyecto. El repositorio incluye `.vscode/tasks.json` para configurar el atajo **Ctrl+Shift+B**. Si solo copias el script, copia también la carpeta `.vscode`.
2. Abre `Sync-Vba-Template.ps1` y asigna a `$presentationPath` la ruta completa de la presentación de PowerPoint.
3. Guarda en la carpeta `vba-src`, junto al script, los archivos `.bas`, `.cls` y `.frm` que quieras mantener en el proyecto. La carpeta se crea automáticamente si no existe.

## Uso

Para exportar los componentes VBA por primera vez, abre PowerShell en la carpeta que contiene `Sync-Vba-Template.ps1` y ejecuta:

```powershell
# Exportar módulos de la presentación abierta a vba-src
.\Sync-Vba-Template.ps1 -Export

# Sincronizar la presentación abierta desde vba-src
.\Sync-Vba-Template.ps1
```

Después de abrir la carpeta del proyecto en Visual Studio Code, pulsa **Ctrl+Shift+B** para sincronizar. La tarea incluida solo sincroniza; para exportar módulos, usa `-Export` en PowerShell.

Al exportar, el script no sobrescribe archivos fuente existentes en `vba-src`.

## Créditos y uso

Creado por EamForexon tm con la asistencia de Copilot y GPT Luna 6. El uso, la modificación y la distribución de este código son libres y de código abierto.
