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

1. Copy `Sync-Vba-Template.ps1` into your project folder.
2. Open `Sync-Vba-Template.ps1` and set `$presentationPath` to the full path of your PowerPoint presentation.
3. Keep the `.bas`, `.cls`, and `.frm` files you want in the project in the `vba-src` folder next to the script. The folder is created automatically if it does not exist.

## Usage

Open PowerShell in the folder containing `Sync-Vba-Template.ps1`:

```powershell
# Export modules from the open presentation into vba-src
.\Sync-Vba-Template.ps1 -Export

# Synchronize the open presentation from vba-src
.\Sync-Vba-Template.ps1
```

You can also press **Ctrl+Shift+B** in Visual Studio Code to run synchronization. This shortcut does not export modules; use `-Export` in PowerShell for that.

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

1. Copia `Sync-Vba-Template.ps1` en la carpeta de tu proyecto.
2. Abre `Sync-Vba-Template.ps1` y asigna a `$presentationPath` la ruta completa de la presentación de PowerPoint.
3. Guarda en la carpeta `vba-src`, junto al script, los archivos `.bas`, `.cls` y `.frm` que quieras mantener en el proyecto. La carpeta se crea automáticamente si no existe.

## Uso

Abre PowerShell en la carpeta que contiene `Sync-Vba-Template.ps1`:

```powershell
# Exportar módulos de la presentación abierta a vba-src
.\Sync-Vba-Template.ps1 -Export

# Sincronizar la presentación abierta desde vba-src
.\Sync-Vba-Template.ps1
```

También puedes pulsar **Ctrl+Shift+B** en Visual Studio Code para sincronizar. Este atajo no exporta módulos; para exportar, usa `-Export` en PowerShell.

Al exportar, el script no sobrescribe archivos fuente existentes en `vba-src`.

## Créditos y uso

Creado por EamForexon tm con la asistencia de Copilot y GPT Luna 6. El uso, la modificación y la distribución de este código son libres y de código abierto.
