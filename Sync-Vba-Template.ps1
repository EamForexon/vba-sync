<#
Sync-Vba - reusable template for PowerPoint projects

What it does:
  - With -Export, exports standard modules (.bas), class modules (.cls), and
    forms (.frm) from the VBA project to the "vba-src" folder next to this
    script.
  - With no parameters, imports and updates those components from "vba-src".
    It also removes components of those types that do not have a matching
    file in that folder. Back up your presentation before synchronizing.

How to use it:
  1. Copy this script into your project folder.
  2. Edit $presentationPath below and set it to the full path of your
     PowerPoint presentation.
  3. Open that presentation in PowerPoint. If the VBA project is protected,
     enter its password in the VBA editor to unlock it. You do not need to
     remove or disable the password; the project just needs to be unlocked
     while the script runs.
  4. In PowerPoint, enable "Trust access to the VBA project object model"
     in the Trust Center.
  5. Open PowerShell in the script's folder and run:
       .\Sync-Vba-Template.ps1 -Export   # export the components
       .\Sync-Vba-Template.ps1           # synchronize from vba-src
     You can also synchronize from Visual Studio Code with Ctrl+Shift+B.
     That shortcut runs synchronization only; use -Export to export.

The vba-src folder is created automatically next to the script. To synchronize,
put the .bas, .cls, and .frm files you want to keep in the project there.

Created by EamForexon tm with the assistance of Copilot and GPT Luna 6.
The use, modification, and distribution of this code are free and open source.
#>

[CmdletBinding()]
param(
    [switch]$Export
)

$ErrorActionPreference = 'Stop'
$presentationPath = 'C:\Path\To\YourPresentation.pptm' # EDIT: full path to your presentation
$sourceDirectory = Join-Path $PSScriptRoot 'vba-src'

if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $sourceDirectory | Out-Null
}

try {
    $powerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
} catch {
    throw 'Open the presentation in PowerPoint first and leave the application running.'
}

$presentation = $null
$expectedPath = [IO.Path]::GetFullPath($presentationPath)
foreach ($openPresentation in $powerPoint.Presentations) {
    if ([string]::Equals([IO.Path]::GetFullPath($openPresentation.FullName), $expectedPath, [StringComparison]::OrdinalIgnoreCase)) {
        $presentation = $openPresentation
        break
    }
}

if (-not $presentation) {
    throw "The expected presentation is not open: $presentationPath"
}

try {
    $project = $presentation.VBProject
    if ($project.Protection -ne 0) {
        throw 'The VBA project is still locked. Unlock it manually in the VBA editor and run the command again.'
    }
} catch {
    throw "Cannot access the VBA project. Unlock it manually in PowerPoint and verify that 'Trust access to the VBA project object model' is enabled. Details: $($_.Exception.Message)"
}

function Get-ComponentExtension([int]$type) {
    switch ($type) {
        1 { return '.bas' }
        2 { return '.cls' }
        3 { return '.frm' }
        default { return $null }
    }
}

if ($Export) {
    $components = @()
    for ($index = 1; $index -le $project.VBComponents.Count; $index++) {
        $component = $project.VBComponents.Item($index)
        $extension = Get-ComponentExtension $component.Type
        if ($extension) {
            $components += [pscustomobject]@{ Component = $component; Extension = $extension }
        }
    }

    $destinations = foreach ($entry in $components) {
        $path = Join-Path $sourceDirectory ($entry.Component.Name + $entry.Extension)
        if (Test-Path -LiteralPath $path) {
            throw "Export cancelled to avoid overwriting an existing source file: $path"
        }
        if ($entry.Extension -eq '.frm' -and (Test-Path -LiteralPath ([IO.Path]::ChangeExtension($path, '.frx')))) {
            throw "Export cancelled to avoid overwriting an existing form file: $path"
        }
        [pscustomobject]@{ Component = $entry.Component; Path = $path }
    }

    foreach ($destination in $destinations) {
        $destination.Component.Export($destination.Path)
        Write-Host "Exported: $([IO.Path]::GetFileName($destination.Path))"
    }

    Write-Host "Export complete. Sources: $sourceDirectory"
    return
}

$sourceFiles = @(Get-ChildItem -LiteralPath $sourceDirectory -File | Where-Object { $_.Extension -in @('.bas', '.cls', '.frm') })
if ($sourceFiles.Count -eq 0) {
    throw "No .bas, .cls, or .frm files found in $sourceDirectory. Run this script with -Export first."
}

$desiredModules = @{}
$items = @()
foreach ($file in $sourceFiles) {
    $text = [IO.File]::ReadAllText($file.FullName)
    $match = [regex]::Match($text, '(?m)^Attribute VB_Name\s*=\s*"([^"]+)"')
    if (-not $match.Success) {
        throw "Attribute VB_Name not found in $($file.Name). The presentation was not modified."
    }

    $moduleName = $match.Groups[1].Value
    if ($desiredModules.ContainsKey($moduleName)) {
        throw "More than one source file found for module '$moduleName'. The presentation was not modified."
    }

    $expectedType = switch ($file.Extension.ToLowerInvariant()) {
        '.bas' { 1 }
        '.cls' { 2 }
        '.frm' { 3 }
    }
    $desiredModules[$moduleName] = $true
    $items += [pscustomobject]@{ File = $file; Name = $moduleName; Type = $expectedType }
}

foreach ($item in $items) {
    for ($index = 1; $index -le $project.VBComponents.Count; $index++) {
        $component = $project.VBComponents.Item($index)
        if ($component.Name -eq $item.Name -and $component.Type -ne $item.Type) {
            throw "The type of '$($item.Name)' does not match $($item.File.Extension). The presentation was not modified."
        }
    }
}

$presentation.Save()
foreach ($item in $items) {
    $existing = $null
    for ($index = 1; $index -le $project.VBComponents.Count; $index++) {
        $component = $project.VBComponents.Item($index)
        if ($component.Name -eq $item.Name -and (Get-ComponentExtension $component.Type)) {
            $existing = $component
            break
        }
    }

    $backupName = $null
    if ($existing) {
        $backupName = $existing.Name
        $existing.Name = 'VSCBackup_' + [guid]::NewGuid().ToString('N').Substring(0, 12)
    }

    $imported = $null
    try {
        $imported = $project.VBComponents.Import($item.File.FullName)
        if ($imported.Name -ne $item.Name) {
            throw "PowerPoint imported '$($item.File.Name)' as '$($imported.Name)' instead of '$($item.Name)'."
        }
        if ($existing) {
            $project.VBComponents.Remove($existing)
        }
    } catch {
        if ($imported) {
            try { $project.VBComponents.Remove($imported) } catch { }
        }
        if ($existing) {
            try { $existing.Name = $backupName } catch { }
        }
        throw "Failed to import $($item.File.Name). VBA changes will not be saved; close the presentation without saving if any partial changes remain. Details: $($_.Exception.Message)"
    }

    Write-Host "Synchronized: $($item.File.Name)"
}

for ($index = $project.VBComponents.Count; $index -ge 1; $index--) {
    $component = $project.VBComponents.Item($index)
    if ((Get-ComponentExtension $component.Type) -and -not $desiredModules.ContainsKey($component.Name)) {
        $removedName = $component.Name
        $project.VBComponents.Remove($component)
        Write-Host "Removed from the project because it no longer exists in vba-src: $removedName"
    }
}

$presentation.Save()
Write-Host "Synchronization complete and saved to: $presentationPath"
