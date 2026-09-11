<#
.SYNOPSIS
    Removes the Zotero Citations add-in from PowerPoint and stops the helper server.

.DESCRIPTION
    Undoes install.ps1: stops the helper server, removes the autostart shortcut,
    unregisters the add-in, and deletes the installed files.
    The localhost certificate is kept unless -RemoveCertificate is given.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
#>
[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA "ZoteroCitations"),
    [switch]$RemoveCertificate
)

$ErrorActionPreference = "Stop"

$AddInId = "3bad9358-5068-4f44-b97d-c6dc509f510b"
$DeveloperKey = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
$ShortcutPath = Join-Path ([Environment]::GetFolderPath("Startup")) "Zotero Citations Server.lnk"

$servers = Get-CimInstance Win32_Process -Filter "Name = 'python.exe' OR Name = 'pythonw.exe' OR Name = 'ZoteroHelper.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like "*$InstallDir*" }
foreach ($server in $servers) {
    Write-Host "Stopping helper server (PID $($server.ProcessId))"
    Stop-Process -Id $server.ProcessId -Force -ErrorAction SilentlyContinue
}

Remove-Item -Path $ShortcutPath -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $DeveloperKey -Name $AddInId -Force -ErrorAction SilentlyContinue
Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Unregistered the add-in and removed $InstallDir"

if ($RemoveCertificate) {
    Get-ChildItem Cert:\CurrentUser\Root |
        Where-Object { $_.Subject -like "*Developer CA for Zotero Citations*" } |
        Remove-Item -Force
    Write-Host "Removed the trusted localhost certificate"
} else {
    Write-Host "Kept the localhost certificate (use -RemoveCertificate to delete it)"
}

Write-Host "Restart PowerPoint to unload the pane."
