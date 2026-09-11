<#
.SYNOPSIS
    Installs the Zotero Citations add-in for PowerPoint. No Node.js, no npm, no build step.

.DESCRIPTION
    1. copies the add-in web files and the helper server to %LOCALAPPDATA%\ZoteroCitations
    2. makes sure a trusted localhost certificate exists (for https://localhost:23000)
    3. registers the add-in with PowerPoint (developer sideload)
    4. starts the helper server and, by default, keeps it running at every logon

    Run it again at any time to update the installed copy.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\install.ps1

.EXAMPLE
    # from WSL (Windows PowerShell can run this script over the \\wsl.localhost path)
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File install.ps1
#>
[CmdletBinding()]
param(
    [string]$Source,
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA "ZoteroCitations"),
    [string]$CertDir = (Join-Path $env:USERPROFILE ".office-addin-dev-certs"),
    [switch]$NoAutostart
)
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Default to the folder holding this script ($PSScriptRoot is empty when -File is used with a UNC path).
if (-not $Source) { $Source = $PSScriptRoot }
if (-not $Source) { $Source = Split-Path -Parent $PSCommandPath }
if (-not $Source) { throw "Pass -Source <path to the repository>." }

$AddInId = "3bad9358-5068-4f44-b97d-c6dc509f510b"
$DeveloperKey = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
$ShortcutPath = Join-Path ([Environment]::GetFolderPath("Startup")) "Zotero Citations Server.lnk"

function Find-Python {
    param([switch]$Windowless)

    $exeName = if ($Windowless) { "pythonw.exe" } else { "python.exe" }
    $candidates = @()

    # uv keeps its managed interpreters in %APPDATA%\uv\python
    $uvRoot = Join-Path $env:APPDATA "uv\python"
    if (Test-Path $uvRoot) {
        foreach ($dir in Get-ChildItem $uvRoot -Directory) {
            $candidates += (Join-Path $dir.FullName $exeName)
        }
    }

    $onPath = Get-Command $exeName -ErrorAction SilentlyContinue
    if ($onPath) { $candidates += $onPath.Source }

    foreach ($candidate in $candidates) {
        if (-not (Test-Path $candidate)) { continue }
        if ($candidate -like "*\WindowsApps\*") { continue }  # Microsoft Store alias, not a real Python
        & $candidate -c "import sys" 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0) { return $candidate }
    }
    return $null
}

function Test-DevCertificateUsable {
    param([string]$Directory)

    $crt = Join-Path $Directory "localhost.crt"
    $key = Join-Path $Directory "localhost.key"
    if (-not (Test-Path $crt) -or -not (Test-Path $key)) { return $false }

    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($crt)
    } catch {
        return $false
    }
    if ($cert.NotAfter -lt (Get-Date).AddDays(7)) { return $false }

    $san = $cert.Extensions | Where-Object { $_.Oid.Value -eq "2.5.29.17" } | Select-Object -First 1
    if (-not $san -or $san.Format($false) -notmatch "localhost") { return $false }

    $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
    $chain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
    return $chain.Build($cert)
}

function Write-Pem {
    param([byte[]]$Der, [string]$Label, [string]$Path)

    $base64 = [Convert]::ToBase64String($Der)
    $builder = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $base64.Length; $i += 64) {
        [void]$builder.AppendLine($base64.Substring($i, [Math]::Min(64, $base64.Length - $i)))
    }
    [System.IO.File]::WriteAllText($Path, "-----BEGIN $Label-----`n" + $builder.ToString() + "-----END $Label-----`n")
}

function New-DevCertificates {
    param([string]$Directory)

    New-Item -ItemType Directory -Force -Path $Directory | Out-Null

    # Drop certificates from an earlier install that have expired.
    Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Subject -like "*Zotero Citations Add-in*" -and $_.NotAfter -lt (Get-Date) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
    Get-ChildItem Cert:\CurrentUser\Root |
        Where-Object { $_.Subject -like "*Zotero Citations Add-in*" -and $_.NotAfter -lt (Get-Date) } |
        Remove-Item -Force -ErrorAction SilentlyContinue

    $ca = New-SelfSignedCertificate -Subject "CN=Developer CA for Zotero Citations Add-in" `
        -CertStoreLocation Cert:\CurrentUser\My -Type Custom -KeyAlgorithm RSA -KeyLength 2048 `
        -KeyExportPolicy Exportable -KeyUsage CertSign, CRLSign, DigitalSignature -KeyUsageProperty Sign `
        -NotAfter (Get-Date).AddYears(5) -TextExtension @("2.5.29.19={text}ca=1&pathlength=0")

    $leaf = New-SelfSignedCertificate -Subject "CN=localhost" -Signer $ca `
        -CertStoreLocation Cert:\CurrentUser\My -Type Custom -KeyAlgorithm RSA -KeyLength 2048 `
        -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(1) `
        -TextExtension @(
            "2.5.29.19={text}ca=0",
            "2.5.29.17={text}DNS=localhost&IPAddress=127.0.0.1",
            "2.5.29.37={text}1.3.6.1.5.5.7.3.1"
        )

    Write-Pem -Der $ca.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert) `
        -Label "CERTIFICATE" -Path (Join-Path $Directory "ca.crt")
    Write-Pem -Der $leaf.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert) `
        -Label "CERTIFICATE" -Path (Join-Path $Directory "localhost.crt")

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($leaf)
    if (-not $rsa) { $rsa = $leaf.PrivateKey }
    if (-not $rsa -or -not $rsa.Key) { throw "Could not export the private key of the generated certificate." }
    Write-Pem -Der $rsa.Key.Export([System.Security.Cryptography.CngKeyBlobFormat]::Pkcs8PrivateBlob) `
        -Label "PRIVATE KEY" -Path (Join-Path $Directory "localhost.key")

    # Trust the CA for the current user (no admin rights needed).
    Import-Certificate -FilePath (Join-Path $Directory "ca.crt") -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
}

function New-Shortcut {
    param([string]$Path, [string]$Target, [string]$Arguments, [string]$WorkingDirectory)

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($Path)
    $shortcut.TargetPath = $Target
    $shortcut.Arguments = $Arguments
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.Description = "Local helper server for the Zotero Citations PowerPoint add-in"
    $shortcut.WindowStyle = 7
    $shortcut.Save()
}

function Test-PortListening {
    param([int]$Port)

    return [bool](Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue)
}

# --- 1. check the source tree ------------------------------------------------

$serverSource = Join-Path $Source "shared\zotero_proxy_server.py"
$wwwSource = Join-Path $Source "zotero-addon\www"
$manifestSource = Join-Path $Source "zotero-addon\manifest.xml"
foreach ($path in @($serverSource, $wwwSource, $manifestSource)) {
    if (-not (Test-Path $path)) {
        throw "Not found: $path`nRun install.ps1 from the repository root, or pass -Source <repository path>."
    }
}

$manifestText = Get-Content $manifestSource -Raw
if ($manifestText -match "https://localhost:(\d+)/taskpane\.html") {
    $staticPort = [int]$Matches[1]
} else {
    throw "manifest.xml has no https://localhost:<port>/taskpane.html SourceLocation."
}

# --- 2. find Python --------------------------------------------------------

$python = Find-Python
$pythonw = Find-Python -Windowless
if (-not $python -and -not $pythonw) {
    throw "No Python found. Install it with 'winget install astral-sh.uv' (or from python.org) and run install.ps1 again."
}
if (-not $pythonw) { $pythonw = $python }
if (-not $python) { $python = $pythonw }

# --- 3. copy the add-in and the server -------------------------------------

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $InstallDir "www") | Out-Null
Copy-Item -Path $serverSource -Destination (Join-Path $InstallDir "server.py") -Force
Copy-Item -Path (Join-Path $wwwSource "*") -Destination (Join-Path $InstallDir "www") -Recurse -Force
Copy-Item -Path $manifestSource -Destination (Join-Path $InstallDir "manifest.xml") -Force
Write-Host "Installed files in $InstallDir"

# --- 4. localhost certificate ----------------------------------------------

if (-not (Test-DevCertificateUsable -Directory $CertDir)) {
    Write-Host "Creating a trusted localhost certificate in $CertDir"
    New-DevCertificates -Directory $CertDir
    if (-not (Test-DevCertificateUsable -Directory $CertDir)) {
        throw "The generated certificate is not usable. Check $CertDir and the CurrentUser certificate stores."
    }
} else {
    Write-Host "Reusing the trusted localhost certificate in $CertDir"
}

# --- 5. register the add-in with PowerPoint --------------------------------

if (-not (Test-Path $DeveloperKey)) { New-Item -Path $DeveloperKey -Force | Out-Null }
New-ItemProperty -Path $DeveloperKey -Name $AddInId `
    -Value (Join-Path $InstallDir "manifest.xml") -PropertyType String -Force | Out-Null
Write-Host "Registered the add-in for PowerPoint"

# --- 6. launcher and autostart ---------------------------------------------

$launcherPath = Join-Path $InstallDir "run-server.cmd"
$launcher = @(
    "@echo off",
    "title Zotero Citations helper server",
    "cd /d `"%~dp0`"",
    "`"$python`" `"%~dp0server.py`" --cert-dir `"$CertDir`" --static-port $staticPort",
    "echo.",
    "echo The server stopped. Press any key to close this window.",
    "pause >nul"
) -join "`r`n"
[System.IO.File]::WriteAllText($launcherPath, $launcher + "`r`n")

if ($NoAutostart) {
    Remove-Item -Path $ShortcutPath -Force -ErrorAction SilentlyContinue
    Write-Host "Autostart disabled (run $launcherPath when you need the add-in)"
} else {
    New-Shortcut -Path $ShortcutPath -Target $pythonw -WorkingDirectory $InstallDir `
        -Arguments "`"$InstallDir\server.py`" --cert-dir `"$CertDir`" --static-port $staticPort --log `"$InstallDir\server.log`""
    Write-Host "The helper server starts at every logon"
}

# --- 7. (re)start the helper server and verify ------------------------------

# Stop a server that is still running the files we just replaced.
$stale = Get-CimInstance Win32_Process -Filter "Name = 'python.exe' OR Name = 'pythonw.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like "*$InstallDir*" }
foreach ($process in $stale) {
    Write-Host "Stopping the previous helper server (PID $($process.ProcessId))"
    Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
}
if ($stale) { Start-Sleep -Milliseconds 500 }

if (Test-PortListening -Port 8000) {
    Write-Host "Port 8000 is already in use; not starting a second server"
} else {
    Start-Process -FilePath $pythonw -WorkingDirectory $InstallDir -WindowStyle Hidden `
        -ArgumentList "`"$InstallDir\server.py`"", "--cert-dir", "`"$CertDir`"", "--static-port", "$staticPort", "--log", "`"$InstallDir\server.log`""
    Start-Sleep -Milliseconds 1500
    if (Test-PortListening -Port 8000) {
        Write-Host "Helper server started"
    } else {
        Write-Warning "The helper server did not start. See $InstallDir\server.log"
    }
}

try {
    $pane = Invoke-WebRequest -Uri "https://localhost:$staticPort/taskpane.html" -UseBasicParsing -TimeoutSec 10
    Write-Host "Task pane reachable over HTTPS (HTTP $($pane.StatusCode))"
} catch {
    Write-Warning "Could not load https://localhost:$staticPort/taskpane.html : $($_.Exception.Message)"
    Write-Warning "Check $InstallDir\server.log and run $launcherPath in a console window."
}

Write-Host ""
Write-Host "Done. Restart PowerPoint, then open Home > Zotero Tools > Open Zotero Pane."
Write-Host "Start it from Home > Add-ins > Zotero Citations; Office shows sideloaded add-ins there instead of pinning them to the Home tab."
Write-Host "Zotero must be running with the Better BibTeX plugin for citing and bibliographies."
Write-Host "Uninstall with: powershell -ExecutionPolicy Bypass -File `"$Source\uninstall.ps1`""
