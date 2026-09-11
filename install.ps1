<#
.SYNOPSIS
    Installs the Zotero Citations add-in for PowerPoint. No Node.js, no npm, no Python.

.DESCRIPTION
    1. compiles the helper (zotero-addon/helper/ZoteroHelper.cs) with the C# compiler that
       ships with Windows and copies it, the add-in web files and the manifest into
       %LOCALAPPDATA%\ZoteroCitations
    2. makes sure a trusted localhost certificate exists (for https://localhost:23000)
    3. registers the add-in with PowerPoint (developer sideload)
    4. starts the helper and, by default, keeps it running at every logon

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

# $PSScriptRoot is empty when -File is used with a UNC path.
if (-not $Source) { $Source = $PSScriptRoot }
if (-not $Source) { $Source = Split-Path -Parent $PSCommandPath }
if (-not $Source) { throw "Pass -Source <path to the repository>." }

$AddInId = "3bad9358-5068-4f44-b97d-c6dc509f510b"
$DeveloperKey = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
$ShortcutPath = Join-Path ([Environment]::GetFolderPath("Startup")) "Zotero Citations Server.lnk"

function Find-CSharpCompiler {
    foreach ($candidate in @(
            "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
            "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe")) {
        if (Test-Path $candidate) { return $candidate }
    }
    throw "The C# compiler that ships with Windows (csc.exe) was not found."
}

function Test-PemCertificateUsable {
    param([string]$Directory)

    $crt = Join-Path $Directory "localhost.crt"
    if (-not (Test-Path $crt)) { return $false }

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

function Test-PfxUsable {
    param([string]$Directory)

    $pfx = Join-Path $Directory "localhost.pfx"
    if (-not (Test-Path $pfx)) { return $false }
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfx, "")
    } catch {
        return $false
    }
    return ($cert.HasPrivateKey -and $cert.NotAfter -gt (Get-Date))
}

function Export-DevCertificatePfx {
    param([string]$Directory)

    $crt = Join-Path $Directory "localhost.crt"
    $pfx = Join-Path $Directory "localhost.pfx"
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($crt)
    if (-not $cert.HasPrivateKey) {
        # A PEM certificate carries no key, so take the matching one from the user store.
        $stored = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
        if (-not $stored) { return $false }
        $cert = $stored
    }
    $empty = New-Object System.Security.SecureString
    Export-PfxCertificate -Cert $cert -FilePath $pfx -Password $empty | Out-Null
    return (Test-Path $pfx)
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

    # The helper reads an unencrypted PKCS#12 file; a PEM private key is not used any more.
    Remove-Item (Join-Path $Directory "localhost.key") -Force -ErrorAction SilentlyContinue
    if (-not (Export-DevCertificatePfx -Directory $Directory)) {
        throw "Could not export the certificate for the helper."
    }

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

$helperSource = Join-Path $Source "zotero-addon\helper\ZoteroHelper.cs"
$wwwSource = Join-Path $Source "zotero-addon\www"
$manifestSource = Join-Path $Source "zotero-addon\manifest.xml"
foreach ($path in @($helperSource, $wwwSource, $manifestSource)) {
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

$compiler = Find-CSharpCompiler

# Stop a helper that is still running the files we are about to replace (Windows keeps the exe locked).
$stale = Get-CimInstance Win32_Process -Filter "Name = 'python.exe' OR Name = 'pythonw.exe' OR Name = 'ZoteroHelper.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like "*$InstallDir*" }
foreach ($process in $stale) {
    Write-Host "Stopping the previous helper server (PID $($process.ProcessId))"
    Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
}
if ($stale) { Start-Sleep -Milliseconds 700 }

# --- 2. compile the helper and copy the add-in ------------------------------

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $InstallDir "www") | Out-Null
Copy-Item -Path (Join-Path $wwwSource "*") -Destination (Join-Path $InstallDir "www") -Recurse -Force
Copy-Item -Path $manifestSource -Destination (Join-Path $InstallDir "manifest.xml") -Force
Remove-Item (Join-Path $InstallDir "server.py") -Force -ErrorAction SilentlyContinue  # from the Python version

$helperExe = Join-Path $InstallDir "ZoteroHelper.exe"
Write-Host "Compiling the helper with $compiler"
& $compiler /nologo /optimize+ /r:System.Web.Extensions.dll /out:"$helperExe" "$helperSource"
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $helperExe)) {
    throw "Compiling the helper failed (see the compiler output above)."
}
Write-Host "Installed files in $InstallDir"

# --- 3. localhost certificate ----------------------------------------------

# A trusted PEM certificate from an earlier version only needs its PKCS#12 twin.
if ((Test-PemCertificateUsable -Directory $CertDir) -and -not (Test-PfxUsable -Directory $CertDir)) {
    Export-DevCertificatePfx -Directory $CertDir | Out-Null
    Remove-Item (Join-Path $CertDir "localhost.key") -Force -ErrorAction SilentlyContinue
}

if ((Test-PemCertificateUsable -Directory $CertDir) -and (Test-PfxUsable -Directory $CertDir)) {
    Write-Host "Reusing the trusted localhost certificate in $CertDir"
} else {
    Write-Host "Creating a trusted localhost certificate in $CertDir"
    New-DevCertificates -Directory $CertDir
    if (-not ((Test-PemCertificateUsable -Directory $CertDir) -and (Test-PfxUsable -Directory $CertDir))) {
        throw "The generated certificate is not usable. Check $CertDir and the CurrentUser certificate stores."
    }
}

# --- 4. register the add-in with PowerPoint --------------------------------

if (-not (Test-Path $DeveloperKey)) { New-Item -Path $DeveloperKey -Force | Out-Null }
New-ItemProperty -Path $DeveloperKey -Name $AddInId `
    -Value (Join-Path $InstallDir "manifest.xml") -PropertyType String -Force | Out-Null
Write-Host "Registered the add-in for PowerPoint"

# --- 5. launcher and autostart ---------------------------------------------

$helperArguments = "--cert-dir `"$CertDir`" --static-port $staticPort"

$launcherPath = Join-Path $InstallDir "run-server.cmd"
$launcher = @(
    "@echo off",
    "title Zotero Citations helper server",
    "cd /d `"%~dp0`"",
    "`"%~dp0ZoteroHelper.exe`" $helperArguments",
    "echo.",
    "echo The helper stopped. Press any key to close this window.",
    "pause >nul"
) -join "`r`n"
[System.IO.File]::WriteAllText($launcherPath, $launcher + "`r`n")

# The logon shortcut runs this VBScript so the console application stays invisible.
$hiddenLauncherPath = Join-Path $InstallDir "run-hidden.vbs"
# A literal here-string keeps the VBScript quoting readable: VBScript escapes a quote by doubling it.
$hiddenLauncher = @'
' Starts the helper without a window; used by the logon shortcut.
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
exe = here & "\ZoteroHelper.exe"
args = "--cert-dir ""__CERTDIR__"" --static-port __PORT__ --hidden --log """ & here & "\server.log"""
shell.Run """" & exe & """" & " " & args, 0, False
'@
$hiddenLauncher = $hiddenLauncher.Replace("__CERTDIR__", $CertDir).Replace("__PORT__", "$staticPort")
[System.IO.File]::WriteAllText($hiddenLauncherPath, $hiddenLauncher + "`r`n")

if ($NoAutostart) {
    Remove-Item -Path $ShortcutPath -Force -ErrorAction SilentlyContinue
    Write-Host "Autostart disabled (run $launcherPath when you need the add-in)"
} else {
    New-Shortcut -Path $ShortcutPath -Target (Join-Path $env:SystemRoot "System32\wscript.exe") `
        -Arguments "`"$hiddenLauncherPath`"" -WorkingDirectory $InstallDir
    Write-Host "The helper starts at every logon"
}

# --- 6. (re)start the helper and verify -------------------------------------

# The previous helper was stopped before compiling, so this starts the new build.

if (Test-PortListening -Port 8000) {
    Write-Host "Port 8000 is already in use; not starting a second server"
} else {
    Start-Process -FilePath $helperExe -WorkingDirectory $InstallDir -WindowStyle Hidden `
        -ArgumentList "--cert-dir", "`"$CertDir`"", "--static-port", "$staticPort", "--log", "`"$InstallDir\server.log`""
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
Write-Host "Done. Restart PowerPoint, then start the add-in from Home > Add-ins > Zotero Citations."
Write-Host "Office keeps sideloaded add-ins in the Add-ins tab; the Zotero Tools group appears on the Home tab while it is running."
Write-Host "Zotero must be running with the Better BibTeX plugin for citing and bibliographies."
Write-Host "Uninstall with: powershell -ExecutionPolicy Bypass -File `"$Source\uninstall.ps1`""
