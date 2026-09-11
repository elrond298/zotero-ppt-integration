<#
.SYNOPSIS
    Smoke test for the C# helper (zotero-addon/helper/ZoteroHelper.cs).

.DESCRIPTION
    Compiles the helper with the C# compiler that ships with Windows, starts it
    against a stub Better BibTeX and checks every endpoint the add-in uses.
    Runs on Windows PowerShell 5.1, needs no network, no Python and no Node.js.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\tools\test-helper.ps1
#>
[CmdletBinding()]
param(
    [string]$Source,
    [int]$BasePort = 23100
)

$ErrorActionPreference = "Stop"

# $PSScriptRoot is empty when this script is started with -File over a UNC path.
if (-not $Source) { $Source = Split-Path -Parent $PSScriptRoot }
if (-not $Source) { $Source = Split-Path -Parent (Split-Path -Parent $PSCommandPath) }
if (-not $Source) { throw "Pass -Source <repository path>." }

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) { $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
if (-not (Test-Path $csc)) { throw "No built-in C# compiler found (csc.exe)." }

$helperSource = Join-Path $Source "zotero-addon\helper\ZoteroHelper.cs"
$wwwSource = Join-Path $Source "zotero-addon\www"
foreach ($path in @($helperSource, $wwwSource)) {
    if (-not (Test-Path $path)) { throw "Not found: $path (pass -Source <repository path>)" }
}

$apiPort = $BasePort
$staticPort = $BasePort + 1
$downPort = $BasePort + 2
$stubPort = $BasePort + 3
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("zotero-helper-test-" + $PID)
$passed = 0
$failures = New-Object System.Collections.Generic.List[string]
$processes = New-Object System.Collections.Generic.List[object]

function Assert-Equal {
    param([string]$Name, $Expected, $Actual)
    if ("$Expected" -eq "$Actual") {
        $script:passed++
        Write-Host ("  ok   " + $Name)
    } else {
        $script:failures.Add("$Name (expected '$Expected', got '$Actual')")
        Write-Host ("  FAIL " + $Name + " (expected '$Expected', got '$Actual')")
    }
}

function Assert-Match {
    param([string]$Name, [string]$Pattern, [string]$Text)
    if ($Text -match $Pattern) {
        $script:passed++
        Write-Host ("  ok   " + $Name)
    } else {
        $script:failures.Add("$Name (no match for '$Pattern')")
        Write-Host ("  FAIL " + $Name + " (no match for '$Pattern')")
    }
}

function Invoke-Endpoint {
    param([string]$Uri, [string]$Method = "GET", [string]$Body)
    $response = $null
    try {
        if ($Body) {
            $response = Invoke-WebRequest -Method $Method -Uri $Uri -Body $Body -ContentType "application/json" -UseBasicParsing -TimeoutSec 20
        } else {
            $response = Invoke-WebRequest -Method $Method -Uri $Uri -UseBasicParsing -TimeoutSec 20
        }
    } catch [System.Net.WebException] {
        $webResponse = $_.Exception.Response
        if (-not $webResponse) { throw }
        # PowerShell 5.1 buffers the error body before it throws, so read it from the error details.
        $content = ""
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $content = $_.ErrorDetails.Message }
        return New-Object PSObject -Property @{
            Status = [int]$webResponse.StatusCode
            Content = $content
            Headers = $webResponse.Headers
        }
    }
    return New-Object PSObject -Property @{
        Status = [int]$response.StatusCode
        Content = $response.Content
        Headers = $response.Headers
    }
}

function Invoke-Https {
    # curl.exe ships with Windows, so the HTTPS checks need no certificate trust setup.
    param([string]$Uri)
    $stamp = [Guid]::NewGuid().ToString("N")
    $bodyFile = Join-Path $env:TEMP ("zotero-test-body-" + $stamp)
    $headerFile = Join-Path $env:TEMP ("zotero-test-head-" + $stamp)
    $status = & curl.exe -sk --path-as-is -o $bodyFile -D $headerFile -w "%{http_code}" $Uri
    $content = ""
    if (Test-Path $bodyFile) { $content = Get-Content $bodyFile -Raw }
    $headers = New-Object System.Collections.Hashtable([System.StringComparer]::OrdinalIgnoreCase)
    if (Test-Path $headerFile) {
        foreach ($line in Get-Content $headerFile) {
            $colon = $line.IndexOf(":")
            if ($colon -gt 0) { $headers[$line.Substring(0, $colon).Trim()] = $line.Substring($colon + 1).Trim() }
        }
    }
    Remove-Item $bodyFile, $headerFile -Force -ErrorAction SilentlyContinue
    return New-Object PSObject -Property @{ Status = [int]$status; Content = "$content"; Headers = $headers }
}

try {
    # ------------------------------------------------------------------ setup
    New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
    $exePath = Join-Path $tempDir "ZoteroHelper.exe"
    Write-Host "Compiling $helperSource"
    & $csc /nologo /optimize+ /r:System.Web.Extensions.dll /out:"$exePath" "$helperSource"
    if (-not (Test-Path $exePath)) { throw "Compilation failed." }

    $certDir = Join-Path $tempDir "certs"
    New-Item -ItemType Directory -Force -Path $certDir | Out-Null
    $cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable -NotAfter (Get-Date).AddDays(30)
    $emptyPassword = New-Object System.Security.SecureString
    Export-PfxCertificate -Cert $cert -FilePath (Join-Path $certDir "localhost.pfx") -Password $emptyPassword | Out-Null
    Remove-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)" -Force -ErrorAction SilentlyContinue

    # The HTTPS checks shell out to curl.exe, so the test certificate needs no trust setup.

    # ------------------------------------------------------- stub Better BibTeX
    $stubScript = Join-Path $tempDir "stub.ps1"
    @'
param([int]$Port)
$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
$listener.Start()
while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
        $stream = $client.GetStream()
        $buffer = New-Object byte[] 16384
        $builder = New-Object System.Text.StringBuilder
        while ($true) {
            $read = $stream.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            [void]$builder.Append([System.Text.Encoding]::ASCII.GetString($buffer, 0, $read))
            if ($builder.ToString() -match "`r`n`r`n") { break }
        }
        $request = $builder.ToString()
        if ($request -match "/better-bibtex/cayw") {
            $body = '[{"citationKey":"smith2020","item":{"creators":[{"lastName":"Smith"}],"date":"2020"}}]'
        } elseif ($request -match "/better-bibtex/json-rpc") {
            $body = '{"jsonrpc":"2.0","result":"Smith, J. (2020). A title. Journal."}'
        } else {
            $body = '{"jsonrpc":"2.0","error":{"message":"unexpected endpoint"}}'
        }
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
        $head = "HTTP/1.1 200 OK`r`nContent-Type: application/json`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n"
        $headBytes = [System.Text.Encoding]::ASCII.GetBytes($head)
        $stream.Write($headBytes, 0, $headBytes.Length)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush()
    } finally {
        $client.Close()
    }
}
'@ | Set-Content -Path $stubScript -Encoding UTF8

    $stub = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $stubScript, "-Port", $stubPort -WindowStyle Hidden -PassThru
    $processes.Add($stub)

    Start-Sleep -Seconds 2

    # ------------------------------------------------------------ start helper
    $helper = Start-Process -FilePath $exePath -WorkingDirectory $tempDir -WindowStyle Hidden -PassThru -ArgumentList @(
        "--api-port", $apiPort, "--static-port", $staticPort,
        "--upstream", "http://127.0.0.1:$stubPort",
        "--www", $wwwSource, "--cert-dir", $certDir,
        "--log", (Join-Path $tempDir "helper.log")
    )
    $processes.Add($helper)

    # A second instance pointed at a closed port, to check the "Zotero is down" path.
    $orphan = Start-Process -FilePath $exePath -WorkingDirectory $tempDir -WindowStyle Hidden -PassThru -ArgumentList @(
        "--api-port", ($downPort), "--no-static", "--upstream", "http://127.0.0.1:$($BasePort + 9)",
        "--log", (Join-Path $tempDir "orphan.log")
    )
    $processes.Add($orphan)

    Start-Sleep -Seconds 3

    $api = "http://localhost:$apiPort"
    $site = "https://localhost:$staticPort"

    Write-Host "API"
    $health = Invoke-Endpoint -Uri "$api/health"
    Assert-Equal "GET /health status" 200 $health.Status
    Assert-Match "GET /health body" '"ok"' $health.Content
    Assert-Equal "GET /health CORS header" "*" $health.Headers["Access-Control-Allow-Origin"]

    $citations = Invoke-Endpoint -Uri "$api/zotero"
    Assert-Equal "GET /zotero status" 200 $citations.Status
    Assert-Match "GET /zotero body passes through" "smith2020" $citations.Content

    $bibliography = Invoke-Endpoint -Uri "$api/bibliography" -Method "POST" -Body '{"keys":["smith2020"],"style":"apa"}'
    Assert-Equal "POST /bibliography status" 200 $bibliography.Status
    Assert-Match "POST /bibliography body" "A title" $bibliography.Content

    $noKeys = Invoke-Endpoint -Uri "$api/bibliography" -Method "POST" -Body '{"keys":[],"style":"apa"}'
    Assert-Equal "POST /bibliography without keys" 400 $noKeys.Status
    Assert-Match "POST /bibliography without keys body" "No citation keys" $noKeys.Content

    $unknown = Invoke-Endpoint -Uri "$api/nope"
    Assert-Equal "GET /nope status" 404 $unknown.Status

    $down = Invoke-Endpoint -Uri "http://localhost:$downPort/zotero"
    Assert-Equal "GET /zotero with Zotero down" 500 $down.Status
    Assert-Match "GET /zotero with Zotero down body" "Better BibTeX" $down.Content

    Write-Host "Static files over HTTPS"
    $pane = Invoke-Https -Uri "$site/"
    Assert-Equal "GET / status" 200 $pane.Status
    Assert-Match "GET / serves the pane" "Zotero Citation Manager" $pane.Content
    Assert-Equal "GET / is not cached" "no-store" $pane.Headers["Cache-Control"]

    foreach ($path in @("/taskpane.html", "/frontend_core.js", "/style.css", "/commands.html", "/assets/icon-32.png")) {
        $file = Invoke-Https -Uri "$site$path"
        Assert-Equal "GET $path status" 200 $file.Status
    }

    $traversal = Invoke-Https -Uri "$site/../ZoteroHelper.exe"
    Assert-Equal "path traversal is blocked" 404 $traversal.Status

    Write-Host ""
    Write-Host ("Passed: " + $passed + "   Failed: " + $failures.Count)
} finally {
    foreach ($process in $processes) {
        if ($process -and -not $process.HasExited) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
    }
    # Only the instances this test started are stopped above; a helper installed for the add-in is left alone.
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failures.Count -gt 0) {
    Write-Host ""
    Write-Host "Failures:"
    foreach ($failure in $failures) { Write-Host ("  - " + $failure) }
    exit 1
}
exit 0
