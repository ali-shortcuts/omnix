# ============================================================================
# package.ps1 — stages the validated VSTO handoff into installer/payload/.
#
# The build step must first create build/compiled-payload/<host>. Packaging no
# longer reads transient src/<host>/bin/Release folders directly. This creates
# a strict boundary: source compilation -> validated immutable handoff -> setup.
#
# Release-safety invariants:
#   * Packaging MUST fail if any host DLL/.vsto/.dll.manifest is missing.
#   * OMNIX.Core.dll MUST be present in the staged payload.
#   * Installed payload identity MUST bind the exact Git source commit to the
#     exact Core + Excel + Word + PowerPoint assembly hashes shipped in setup.
#   * A successful source build is not enough; a hollow installer is a failure.
#   * Only the PUBLIC OMNIX.cer may be staged. No PFX/private key may enter the
#     installer payload or CI artifact.
# ============================================================================
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
$handoffRoot = Join-Path $PSScriptRoot 'compiled-payload'
$payload = Join-Path $root 'installer\payload'

if (-not (Test-Path -LiteralPath $handoffRoot -PathType Container)) {
    throw "COMPILED_HANDOFF_GUARD: validated build handoff is missing: $handoffRoot"
}

if (Test-Path $payload) { Remove-Item -Recurse -Force $payload }
New-Item -ItemType Directory -Force -Path $payload | Out-Null

$hosts = @('OMNIX.Excel', 'OMNIX.Word', 'OMNIX.PowerPoint')
$allowedExtensions = @('.dll', '.vsto', '.manifest', '.config')
$copiedSources = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)

function Copy-PayloadFile([string]$sourcePath) {
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required packaging input is missing: $sourcePath"
    }
    $item = Get-Item -LiteralPath $sourcePath
    if ($item.Length -le 0) {
        throw "Required packaging input is empty: $sourcePath"
    }
    Copy-Item -LiteralPath $item.FullName -Destination $payload -Force
    [void]$copiedSources.Add($item.FullName)
}

foreach ($hostProject in $hosts) {
    $sourceDir = Join-Path $handoffRoot $hostProject
    if (-not (Test-Path -LiteralPath $sourceDir -PathType Container)) {
        throw "COMPILED_HANDOFF_GUARD: host handoff directory missing: $sourceDir"
    }

    foreach ($requiredName in @(
        "$hostProject.dll",
        "$hostProject.dll.manifest",
        "$hostProject.vsto"
    )) {
        Copy-PayloadFile (Join-Path $sourceDir $requiredName)
    }

    $runtimeFiles = @(Get-ChildItem -LiteralPath $sourceDir -File -ErrorAction Stop | Where-Object {
        $allowedExtensions -contains $_.Extension
    })
    if ($runtimeFiles.Count -lt 3) {
        throw "COMPILED_HANDOFF_GUARD: too few runtime files for $hostProject ($($runtimeFiles.Count))."
    }
    foreach ($file in $runtimeFiles) {
        Copy-PayloadFile $file.FullName
    }
}

$requiredPayload = @(
    'OMNIX.Excel.dll',
    'OMNIX.Excel.dll.manifest',
    'OMNIX.Excel.vsto',
    'OMNIX.Word.dll',
    'OMNIX.Word.dll.manifest',
    'OMNIX.Word.vsto',
    'OMNIX.PowerPoint.dll',
    'OMNIX.PowerPoint.dll.manifest',
    'OMNIX.PowerPoint.vsto',
    'OMNIX.Core.dll'
)
foreach ($name in $requiredPayload) {
    $staged = Join-Path $payload $name
    if (-not (Test-Path -LiteralPath $staged -PathType Leaf)) {
        throw "HOLLOW_INSTALLER_GUARD: required staged payload file is missing: $name"
    }
    if ((Get-Item -LiteralPath $staged).Length -le 0) {
        throw "HOLLOW_INSTALLER_GUARD: staged payload file is empty: $name"
    }
}

$pfxFiles = @(Get-ChildItem (Join-Path $root 'build\cert') -Filter '*.pfx' -File -ErrorAction SilentlyContinue)
if ($pfxFiles.Count -gt 0) {
    throw 'Private signing key file(s) were found under build/cert. OMNIX packaging is fail-closed: remove PFX files and keep the private key in the Windows certificate store.'
}

$cert = Join-Path $root 'build\cert\OMNIX.cer'
if (Test-Path -LiteralPath $cert -PathType Leaf) {
    Copy-Item -LiteralPath $cert -Destination $payload -Force
    Write-Host 'OMNIX.cer staged (public certificate only).'
} else {
    Write-Warning 'No OMNIX.cer found — the installer will run without the development certificate trust helper.'
}

$requiredScripts = @(
    'post-install-verify.ps1',
    'install-vsto-trust.ps1',
    'classify-dev-cert.ps1',
    'office-registration-maintenance.ps1',
    'install-maintenance-task.ps1'
)
foreach ($scriptName in $requiredScripts) {
    $scriptPath = Join-Path $root ("build\" + $scriptName)
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        throw "$scriptName is required for installer verification/maintenance."
    }
    Copy-Item -LiteralPath $scriptPath -Destination $payload -Force
    Write-Host "$scriptName staged."
}

$payloadPrivateKeys = @(Get-ChildItem $payload -Recurse -Filter '*.pfx' -File -ErrorAction SilentlyContinue)
if ($payloadPrivateKeys.Count -gt 0) {
    throw 'Private key detected in installer payload. Packaging aborted.'
}

# Create a deterministic, installed build identity BEFORE the payload inventory is generated.
# This is the cryptographic source -> installed-payload bridge used by real-machine release evidence.
$sourceCommit = (& git -C $root rev-parse HEAD 2>$null).Trim().ToLowerInvariant()
if ($LASTEXITCODE -ne 0 -or $sourceCommit -notmatch '^[0-9a-f]{40}$') {
    throw 'PAYLOAD_IDENTITY_GUARD: could not resolve the exact Git source commit.'
}

function Payload-Hash([string]$name) {
    $path = Join-Path $payload $name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "PAYLOAD_IDENTITY_GUARD: required assembly missing: $name"
    }
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash.ToLowerInvariant()
}

$buildIdentity = [ordered]@{
    TestId = 'OMNIX-BUILD-IDENTITY-001'
    EvidenceSchema = 1
    SourceCommit = $sourceCommit
    CoreSha256 = Payload-Hash 'OMNIX.Core.dll'
    ExcelSha256 = Payload-Hash 'OMNIX.Excel.dll'
    WordSha256 = Payload-Hash 'OMNIX.Word.dll'
    PowerPointSha256 = Payload-Hash 'OMNIX.PowerPoint.dll'
    Scope = 'Exact source commit and primary managed assemblies staged into this installer payload.'
}
$buildIdentityPath = Join-Path $payload 'OMNIX-build-identity.json'
$buildIdentity | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $buildIdentityPath -Encoding ASCII
$identityCheck = Get-Content -LiteralPath $buildIdentityPath -Raw | ConvertFrom-Json
if ($identityCheck.TestId -ne 'OMNIX-BUILD-IDENTITY-001' -or
    [int]$identityCheck.EvidenceSchema -lt 1 -or
    [string]$identityCheck.SourceCommit -ne $sourceCommit -or
    [string]$identityCheck.CoreSha256 -ne (Payload-Hash 'OMNIX.Core.dll')) {
    throw 'PAYLOAD_IDENTITY_GUARD: generated build identity did not round-trip correctly.'
}
Write-Host "Installed payload identity created for source $sourceCommit."

$readmeFirst = Join-Path $payload 'README-first.txt'
@'
OMNIX — what to do next
=======================
1) Close and reopen Excel / Word / PowerPoint.
2) Look for the OMNIX ribbon entry supported by your Office host/version.
3) Click "Open Workspace" - the panel opens docked to the right of your document.
4) Open Settings inside the panel, choose a provider, paste your API key
   (it is stored encrypted with Windows DPAPI) and press "Test connection".

Office registration maintenance
-------------------------------
During your authorized installation OMNIX detects installed Excel, Word and PowerPoint hosts and
registers only its own per-user VSTO add-in keys. A LIMITED per-user maintenance task may also be
installed to re-scan supported Office hosts at sign-in, so a newly installed supported Office host
can receive the OMNIX registration without reinstalling OMNIX. The task never clears Office
Resiliency/DisabledItems, never changes Trust Center, never elevates, and never reads documents.
You can also run office-registration-maintenance.ps1 manually for repair/audit.

Privacy Mode default is "Ask before sending": before any request goes to a
cloud provider, OMNIX asks you once. "Local Only" keeps data on this PC.

Logs: %LOCALAPPDATA%\OMNIX\logs\
'@ | Set-Content -Path $readmeFirst -Encoding UTF8

$payloadFiles = @(Get-ChildItem -LiteralPath $payload -File | Sort-Object Name)
Write-Host "Staged $($copiedSources.Count) validated runtime files; payload contains $($payloadFiles.Count) files."
foreach ($file in $payloadFiles) {
    Write-Host ("  {0,-55} {1,12} bytes" -f $file.Name, $file.Length)
}

$inventoryDir = Join-Path $root 'build\artifact'
New-Item -ItemType Directory -Force -Path $inventoryDir | Out-Null
$inventory = foreach ($file in $payloadFiles) {
    [ordered]@{
        Name = $file.Name
        SizeBytes = [int64]$file.Length
        Sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $file.FullName).Hash.ToLowerInvariant()
    }
}
$inventory | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $inventoryDir 'payload-inventory.json') -Encoding UTF8

Write-Host 'OMNIX PAYLOAD GATE: PASS — all three Office hosts, OMNIX.Core, installed build identity and maintenance helpers are present.'
