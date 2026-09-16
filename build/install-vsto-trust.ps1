[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$InstallDir,[Parameter(Mandatory=$true)][string]$HostNames,[switch]$Silent)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$logRoot=Join-Path $env:LOCALAPPDATA 'OMNIX\logs'
New-Item -ItemType Directory -Force $logRoot | Out-Null
try {
    if(Get-Process EXCEL,WINWORD,POWERPNT -ErrorAction SilentlyContinue){throw 'Close Office before installing add-ins.'}
    # Microsoft's deployment installer owns its trust prompt. Never import a root certificate.
    $candidates=@()
    foreach($view in @([Microsoft.Win32.RegistryView]::Registry64,[Microsoft.Win32.RegistryView]::Registry32)) {
        $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$view)
        try {
            foreach($version in @('v4','v4R')) {
                $key=$base.OpenSubKey('SOFTWARE\Microsoft\VSTO Runtime Setup\'+$version)
                if($null -ne $key){try{$candidate=[string]$key.GetValue('InstallerPath');if($candidate){$candidates+=$candidate}}finally{$key.Dispose()}}
            }
        }finally{$base.Dispose()}
    }
    foreach($folder in @($env:CommonProgramW6432,$env:CommonProgramFiles,${env:CommonProgramFiles(x86)})) {
        if($folder){$candidates+=(Join-Path $folder 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe')}
    }
    $installer=$candidates | Where-Object {Test-Path -LiteralPath $_ -PathType Leaf} | Select-Object -First 1
    if(-not $installer){throw 'Microsoft VSTOInstaller.exe was not found. Repair the VSTO Runtime.'}
    foreach($officeName in $HostNames.Split(',')) {
        if($officeName -notin @('Excel','Word','PowerPoint')){throw 'Unsupported Office host.'}
        $manifest=Join-Path $InstallDir ("OMNIX.$officeName.vsto")
        if(-not(Test-Path -LiteralPath $manifest -PathType Leaf)){throw "Missing $officeName manifest."}
        $arguments='/Install "'+$manifest+'"'
        if($Silent){$arguments+=' /Silent'}
        $process=Start-Process -FilePath $installer -ArgumentList $arguments -Wait -PassThru
        if($process.ExitCode -ne 0){throw "Microsoft VSTO deployment for $officeName returned $($process.ExitCode). Review its deployment/trust error; security policy was not changed."}
    }
    'Microsoft VSTO deployment succeeded.' | Set-Content (Join-Path $logRoot 'vsto-deployment.log')
    exit 0
}catch {
    $_.Exception.Message | Set-Content (Join-Path $logRoot 'vsto-deployment.log')
    Write-Error $_ -ErrorAction Continue
    exit 10
}
