Set-StrictMode -Version 2.0

# Moi tac vu co runspace rieng; UI chi doc log va gui co dung.
function Start-HddtDesktopRun {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][hashtable]$Values,
        [ValidateSet('download', 'local')][string]$Mode = 'download'
    )
    $context = [hashtable]::Synchronized(@{ StopRequested = $false; Completed = $false; ExitCode = 1; Error = '' })
    $copy = Copy-HddtConfigValues $Values
    $secrets = @()
    foreach ($key in @('GDT_PASSWORD', 'PROXY_PASSWORD')) {
        if ($copy.ContainsKey($key) -and -not [string]::IsNullOrEmpty($copy[$key])) { $secrets += $copy[$key] }
    }
    $entry = 'Invoke-Hddt.ps1'
    if ($Mode -eq 'local') { $entry = 'Parse-LocalXml.ps1' }
    $runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = 'MTA'
    $runspace.ThreadOptions = 'ReuseThread'
    $pipeline = [Management.Automation.PowerShell]::Create()
    try {
        $runspace.Open()
        $pipeline.Runspace = $runspace
        $worker = @'
param($Entry, $Values, $Context)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$global:ProgressPreference = 'SilentlyContinue'
$LASTEXITCODE = 0
try {
    & $Entry -Values $Values -RunContext $Context
    $Context['ExitCode'] = $LASTEXITCODE
}
catch {
    $Context['Error'] = $_.Exception.Message
    $Context['ExitCode'] = 1
}
finally {
    $Values.Clear()
    $Context['Completed'] = $true
}
'@
        $null = $pipeline.AddScript($worker).AddArgument((Join-Path $Root $entry)).AddArgument($copy).AddArgument($context)
        $handle = $pipeline.BeginInvoke()
        return [pscustomobject]@{
            Pipeline = $pipeline; Runspace = $runspace; Handle = $handle; Context = $context
            InformationIndex = 0; Secrets = $secrets; Values = $copy; Disposed = $false
        }
    }
    catch {
        $copy.Clear()
        $pipeline.Dispose()
        $runspace.Dispose()
        throw
    }
}

function Protect-HddtDesktopText {
    param([string]$Text, [string[]]$Secrets = @())
    $safe = Protect-ExcelErrorText $Text
    foreach ($secret in $Secrets) {
        if (-not [string]::IsNullOrEmpty($secret)) { $safe = $safe.Replace($secret, '[REDACTED]') }
    }
    return $safe
}

function Receive-HddtDesktopLog {
    param([Parameter(Mandatory = $true)]$Run)
    $stream = $Run.Pipeline.Streams.Information
    while ($Run.InformationIndex -lt $stream.Count) {
        $record = $stream[$Run.InformationIndex]
        $Run.InformationIndex++
        Protect-HddtDesktopText -Text ([string]$record.MessageData) -Secrets $Run.Secrets
    }
}

function Request-HddtDesktopStop {
    param([Parameter(Mandatory = $true)]$Run)
    $Run.Context['StopRequested'] = $true
}

function Complete-HddtDesktopRun {
    param([Parameter(Mandatory = $true)]$Run)
    if ($Run.Disposed) { return }
    if (-not $Run.Handle.IsCompleted) { throw 'Tac vu chua ket thuc; hay dung an toan va cho hoan tat.' }
    try {
        $null = $Run.Pipeline.EndInvoke($Run.Handle)
    }
    catch {
        $Run.Context['ExitCode'] = 1
        $Run.Context['Error'] = $_.Exception.Message
    }
    finally {
        $Run.Context['Error'] = Protect-HddtDesktopText -Text ([string]$Run.Context['Error']) -Secrets $Run.Secrets
        $Run.Values.Clear()
        $Run.Secrets = @()
        $Run.Pipeline.Dispose()
        $Run.Runspace.Dispose()
        $Run.Disposed = $true
    }
}

function Get-HddtDesktopProgress {
    param([string]$Line)
    if ($Line -match '\[(\d+)/(\d+)\s*\|\s*(\d+)%\]') {
        return [pscustomobject]@{ Current = [int]$Matches[1]; Total = [int]$Matches[2]; Percent = [int]$Matches[3] }
    }
    return $null
}
