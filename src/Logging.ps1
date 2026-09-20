Set-StrictMode -Version 2.0

$script:HddtLogFile = $null
$script:HddtLogLevel = 'INFO'
$script:HddtLogToFile = $false
$script:HddtLevelRank = @{ DEBUG = 10; INFO = 20; WARN = 30; ERROR = 40 }

function Initialize-HddtConsole {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    try { [Console]::InputEncoding = $utf8 } catch { }
    try { [Console]::OutputEncoding = $utf8 } catch { }
    $global:OutputEncoding = $utf8

    # Windows PowerShell 5.1 vẽ lại thanh tiến trình của Invoke-WebRequest liên
    # tục. Tắt UI động để terminal chỉ thêm dòng mới và không bị nhấp nháy.
    $global:ProgressPreference = 'SilentlyContinue'
    try { $Host.UI.RawUI.WindowTitle = 'HDDT Downloader for Windows' } catch { }
}

function Start-HddtLogging {
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        [ValidateSet('debug', 'info', 'warn', 'error')][string]$Level = 'info',
        [bool]$ToFile = $true,
        [string]$RunName = 'run'
    )

    $script:HddtLogLevel = $Level.ToUpperInvariant()
    $script:HddtLogToFile = $ToFile
    if ($ToFile) {
        New-Item -ItemType Directory -Path $Directory -Force | Out-Null
        $script:HddtLogFile = Join-Path $Directory ('{0}_{1}.log' -f $RunName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
        [IO.File]::WriteAllText($script:HddtLogFile, '', (New-Object Text.UTF8Encoding($true)))
    }
    return $script:HddtLogFile
}

function Write-HddtLog {
    param(
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level = 'INFO',
        [Parameter(Mandatory = $true)][string]$Message
    )

    $normalizedLevel = $Level.ToUpperInvariant()
    if ($script:HddtLevelRank[$normalizedLevel] -lt $script:HddtLevelRank[$script:HddtLogLevel]) { return }

    $consoleLine = '[{0}] [{1,-5}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $normalizedLevel, $Message
    switch ($normalizedLevel) {
        'WARN' { Write-Host $consoleLine -ForegroundColor Yellow }
        'ERROR' { Write-Host $consoleLine -ForegroundColor Red }
        default { Write-Host $consoleLine }
    }

    if ($script:HddtLogToFile -and -not [string]::IsNullOrWhiteSpace($script:HddtLogFile)) {
        $fileLine = '[{0}] [{1,-5}] {2}{3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $normalizedLevel, $Message, [Environment]::NewLine
        [IO.File]::AppendAllText($script:HddtLogFile, $fileLine, (New-Object Text.UTF8Encoding($false)))
    }
}

function Get-HddtLogFile {
    return $script:HddtLogFile
}
