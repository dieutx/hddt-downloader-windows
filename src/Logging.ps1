Set-StrictMode -Version 2.0

$script:HddtLogFile = $null
$script:HddtLogLevel = 'INFO'
$script:HddtLogToFile = $false
$script:HddtLevelRank = @{ DEBUG = 10; INFO = 20; WARN = 30; ERROR = 40 }
$script:HddtStartedUtc = [datetime]::UtcNow

# Trạng thái dùng chung giữa main thread và các worker runspace (tải XML song
# song). Logging.ps1 là file nền nên được nạp trước tiên; các logic điều tiết
# và pipeline nằm ở src/XmlScheduler.ps1.
$script:HddtSharedState = $null
# Khi bật, worker chỉ đẩy log vào shared.LogQueue; main thread là nơi duy nhất
# ghi ra console/file để log không bị đan xen giữa các runspace.
$script:HddtLogForwardToShared = $false

function Set-HddtSharedState {
    param($Shared)
    $script:HddtSharedState = $Shared
}

function Get-HddtSharedState {
    return $script:HddtSharedState
}

# Đọc một khóa của trạng thái dùng chung; trả về $Default khi chưa có pipeline.
function Get-HddtSharedValue {
    param([string]$Key, $Default = $null)
    $shared = $script:HddtSharedState
    if ($null -eq $shared) { return $Default }
    if (-not $shared.ContainsKey($Key)) { return $Default }
    return $shared[$Key]
}

# Ghi một khóa; mọi read-modify-write (tăng/giảm đếm, điều tiết) phải tự giữ
# lock bằng Monitor.Enter/Exit trên shared.SyncRoot ngay trong chỗ gọi.
function Set-HddtSharedValue {
    param([string]$Key, $Value)
    $shared = $script:HddtSharedState
    if ($null -eq $shared) { return }
    [Threading.Monitor]::Enter($shared.SyncRoot)
    try { $shared[$Key] = $Value }
    finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
}

# Thời gian tính từ lúc bắt đầu phiên, dùng để ghi thời lượng vào log INFO.
function Get-HddtElapsedText {
    $elapsed = [datetime]::UtcNow - $script:HddtStartedUtc
    if ($elapsed.TotalSeconds -lt 60) { return ('{0:0.0}s' -f $elapsed.TotalSeconds) }
    return ('{0:hh\:mm\:ss}' -f $elapsed)
}

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
    $script:HddtStartedUtc = [datetime]::UtcNow
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

    if ($script:HddtLogForwardToShared -and $null -ne $script:HddtSharedState) {
        $shared = $script:HddtSharedState
        [Threading.Monitor]::Enter($shared.SyncRoot)
        try { $shared.LogQueue.Add([pscustomobject]@{ Level = $normalizedLevel; Message = $Message }) }
        finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
        return
    }

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
