Set-StrictMode -Version 2.0

# Báo cho src/Http.ps1 biết bộ điều tiết XML đã sẵn sàng: request ExportXml
# sẽ đi qua Enter/Exit-GdtXmlRequestSlot thay vì giãn cách tổng (REQUEST_DELAY_MS).
$script:HddtXmlThrottleAvailable = $true

# Ngưỡng phục hồi sau rate-limit: sau 25 request thành công liên tiếp mới giảm
# giãn cách hoặc tăng lại số kết nối.
$script:HddtXmlRecoveryThreshold = 25

# Script chạy trong mỗi worker runspace: tự nạp src/*, nối vào trạng thái dùng
# chung rồi xử lý các hóa đơn được chia vòng tròn (round-robin).
$script:HddtXmlWorkerScript = @'
param($Root, $Shared, $Config, $WorkerIndex, $WorkerCount, $LogLevel, $Invoices)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$global:ProgressPreference = 'SilentlyContinue'
. (Join-Path $Root 'src/Logging.ps1')
. (Join-Path $Root 'src/Config.ps1')
. (Join-Path $Root 'src/Http.ps1')
. (Join-Path $Root 'src/InvoiceApi.ps1')
. (Join-Path $Root 'src/Login.ps1')
. (Join-Path $Root 'src/BrowserProfile.ps1')
. (Join-Path $Root 'src/XmlScheduler.ps1')
Set-HddtSharedState -Shared $Shared
$script:HddtLogForwardToShared = $true
$script:HddtLogLevel = ([string]$LogLevel).ToUpperInvariant()
$script:HddtLogToFile = $false
$script:HddtLogFile = $null

$workerIndex = [int]$WorkerIndex
$workerCount = [int]$WorkerCount
while ($workerIndex -lt $Invoices.Count) {
    if (Test-HddtStopRequested) { break }
    $invoice = $Invoices[$workerIndex]
    try {
        $result = Get-GdtXmlDownloadResult -Config $Config -Invoice $invoice -PipelineIndex $workerIndex
    }
    catch {
        # Get-GdtXmlDownloadResult đã bọc mọi lỗi; nhánh này chỉ phòng khi
        # chính hàm đó trục trặc để main thread vẫn nhận được một kết quả.
        $result = [pscustomobject]@{
            PipelineIndex = $workerIndex
            Invoice = $invoice
            Label = ''
            Success = $false
            Stage = 'Tải XML'
            XmlFiles = @()
            ResponseTimeMs = -1
            Endpoint = Get-GdtLastRequestUri
            StatusCode = Get-GdtLastStatusCode
            Attempts = Get-GdtLastRequestAttempts
            RetryAfterSeconds = Get-GdtLastRetryAfterSeconds
            ErrorMessage = $_.Exception.Message
        }
    }
    Add-HddtXmlResult -Result $result
    $workerIndex += $workerCount
}
'@

# Khởi tạo trạng thái dùng chung cho một lần tải XML song song và nối vào
# phạm vi script hiện tại để Http/Logging thấy cùng một đối tượng.
function New-HddtXmlSharedState {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Config)

    $concurrency = [int](Get-HddtConfigValue $Config 'XmlConcurrency' 1)
    $maxConcurrency = [int](Get-HddtConfigValue $Config 'XmlMaxConcurrency' 1)
    $intervalMs = [int](Get-HddtConfigValue $Config 'XmlRequestIntervalMs' 0)
    if ($maxConcurrency -lt 1) { $maxConcurrency = 1 }
    if ($concurrency -lt 1) { $concurrency = 1 }
    if ($concurrency -gt $maxConcurrency) { $concurrency = $maxConcurrency }

    $shared = [hashtable]::Synchronized(@{
        SyncRoot = (New-Object object)
        StopRequested = $false
        Token = [string](Get-HddtConfigValue $Config 'Token' '')
        AuthRefreshCount = 0
        Results = (New-Object System.Collections.Generic.List[object])
        LogQueue = (New-Object System.Collections.Generic.List[object])
        XmlBaseConcurrency = $concurrency
        XmlMaxConcurrency = $maxConcurrency
        XmlCurrentConcurrency = $concurrency
        XmlBaseIntervalMs = $intervalMs
        XmlCurrentIntervalMs = $intervalMs
        XmlSuccessStreak = 0
        XmlRateLimitCount = 0
        XmlGlobalCooldownUntilUtc = [datetime]::MinValue
        XmlLastRequestStartUtc = [datetime]::MinValue
        XmlActiveCount = 0
        XmlPeakConcurrency = 0
        XmlCompletedCount = 0
        XmlTotalResponseTimeMs = 0
        CooldownFallbackSeconds = 15
        StartUtc = [datetime]::UtcNow
    })
    Set-HddtSharedState -Shared $shared
    return $shared
}

# Chặn cho tới khi được phép gửi một request XML: hết cooldown rate-limit,
# số kết nối đang mở dưới mức cho phép và đã qua giãn cách giữa hai request.
# -TryOnly: không chờ, trả về $false nếu chưa lấy được slot (dùng cho kiểm
# thử số kết nối mà không cần ngủ thật).
function Enter-GdtXmlRequestSlot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [switch]$TryOnly
    )
    $shared = Get-HddtSharedState
    if ($null -eq $shared) { return $true }
    $lock = $shared.SyncRoot

    while ($true) {
        if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; không gửi request mới.' }
        $waitMs = 0
        $now = [datetime]::UtcNow
        [Threading.Monitor]::Enter($lock)
        try {
            $cooldownUntil = [datetime]$shared.XmlGlobalCooldownUntilUtc
            if ($now -lt $cooldownUntil) {
                $waitMs = [int][Math]::Ceiling(($cooldownUntil - $now).TotalMilliseconds)
            }
            elseif ([int]$shared.XmlActiveCount -ge [int]$shared.XmlCurrentConcurrency) {
                # Đã dùng hết số kết nối: chờ worker khác trả slot về.
                $waitMs = 100
            }
            else {
                $lastStart = [datetime]$shared.XmlLastRequestStartUtc
                $interval = [int]$shared.XmlCurrentIntervalMs
                if ($lastStart -ne [datetime]::MinValue -and $interval -gt 0) {
                    $remaining = [int][Math]::Ceiling($interval - ($now - $lastStart).TotalMilliseconds)
                    if ($remaining -gt 0) { $waitMs = $remaining }
                }
                if ($waitMs -le 0) {
                    $shared.XmlActiveCount = ([int]$shared.XmlActiveCount + 1)
                    $shared.XmlLastRequestStartUtc = $now
                    if ([int]$shared.XmlActiveCount -gt [int]$shared.XmlPeakConcurrency) {
                        $shared.XmlPeakConcurrency = [int]$shared.XmlActiveCount
                    }
                    return $true
                }
            }
        }
        finally { [Threading.Monitor]::Exit($lock) }

        if ($TryOnly) { return $false }
        if ($waitMs -gt 500) { $waitMs = 500 }
        Start-Sleep -Milliseconds $waitMs
    }
}

# Trả slot sau khi request kết thúc. ResponseTimeMs < 0 nghĩa là request thất
# bại nên không tính vào thời gian phản hồi trung bình.
function Exit-GdtXmlRequestSlot {
    [CmdletBinding()]
    param([int]$ResponseTimeMs = -1)
    $shared = Get-HddtSharedState
    if ($null -eq $shared) { return }
    [Threading.Monitor]::Enter($shared.SyncRoot)
    try {
        if ([int]$shared.XmlActiveCount -gt 0) { $shared.XmlActiveCount = ([int]$shared.XmlActiveCount - 1) }
        if ($ResponseTimeMs -ge 0) {
            $shared.XmlCompletedCount = ([int]$shared.XmlCompletedCount + 1)
            $shared.XmlTotalResponseTimeMs = ([int]$shared.XmlTotalResponseTimeMs + $ResponseTimeMs)
        }
    }
    finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
}

# HTTP 429 với request XML: đặt cooldown, giảm một kết nối và tăng giãn cách.
# Không jitter khi server gửi Retry-After; không có Retry-After thì dùng
# CooldownFallbackSeconds ±2s.
function Register-GdtXmlRateLimit {
    [CmdletBinding()]
    param([int]$RetryAfterSeconds = 0)
    $shared = Get-HddtSharedState
    if ($null -eq $shared) { return }

    $cooldownSeconds = 0
    $oldConcurrency = 0
    $newConcurrency = 0
    $oldInterval = 0
    $newInterval = 0
    [Threading.Monitor]::Enter($shared.SyncRoot)
    try {
        if ($RetryAfterSeconds -gt 0) {
            $cooldownSeconds = $RetryAfterSeconds
        }
        else {
            $fallback = [int]$shared.CooldownFallbackSeconds
            # Không có Retry-After thì dùng fallback ±2s; fallback = 0 nghĩa là
            # không chờ (dùng cho kiểm thử và cấu hình không muốn chặn).
            if ($fallback -gt 0) { $cooldownSeconds = [Math]::Max(0, $fallback + (Get-Random -Minimum -2 -Maximum 3)) }
        }
        $proposedCooldownUntilUtc = ([datetime]::UtcNow).AddSeconds($cooldownSeconds)
        $existingCooldownUntilUtc = [datetime]$shared.XmlGlobalCooldownUntilUtc
        if ($proposedCooldownUntilUtc -gt $existingCooldownUntilUtc) {
            $shared.XmlGlobalCooldownUntilUtc = $proposedCooldownUntilUtc
        }

        $oldConcurrency = [int]$shared.XmlCurrentConcurrency
        $newConcurrency = [Math]::Max(1, $oldConcurrency - 1)
        $shared.XmlCurrentConcurrency = $newConcurrency

        $oldInterval = [int]$shared.XmlCurrentIntervalMs
        $newInterval = [Math]::Min(5000, [Math]::Max([int]$shared.XmlBaseIntervalMs, [int][Math]::Floor($oldInterval * 1.5)))
        $shared.XmlCurrentIntervalMs = $newInterval

        $shared.XmlSuccessStreak = 0
        $shared.XmlRateLimitCount = ([int]$shared.XmlRateLimitCount + 1)
    }
    finally { [Threading.Monitor]::Exit($shared.SyncRoot) }

    Write-HddtLog WARN ('[XML THROTTLE] HTTP 429 | cooldown {0}s | concurrency {1}→{2} | interval {3}→{4} ms' -f $cooldownSeconds, $oldConcurrency, $newConcurrency, $oldInterval, $newInterval)
}

# Request XML thành công liên tiếp: sau ngưỡng phục hồi thì giảm giãn cách
# trước, hết giãn cách mới tăng lại số kết nối (không lấy lại quá mức cấu hình).
function Register-GdtXmlSuccess {
    [CmdletBinding()]
    param()
    $shared = Get-HddtSharedState
    if ($null -eq $shared) { return }

    $logMessage = ''
    [Threading.Monitor]::Enter($shared.SyncRoot)
    try {
        $shared.XmlSuccessStreak = ([int]$shared.XmlSuccessStreak + 1)
        if ([int]$shared.XmlSuccessStreak -lt $script:HddtXmlRecoveryThreshold) { return }
        $shared.XmlSuccessStreak = 0

        $interval = [int]$shared.XmlCurrentIntervalMs
        $baseInterval = [int]$shared.XmlBaseIntervalMs
        $concurrency = [int]$shared.XmlCurrentConcurrency
        $maxConcurrency = [int]$shared.XmlMaxConcurrency
        if ($interval -gt $baseInterval) {
            $newInterval = [Math]::Max($baseInterval, [int][Math]::Floor($interval * 0.85))
            if ($newInterval -lt $interval) {
                $shared.XmlCurrentIntervalMs = $newInterval
                $logMessage = ('25 request thành công | interval {0}→{1} ms' -f $interval, $newInterval)
            }
        }
        elseif ($concurrency -lt $maxConcurrency) {
            $shared.XmlCurrentConcurrency = ($concurrency + 1)
            $logMessage = ('Kết nối ổn định | concurrency {0}→{1}' -f $concurrency, ($concurrency + 1))
        }
    }
    finally { [Threading.Monitor]::Exit($shared.SyncRoot) }

    if (-not [string]::IsNullOrEmpty($logMessage)) {
        Write-HddtLog INFO ('[XML THROTTLE] {0}' -f $logMessage)
    }
}

# Số liệu hiện tại của bộ điều tiết, dùng cho log thống kê và kiểm thử.
function Get-GdtXmlThrottleSnapshot {
    $shared = Get-HddtSharedState
    if ($null -eq $shared) { return $null }
    return [pscustomobject]@{
        BaseConcurrency = [int]$shared.XmlBaseConcurrency
        MaxConcurrency = [int]$shared.XmlMaxConcurrency
        CurrentConcurrency = [int]$shared.XmlCurrentConcurrency
        BaseIntervalMs = [int]$shared.XmlBaseIntervalMs
        CurrentIntervalMs = [int]$shared.XmlCurrentIntervalMs
        ActiveCount = [int]$shared.XmlActiveCount
        PeakConcurrency = [int]$shared.XmlPeakConcurrency
        CompletedCount = [int]$shared.XmlCompletedCount
        TotalResponseTimeMs = [int]$shared.XmlTotalResponseTimeMs
        RateLimitCount = [int]$shared.XmlRateLimitCount
        SuccessStreak = [int]$shared.XmlSuccessStreak
        AuthRefreshCount = [int]$shared.AuthRefreshCount
    }
}

function Add-HddtXmlResult {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Result)
    $shared = Get-HddtSharedState
    if ($null -eq $shared) { return }
    [Threading.Monitor]::Enter($shared.SyncRoot)
    try { $shared.Results.Add($Result) }
    finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
}

# Lấy toàn bộ kết quả đang chờ và làm rỗng hàng đợi (main thread gọi dần).
function Receive-HddtXmlResults {
    $shared = Get-HddtSharedState
    if ($null -eq $shared) { return @() }
    $items = $null
    [Threading.Monitor]::Enter($shared.SyncRoot)
    try {
        $items = $shared.Results.ToArray()
        $shared.Results.Clear()
    }
    finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
    return @($items)
}

# Lấy log các worker đẩy về và làm rỗng hàng đợi; chỉ main thread ghi file.
function Receive-HddtXmlLogEntries {
    $shared = Get-HddtSharedState
    if ($null -eq $shared) { return @() }
    $items = $null
    [Threading.Monitor]::Enter($shared.SyncRoot)
    try {
        $items = $shared.LogQueue.ToArray()
        $shared.LogQueue.Clear()
    }
    finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
    return @($items)
}

function Write-HddtSharedLogEntries {
    foreach ($entry in @(Receive-HddtXmlLogEntries)) {
        Write-HddtLog -Level ([string]$entry.Level) -Message ([string]$entry.Message)
    }
}

# Tải XML cho một hóa đơn trong worker: mọi lỗi được trả về dạng kết quả để
# main thread ghi đúng dòng lỗi; không ném lỗi ra ngoài vòng lặp worker.
function Get-GdtXmlDownloadResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Invoice,
        [Parameter(Mandatory = $true)][int]$PipelineIndex
    )

    $label = Get-InvoiceLabel -Invoice $Invoice
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try {
        Write-HddtLog DEBUG ('Bắt đầu hóa đơn {0}: {1} [{2}]' -f ($PipelineIndex + 1), $label, $Invoice.Source)
        $xmlFiles = @(Save-GdtInvoiceXml -Config $Config -Invoice $Invoice)
        $watch.Stop()
        return [pscustomobject]@{
            PipelineIndex = $PipelineIndex
            Invoice = $Invoice
            Label = $label
            Success = $true
            Stage = ''
            XmlFiles = $xmlFiles
            ResponseTimeMs = [int]$watch.ElapsedMilliseconds
            Endpoint = Get-GdtLastRequestUri
            StatusCode = Get-GdtLastStatusCode
            Attempts = Get-GdtLastRequestAttempts
            RetryAfterSeconds = Get-GdtLastRetryAfterSeconds
            ErrorMessage = ''
        }
    }
    catch {
        $watch.Stop()
        return [pscustomobject]@{
            PipelineIndex = $PipelineIndex
            Invoice = $Invoice
            Label = $label
            Success = $false
            Stage = 'Tải XML'
            XmlFiles = @()
            ResponseTimeMs = [int]$watch.ElapsedMilliseconds
            Endpoint = Get-GdtLastRequestUri
            StatusCode = Get-GdtLastStatusCode
            Attempts = Get-GdtLastRequestAttempts
            RetryAfterSeconds = Get-GdtLastRetryAfterSeconds
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Mở các worker runspace; mỗi worker xử lý hóa đơn có chỉ số
# workerIndex, workerIndex+workerCount, ...
function Start-HddtXmlPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Invoices,
        [Parameter(Mandatory = $true)][string]$Root
    )

    $shared = Get-HddtSharedState
    if ($null -eq $shared) { throw 'Chưa khởi tạo trạng thái dùng chung; gọi New-HddtXmlSharedState trước.' }

    $workers = New-Object System.Collections.Generic.List[object]
    $workerCount = 0
    if ($Invoices.Count -gt 0) {
        $workerCount = [Math]::Min([int]$shared.XmlMaxConcurrency, $Invoices.Count)
        $logLevel = [string](Get-HddtConfigValue $Config 'LogLevel' 'info')
        try {
            for ($workerIndex = 0; $workerIndex -lt $workerCount; $workerIndex++) {
                $runspace = [runspacefactory]::CreateRunspace()
                $runspace.Open()
                $powerShell = [powershell]::Create()
                $powerShell.Runspace = $runspace
                $null = $powerShell.AddScript($script:HddtXmlWorkerScript)
                $null = $powerShell.AddArgument($Root)
                $null = $powerShell.AddArgument($shared)
                $null = $powerShell.AddArgument($Config)
                $null = $powerShell.AddArgument($workerIndex)
                $null = $powerShell.AddArgument($workerCount)
                $null = $powerShell.AddArgument($logLevel)
                $null = $powerShell.AddArgument($Invoices)
                $handle = $powerShell.BeginInvoke()
                $workers.Add([pscustomobject]@{
                    PowerShell = $powerShell
                    Runspace = $runspace
                    Handle = $handle
                    Index = $workerIndex
                })
            }
        }
        catch {
            # Không mở được đủ worker: giải phóng runspace đã mở rồi báo lỗi,
            # tránh treo kết nối của phiên chạy.
            foreach ($worker in @($workers)) {
                try { $worker.PowerShell.Dispose() } catch { }
                try { $worker.Runspace.Close(); $worker.Runspace.Dispose() } catch { }
            }
            throw
        }
    }

    return [pscustomobject]@{
        Shared = $shared
        Workers = $workers.ToArray()
        Root = $Root
        Invoices = $Invoices
    }
}

# Đẩy log của worker ra console/file của main thread và nhận kết quả mới.
function Receive-HddtXmlPipeline {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Pipeline)
    Write-HddtSharedLogEntries
    return @(Receive-HddtXmlResults)
}

function Test-HddtXmlPipelineCompleted {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Pipeline)
    foreach ($worker in @($Pipeline.Workers)) {
        if (-not $worker.Handle.IsCompleted) { return $false }
    }
    return $true
}

# Dừng pipeline (an toàn khi đã chạy xong): báo cờ dừng nếu còn worker đang
# chạy, chờ chúng thoát, rồi lấy nốt log/kết quả và giải phóng runspace.
# Trả về các kết quả nhận được trong lúc dừng để main thread ghi tiếp.
function Stop-HddtXmlPipeline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Pipeline,
        [int]$TimeoutSeconds = 60
    )

    if (-not (Test-HddtXmlPipelineCompleted -Pipeline $Pipeline)) {
        # Không đặt cờ khi pipeline đã chạy xong: main thread sẽ tưởng người
        # dùng Ctrl+C và bỏ qua các bước còn lại.
        Set-HddtSharedValue -Key 'StopRequested' -Value $true
        $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
        while (-not (Test-HddtXmlPipelineCompleted -Pipeline $Pipeline) -and [datetime]::UtcNow -lt $deadline) {
            Start-Sleep -Milliseconds 100
        }
    }

    $remaining = @()
    foreach ($worker in @($Pipeline.Workers)) {
        try {
            $null = $worker.PowerShell.EndInvoke($worker.Handle)
        }
        catch {
            Write-HddtLog ERROR ('[TẢI XML] Worker {0} kết thúc bất thường: {1}' -f $worker.Index, $_.Exception.Message)
        }
        foreach ($streamError in @($worker.PowerShell.Streams.Error)) {
            Write-HddtLog ERROR ('[TẢI XML] Worker {0}: {1}' -f $worker.Index, $streamError.ToString())
        }
        try { $worker.PowerShell.Dispose() } catch { }
        try { $worker.Runspace.Close(); $worker.Runspace.Dispose() } catch { }
    }
    $Pipeline.Workers = @()

    Write-HddtSharedLogEntries
    $remaining = @(Receive-HddtXmlResults)
    return $remaining
}

# Chốt trạng thái: đưa token đã làm mới về object cấu hình và trả số liệu.
function Complete-HddtSharedState {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Config)

    $shared = Get-HddtSharedState
    if ($null -eq $shared) { return $null }

    $token = [string](Get-HddtSharedValue -Key 'Token' -Default '')
    if (-not [string]::IsNullOrWhiteSpace($token)) { Set-GdtAuthToken -Config $Config -Token $token }

    $snapshot = Get-GdtXmlThrottleSnapshot
    $averageResponseTimeMs = 0
    if ($null -ne $snapshot -and $snapshot.CompletedCount -gt 0) {
        $averageResponseTimeMs = [int][Math]::Round($snapshot.TotalResponseTimeMs / [double]$snapshot.CompletedCount)
    }
    $durationMs = [int](([datetime]::UtcNow - [datetime]$shared.StartUtc).TotalMilliseconds)
    $peakConcurrency = 0
    $completedCount = 0
    $rateLimitCount = 0
    $authRefreshCount = 0
    if ($null -ne $snapshot) {
        $peakConcurrency = $snapshot.PeakConcurrency
        $completedCount = $snapshot.CompletedCount
        $rateLimitCount = $snapshot.RateLimitCount
        $authRefreshCount = $snapshot.AuthRefreshCount
    }
    return [pscustomobject]@{
        PeakConcurrency = $peakConcurrency
        CompletedCount = $completedCount
        RateLimitCount = $rateLimitCount
        AuthRefreshCount = $authRefreshCount
        AverageResponseTimeMs = $averageResponseTimeMs
        DurationMs = $durationMs
    }
}
