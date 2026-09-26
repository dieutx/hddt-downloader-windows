Set-StrictMode -Version 2.0

# ---------------------------------------------------------------------------
# Tải + parse XML đa luồng.
#
# Chỉ giai đoạn tải/parse XML chạy song song; danh sách (phân trang theo state)
# và phần hóa đơn liên quan vẫn tuần tự. Mỗi hóa đơn được xử lý đúng một lần,
# kết quả mang theo Index và được áp lại theo đúng thứ tự danh sách, nên không
# sinh dòng trùng hay thiếu dù chạy bao nhiêu luồng.
# ---------------------------------------------------------------------------

# Xử lý tải + parse một hóa đơn. Không ném lỗi ra ngoài: mọi lỗi được đóng gói
# vào kết quả để luồng gọi quyết định ghi dòng lỗi hay bỏ qua.
function Invoke-HddtInvoiceXmlJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Invoice
    )

    $result = [pscustomobject]@{
        Ok = $false
        Stopped = $false
        XmlFiles = @()
        Parsed = @()
        Stage = 'Tải XML'
        Error = ''
        Endpoint = ''
        StatusCode = 0
        Attempts = 0
        RetryAfterSeconds = 0
    }

    try {
        Write-HddtLog DEBUG ('Bắt đầu hóa đơn: {0}' -f (Get-InvoiceLabel -Invoice $Invoice))
        $xmlFiles = @(Save-GdtInvoiceXml -Config $Config -Invoice $Invoice)
        $parsedList = New-Object System.Collections.Generic.List[object]
        foreach ($xmlFile in $xmlFiles) {
            $result.Stage = 'Parse XML'
            $parsedList.Add((ConvertFrom-InvoiceXml -Path $xmlFile -Direction $Invoice.Direction -Source $Invoice.Source))
        }
        $result.XmlFiles = $xmlFiles
        $result.Parsed = $parsedList.ToArray()
        $result.Ok = $true
    }
    catch {
        $result.Stopped = Test-HddtStopRequested
        $result.Error = $_.Exception.Message
        $result.Endpoint = Get-GdtLastRequestUri
        $result.StatusCode = Get-GdtLastStatusCode
        $result.Attempts = Get-GdtLastRequestAttempts
        $result.RetryAfterSeconds = Get-GdtLastRetryAfterSeconds
    }
    return $result
}

# Script chạy trong mỗi runspace công nhân: tự nạp module rồi vòng lấy hóa đơn
# từ chỉ số dùng chung cho tới khi hết việc hoặc bị yêu cầu dừng. Giới hạn số
# luồng hiện hành được đọc lại mỗi vòng để giảm/tăng theo phản hồi máy chủ.
$script:HddtXmlWorkerScript = @'
param($Gate, $Config, $Invoices, $Queue, $Root, $LogFile, $LogLevel, $LogToFile)

Set-StrictMode -Version 2.0
. (Join-Path $Root 'src\Logging.ps1')
. (Join-Path $Root 'src\Http.ps1')
. (Join-Path $Root 'src\XmlParser.ps1')
. (Join-Path $Root 'src\InvoiceApi.ps1')
. (Join-Path $Root 'src\Parallel.ps1')
. (Join-Path $Root 'src\Login.ps1')

$script:HddtSharedGate = $Gate
$script:HddtLogLevel = $LogLevel.ToUpperInvariant()
$script:HddtLogToFile = $LogToFile
$script:HddtLogFile = $LogFile
$script:HddtStartedUtc = [datetime]::UtcNow

$total = @($Invoices).Count
while (-not (Test-HddtStopRequested)) {
    $taken = $false
    $index = -1
    [System.Threading.Monitor]::Enter($Gate.Lock)
    try {
        if ([int]$Gate.NextIndex -ge $total) { break }
        if ([int]$Gate.ActiveWorkers -lt [int]$Gate.CurrentLimit) {
            $index = [int]$Gate.NextIndex
            $Gate.NextIndex = $index + 1
            $Gate.ActiveWorkers = [int]$Gate.ActiveWorkers + 1
            $taken = $true
        }
    }
    finally { [System.Threading.Monitor]::Exit($Gate.Lock) }

    if (-not $taken) { Start-Sleep -Milliseconds 60; continue }

    try {
        $invoice = $Invoices[$index]
        $result = Invoke-HddtInvoiceXmlJob -Config $Config -Invoice $invoice
    }
    catch {
        $result = [pscustomobject]@{
            Ok = $false
            Stopped = (Test-HddtStopRequested)
            XmlFiles = @()
            Parsed = @()
            Stage = 'Tải XML'
            Error = $_.Exception.Message
            Endpoint = ''
            StatusCode = 0
            Attempts = 0
            RetryAfterSeconds = 0
        }
    }
    finally {
        [System.Threading.Monitor]::Enter($Gate.Lock)
        try { $Gate.ActiveWorkers = [int]$Gate.ActiveWorkers - 1 }
        finally { [System.Threading.Monitor]::Exit($Gate.Lock) }
    }

    if ($null -ne $result) {
        Add-Member -InputObject $result -NotePropertyName Index -NotePropertyValue $index -Force
        # Queue không an toàn luồng nên mọi thao tác đều khóa bằng gate chung.
        [System.Threading.Monitor]::Enter($Gate.Lock)
        try { $Queue.Enqueue($result) }
        finally { [System.Threading.Monitor]::Exit($Gate.Lock) }
    }
}
'@

function Start-HddtXmlDownloadPool {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Invoices,
        [Parameter(Mandatory = $true)]$Gate,
        [Parameter(Mandatory = $true)][string]$Root,
        [string]$LogFile = '',
        [Parameter(Mandatory = $true)][string]$LogLevel,
        [Parameter(Mandatory = $true)][bool]$LogToFile,
        [Parameter(Mandatory = $true)][int]$WorkerCount
    )

    $queue = New-Object System.Collections.Queue
    $workers = New-Object System.Collections.Generic.List[object]
    $count = [Math]::Max(1, $WorkerCount)
    for ($i = 0; $i -lt $count; $i++) {
        $powershell = [System.Management.Automation.PowerShell]::Create()
        $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $runspace.Open()
        $powershell.Runspace = $runspace
        $null = $powershell.AddScript($script:HddtXmlWorkerScript)
        $null = $powershell.AddParameter('Gate', $Gate)
        $null = $powershell.AddParameter('Config', $Config)
        $null = $powershell.AddParameter('Invoices', $Invoices)
        $null = $powershell.AddParameter('Queue', $queue)
        $null = $powershell.AddParameter('Root', $Root)
        $null = $powershell.AddParameter('LogFile', $LogFile)
        $null = $powershell.AddParameter('LogLevel', $LogLevel)
        $null = $powershell.AddParameter('LogToFile', $LogToFile)
        $handle = $powershell.BeginInvoke()
        $workers.Add([pscustomobject]@{ PowerShell = $powershell; Runspace = $runspace; Handle = $handle })
    }
    return [pscustomobject]@{ Queue = $queue; Workers = $workers; Gate = $Gate }
}

# Lấy một kết quả từ hàng đợi dùng chung (khóa gate); trả $null khi hết.
function Get-HddtXmlPoolItem {
    param([Parameter(Mandatory = $true)]$Pool)
    $item = $null
    Lock-HddtGate $Pool.Gate
    try {
        if ($Pool.Queue.Count -gt 0) { $item = $Pool.Queue.Dequeue() }
    }
    finally { Unlock-HddtGate $Pool.Gate }
    return $item
}

function Complete-HddtXmlDownloadPool {
    param([Parameter(Mandatory = $true)]$Pool)
    foreach ($worker in $Pool.Workers) {
        try { $null = $worker.PowerShell.EndInvoke($worker.Handle) } catch { }
        try { $worker.PowerShell.Dispose() } catch { }
        try { $worker.Runspace.Dispose() } catch { }
    }
}
