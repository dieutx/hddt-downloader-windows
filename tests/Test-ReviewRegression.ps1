Set-StrictMode -Version 2.0

# Cac ca loi tim thay khi review; chi dung du lieu gia va file tam.
$reviewFailures = New-Object 'Collections.Generic.List[string]'
function Test-ReviewCase {
    param([string]$Name, [scriptblock]$Body)
    try { & $Body }
    catch { $reviewFailures.Add($Name + ': ' + $_.Exception.Message) }
}

Test-ReviewCase 'No-overwrite race' {
    $path = Join-Path ([IO.Path]::GetTempPath()) ('hddt-review-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    $validator = ${function:Assert-ExcelPackage}
    try {
        $script:ReviewRacePath = $path
        Set-Item Function:\Assert-ExcelPackage -Value {
            param($Path, $Parts)
            [IO.File]::WriteAllText($script:ReviewRacePath, 'concurrent-user-file')
        }
        $rejected = $false
        try { $null = Export-InvoiceWorkbook -Path $path -SummaryRows @() -DetailRows @() -ErrorRows @() }
        catch { $rejected = $true }
        Assert-Equal $true $rejected 'File appearing during export must not be overwritten'
        Assert-Equal 'concurrent-user-file' ([IO.File]::ReadAllText($path)) 'Concurrent file stays unchanged'
    }
    finally {
        Set-Item Function:\Assert-ExcelPackage -Value $validator
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
}

Test-ReviewCase 'Local error rows' {
    $errors = @(
        [pscustomobject]@{ Source = 'local-xml'; Direction = 'purchase'; Stage = 'Parse XML'; Endpoint = '/fake/a.xml'; Error = 'bad a' }
        [pscustomobject]@{ Source = 'local-xml'; Direction = 'purchase'; Stage = 'Parse XML'; Endpoint = '/fake/b.xml'; Error = 'bad b' }
    )
    $rows = @(New-ExcelErrorRows $errors)
    Assert-Equal 2 $rows.Count 'Different local XML errors must remain separate'
    Assert-Equal 2 @(New-ExcelErrorRows @($errors + $errors)).Count 'Repeated errors for the same local file still deduplicate'
}

Test-ReviewCase 'Missing tax rate' {
    Assert-Equal $null (ConvertTo-ExcelTaxRate '') 'Missing rate must remain blank'
    Assert-Equal $null (ConvertTo-ExcelTaxRate '   ') 'Whitespace rate must remain blank'
    Assert-Equal 0 (ConvertTo-ExcelTaxRate '0%') 'Explicit zero rate stays zero'
}

Test-ReviewCase 'Detail invoice identity' {
    $rows = @()
    foreach ($seller in @('fake-seller-a', 'fake-seller-b')) {
        $summary = [pscustomobject]@{ Direction = 'purchase'; Source = 'api-detail'; XmlFile = ''; InvoiceId = ''; TemplateCode = '1'; InvoiceSeries = 'C26TEST'; InvoiceNumber = '1'; SellerTaxCode = $seller; TaxAmount = 10 }
        $rows += [pscustomobject]@{ InvoiceSummary = $summary; Source = 'api-detail'; AmountBeforeTax = 100; TaxRate = '10%'; TaxAmount = 10; Description = $seller }
    }
    $groups = @(Get-ExcelDetailGroups $rows)
    Assert-Equal 2 $groups.Count 'Different sellers with identical invoice numbers must not merge'
    $excelRows = @(New-ExcelDetailRows $rows)
    Assert-Equal 'fake-seller-b' $excelRows[1].Row.Cells[8].Value 'Second seller identity is preserved'
    Assert-Equal 10 $excelRows[1].Row.Cells[30].Value 'Each invoice retains its own tax check'
}

Test-ReviewCase 'Throttle never below configured delay' {
    $config = New-TestXmlConfig
    $config.RequestDelayMs = 20000
    $config.AdaptiveThrottle = $true
    Register-GdtRateLimit $config
    Assert-Equal $true ((Get-GdtRequestDelayMs $config) -ge 20000) '429 must not speed up list requests'
    $config.XmlRequestIntervalMs = 20000
    $null = New-HddtXmlSharedState $config
    try {
        Register-GdtXmlRateLimit -RetryAfterSeconds 1
        Assert-Equal $true ((Get-GdtXmlThrottleSnapshot).CurrentIntervalMs -ge 20000) '429 must not speed up XML requests'
    }
    finally { Set-HddtSharedState $null }
}

Test-ReviewCase 'HTTP retry metadata' {
    $transport = ${function:Invoke-WebRequest}
    $sleepFunction = ${function:Start-Sleep}
    $waitFunction = ${function:Wait-HddtRetryDelay}
    $script:ReviewHttpHits = 0
    try {
        Set-Item Function:\Invoke-WebRequest -Value {
            param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction)
            $script:ReviewHttpHits++
            if ($script:ReviewHttpHits -le 2) {
                $response = New-Object HddtTest.FakeWebResponse (429)
                throw (New-Object Net.WebException('Synthetic HTTP 429', $null, [Net.WebExceptionStatus]::ProtocolError, $response))
            }
            return [pscustomobject]@{ StatusCode = 200; Content = '{}' }
        }
        Set-Item Function:\Start-Sleep -Value { param($Seconds, $Milliseconds) }
        Set-Item Function:\Wait-HddtRetryDelay -Value { param($Seconds) }
        $config = New-TestXmlConfig
        $null = Invoke-GdtRequest $config ($config.BaseUrl + '/invoices/purchase')
        Assert-Equal 3 (Get-GdtLastRequestAttempts) 'Every HTTP attempt includes independent 429 retries'
        Assert-Equal 200 (Get-GdtLastStatusCode) 'Success replaces previous failure status'
    }
    finally {
        if ($null -ne $transport) { Set-Item Function:\Invoke-WebRequest -Value $transport }
        else { Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue }
        if ($null -ne $sleepFunction) { Set-Item Function:\Start-Sleep -Value $sleepFunction }
        else { Remove-Item Function:\Start-Sleep -ErrorAction SilentlyContinue }
        if ($null -ne $waitFunction) { Set-Item Function:\Wait-HddtRetryDelay -Value $waitFunction }
        else { Remove-Item Function:\Wait-HddtRetryDelay -ErrorAction SilentlyContinue }
    }
}

Test-ReviewCase 'Safe stop during retry' {
    $transport = ${function:Invoke-WebRequest}
    $sleepFunction = ${function:Start-Sleep}
    $script:ReviewStopHits = 0
    $script:ReviewSleepChunks = 0
    try {
        Set-Item Function:\Invoke-WebRequest -Value {
            param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction)
            $script:ReviewStopHits++
            $response = New-Object HddtTest.FakeWebResponse (429)
            throw (New-Object Net.WebException('Synthetic HTTP 429', $null, [Net.WebExceptionStatus]::ProtocolError, $response))
        }
        Set-Item Function:\Start-Sleep -Value {
            param($Seconds, $Milliseconds)
            $script:ReviewSleepChunks++
            Assert-Equal $true ($Milliseconds -gt 0 -and $Milliseconds -le 250) 'Retry checks stop within 250ms'
            Set-HddtStopRequest
        }
        Reset-HddtStopRequest
        $config = New-TestXmlConfig
        $stopped = $false
        try { $null = Invoke-GdtRequest $config ($config.BaseUrl + '/invoices/purchase') }
        catch { $stopped = $true }
        Assert-Equal $true $stopped 'Stop interrupts retry wait'
        Assert-Equal 1 $script:ReviewStopHits 'Stop prevents another HTTP attempt'
        Assert-Equal 1 $script:ReviewSleepChunks 'Stop does not wait the whole backoff'
    }
    finally {
        if ($null -ne $transport) { Set-Item Function:\Invoke-WebRequest -Value $transport }
        else { Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue }
        if ($null -ne $sleepFunction) { Set-Item Function:\Start-Sleep -Value $sleepFunction }
        else { Remove-Item Function:\Start-Sleep -ErrorAction SilentlyContinue }
        Reset-HddtStopRequest
    }
}

Test-ReviewCase 'Token refresh does not block result queue' {
    $config = New-TestXmlConfig
    $shared = New-HddtXmlSharedState $config
    $ready = New-Object Threading.ManualResetEvent($false)
    $release = New-Object Threading.ManualResetEvent($false)
    $shared['ReviewReady'] = $ready
    $shared['ReviewRelease'] = $release
    $worker = [powershell]::Create()
    try {
        $code = @'
param($Root, $Shared, $Config)
$ErrorActionPreference = 'Stop'
. (Join-Path $Root 'src/Logging.ps1')
. (Join-Path $Root 'src/Http.ps1')
Set-HddtSharedState $Shared
function Request-GdtSessionToken {
    param($Config)
    $null = $script:HddtSharedState.ReviewReady.Set()
    $null = $script:HddtSharedState.ReviewRelease.WaitOne(5000)
    return 'review-new-token'
}
Invoke-GdtTokenRefresh $Config -FailedToken $Config.Token
'@
        $null = $worker.AddScript($code).AddArgument($root).AddArgument($shared).AddArgument($config)
        $handle = $worker.BeginInvoke()
        Assert-Equal $true ($ready.WaitOne(5000)) 'Refresh worker reached login'
        $queueAvailable = [Threading.Monitor]::TryEnter($shared.SyncRoot, 0)
        if ($queueAvailable) {
            try {
                Add-HddtXmlResult ([pscustomobject]@{ Label = 'finished invoice' })
                Assert-Equal 1 @(Receive-HddtXmlResults).Count 'Main thread drains results while login is pending'
                Set-HddtSharedValue -Key 'StopRequested' -Value $false
            }
            finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
        }
        $null = $release.Set()
        $tokens = @($worker.EndInvoke($handle))
        Assert-Equal $true $queueAvailable 'Login never holds the shared result/stop lock'
        Assert-Equal 'review-new-token' ([string]$tokens[0]) 'Refresh returns the new token'
    }
    finally {
        $null = $release.Set()
        $worker.Dispose()
        $ready.Dispose()
        $release.Dispose()
        Set-HddtSharedState $null
        Reset-HddtStopRequest
    }
}

Test-ReviewCase 'Lookup with one shared string' {
    $path = Join-Path ([IO.Path]::GetTempPath()) ('hddt-review-lookup-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    $parts = @{
        'xl/workbook.xml' = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="LinkTraCuu" sheetId="1" r:id="rId1"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/worksheets/sheet1.xml' = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="D2" t="s"><v>0</v></c></row></sheetData></worksheet>'
        'xl/sharedStrings.xml' = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>https://lookup.example.invalid/</t></si></sst>'
    }
    $archive = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $parts.Keys) {
            $writer = New-Object IO.StreamWriter($archive.CreateEntry($name).Open(), [Text.Encoding]::UTF8)
            try { $writer.Write($parts[$name]) } finally { $writer.Dispose() }
        }
    }
    finally { $archive.Dispose() }
    try {
        $rows = @(Read-TaiHoaDonDienTuLookupSheet $path)
        Assert-Equal 'https://lookup.example.invalid/' $rows[0]['D'] 'Single shared string keeps the entire link'
    }
    finally { Remove-Item -LiteralPath $path -Force }
}

Test-ReviewCase 'Lookup import overrides the last duplicate' {
    $originalRows = $script:TaiHoaDonDienTuLookupSheetData
    $reader = ${function:Read-TaiHoaDonDienTuLookupSheet}
    try {
        $script:TaiHoaDonDienTuLookupSheetData = @(
            [ordered]@{ A = 'header' }
            [ordered]@{ B = 'fake-provider'; D = 'https://first.example.invalid/' }
            [ordered]@{ B = 'fake-provider'; D = 'https://last.example.invalid/' }
        )
        Set-Item Function:\Read-TaiHoaDonDienTuLookupSheet -Value {
            param($Path, $SheetName)
            return [ordered]@{ B = 'fake-provider'; D = 'https://custom.example.invalid/' }
        }
        $null = Import-TaiHoaDonDienTuLookupTable 'fake.xlsx'
        Assert-Equal 'https://custom.example.invalid/' (Get-TaiHoaDonDienTuLookupLink 'fake-provider') 'Imported row wins over every earlier duplicate'
    }
    finally {
        Set-Item Function:\Read-TaiHoaDonDienTuLookupSheet -Value $reader
        $script:TaiHoaDonDienTuLookupSheetData = $originalRows
        Sync-TaiHoaDonDienTuLookupIndex
    }
}

if ($reviewFailures.Count -gt 0) { throw ('FAILED: review regressions' + [Environment]::NewLine + ($reviewFailures -join [Environment]::NewLine)) }
