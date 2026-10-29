Set-StrictMode -Version 2.0

# Body loi giong GDT, chi dung du lieu gia.
$rangeMessage = '{"message":"Kho\u1ea3ng th\u1eddi gian t\u00ecm ki\u1ebfm kh\u00f4ng \u0111\u01b0\u1ee3c l\u1edbn h\u01a1n 5 ng\u00e0y."}'
Assert-Equal 5 (Get-GdtRangeLimitDays $rangeMessage) 'Decode Vietnamese JSON range limit'
Assert-Equal 3 (Get-GdtRangeLimitDays 'Khoang thoi gian tim kiem toi da la 3 ngay.') 'Accept a future range limit'
Assert-Equal 1 (Get-GdtRangeLimitDays 'Date range cannot exceed 1 day.') 'Accept a one-day limit'
foreach ($message in @('CAPTCHA khong dung 5 ngay', 'Khoang thoi gian tim kiem khong duoc lon hon 0 ngay.', 'Khoang thoi gian tim kiem khong duoc lon hon 999999999999 ngay.', 'HTTP 400', 'Khoang thoi gian tim kiem 5 ngay')) {
    Assert-Equal 0 (Get-GdtRangeLimitDays $message) 'Do not infer limits from unrelated or invalid errors'
}

# Cac duong doc body cua PS 5.1 va PS 7 deu duoc kiem tra tren ca hai runtime.
Add-Type -AssemblyName System.Net.Http
$response = [pscustomobject]@{ Content = (New-Object Net.Http.StringContent($rangeMessage)) }
$exception = New-Object InvalidOperationException('Synthetic response')
Add-Member -InputObject $exception -NotePropertyName Response -NotePropertyValue $response
$record = New-Object Management.Automation.ErrorRecord($exception, 'BodyTest', ([Management.Automation.ErrorCategory]::InvalidOperation), $null)
try { Assert-Equal 5 (Get-GdtErrorRangeLimitDays $record) 'Read HttpResponseMessage content on PowerShell 7' }
finally { $response.Content.Dispose() }
$response = [pscustomobject]@{ Body = $rangeMessage }
Add-Member -InputObject $response -MemberType ScriptMethod -Name GetResponseStream -Value { return New-Object IO.MemoryStream(,[Text.Encoding]::UTF8.GetBytes($this.Body)) }
$exception = New-Object InvalidOperationException('Synthetic stream response')
Add-Member -InputObject $exception -NotePropertyName Response -NotePropertyValue $response
$record = New-Object Management.Automation.ErrorRecord($exception, 'StreamTest', ([Management.Automation.ErrorCategory]::InvalidOperation), $null)
Assert-Equal 5 (Get-GdtErrorRangeLimitDays $record) 'Read UTF-8 WebResponse stream on PowerShell 5.1'

$savedRangeRequest = ${function:Invoke-GdtRequest}
$savedRangeCaptcha = ${function:Get-GdtCaptcha}
$savedRangeWebRequest = ${function:Invoke-WebRequest}
$savedRangeStatus = ${function:Get-HttpStatusCode}
try {
    # Loi HTTP that phai truyen metadata so ngay, khong truyen body nhay cam.
    Set-Item Function:\Get-HttpStatusCode -Value { param($ErrorRecord) return 400 }
    Set-Item Function:\Invoke-WebRequest -Value {
        $record = New-Object Management.Automation.ErrorRecord((New-Object InvalidOperationException('Synthetic HTTP 400')), 'Range400', ([Management.Automation.ErrorCategory]::InvalidOperation), $null)
        $record.ErrorDetails = New-Object Management.Automation.ErrorDetails('{"message":"Khoang thoi gian tim kiem khong duoc lon hon 5 ngay.","details":"private-test-value"}')
        throw $record
    }
    $httpConfig = [pscustomobject]@{ BaseUrl = 'https://example.invalid/api'; Token = 'fake'; RequestDelayMs = 0; AdaptiveThrottle = $false; MaxRetries = 0; HttpTimeoutSeconds = 30 }
    $caughtRange = $false
    try { $null = Invoke-GdtRequest -Config $httpConfig -Uri ($httpConfig.BaseUrl + '/sco-query/invoices/sold') }
    catch {
        $caughtRange = $true
        Assert-Equal 5 $_.Exception.Data['GdtRangeLimitDays'] 'HTTP layer exposes only the numeric limit'
        Assert-Equal $false ($_.Exception.Message -match 'private-test-value') 'Raw response body stays private'
    }
    Assert-Equal $true $caughtRange 'HTTP range rejection is thrown'
    Set-Item Function:\Get-HttpStatusCode -Value $savedRangeStatus
    if ($null -ne $savedRangeWebRequest) { Set-Item Function:\Invoke-WebRequest -Value $savedRangeWebRequest }
    else { Remove-Item Function:\Invoke-WebRequest }

    $script:RangeRequests = New-Object 'Collections.Generic.List[object]'
    $script:RangeCaptchas = 0
    $script:RangeLimit = 5
    $script:RangeRejected = 0
    Set-Item Function:\Get-GdtCaptcha -Value {
        param($Config)
        $script:RangeCaptchas++
        [pscustomobject]@{ Key = ('key' + $script:RangeCaptchas); Code = 'ABC23' }
    }
    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, $RequestProfile)
        $decoded = [uri]::UnescapeDataString($Uri)
        if ($decoded -notmatch 'tdlap=ge=(\d{2}/\d{2}/\d{4})T.*tdlap=le=(\d{2}/\d{2}/\d{4})T') { throw 'Missing test dates' }
        $from = [datetime]::ParseExact($Matches[1], 'dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
        $to = [datetime]::ParseExact($Matches[2], 'dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
        $script:RangeRequests.Add([pscustomobject]@{ From = $from; To = $to; Uri = $Uri })
        if (($to - $from).Days -ge $script:RangeLimit) {
            $script:RangeRejected++
            $error = New-Object InvalidOperationException('Synthetic range limit HTTP 400')
            $error.Data['GdtRangeLimitDays'] = $script:RangeLimit
            throw $error
        }
        $datas = @(for ($day = $from; $day -le $to; $day = $day.AddDays(1)) { @{ shdon = $day.ToString('yyyyMMdd'); tdlap = $day.ToString('yyyy-MM-dd') } })
        return (@{ datas = $datas; total = $datas.Count; state = '' } | ConvertTo-Json -Depth 5 -Compress)
    }
    foreach ($source in @('query', 'sco-query')) {
        foreach ($limit in @(5, 3, 1)) {
            $script:RangeLimit = $limit
            $script:RangeRequests.Clear(); $script:RangeCaptchas = 0; $script:RangeRejected = 0
            $config = [pscustomobject]@{ IncludeRegular = ($source -eq 'query'); IncludeSco = ($source -eq 'sco-query'); FromDate = [datetime]'2026-08-28'; ToDate = [datetime]'2026-10-02'; BaseUrl = 'https://example.invalid/api'; PageSize = 50 }
            $rows = @(Get-GdtInvoiceIndex $config 'purchase')
            Assert-Equal 36 $rows.Count ('All days retained for ' + $source + ' limit ' + $limit)
            Assert-Equal 36 @($rows | Select-Object -ExpandProperty InvoiceNumber -Unique).Count 'No missing or duplicate dates after splitting'
            Assert-Equal 1 $script:RangeRejected 'Learned limit prevents repeated oversized requests'
            Assert-Equal 0 @(Get-GdtIndexErrors).Count 'Recovered range rejection is not an invoice error'
            if ($source -eq 'sco-query') { Assert-Equal $script:RangeRequests.Count $script:RangeCaptchas 'No CAPTCHA retry for a known date error'
            } else { Assert-Equal 0 $script:RangeCaptchas 'Regular query never adds CAPTCHA' }
        }
    }

    $savedAdaptivePage = ${function:Invoke-GdtRequest}
    $script:RangeRequests.Clear(); $script:RangeRejected = 0; $script:RangeLimit = 5
    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, $RequestProfile)
        if ($script:RangeRequests.Count -ge 2) { $script:RangeLimit = 2 }
        & $savedAdaptivePage -Config $Config -Uri $Uri -RequestProfile $RequestProfile
    }
    $config.FromDate = [datetime]'2026-08-01'; $config.ToDate = [datetime]'2026-08-21'
    $rows = @(Get-GdtInvoiceIndex $config 'sold')
    Assert-Equal 21 $rows.Count 'A second reduction in the same run still covers every day'
    Assert-Equal 2 $script:RangeRejected 'Range limit can decrease again from five to two days'
    Assert-Equal 0 @(Get-GdtIndexErrors).Count 'Both reductions recover without period errors'

    # Gioi han doi tren trang thu hai: giu dong cu va quet lai khong nhan doi.
    $script:PartialRangeCalls = 0
    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, $RequestProfile)
        $script:PartialRangeCalls++
        switch ($script:PartialRangeCalls) {
            1 { return '{"datas":[{"shdon":"1"}],"total":2,"state":"next"}' }
            2 {
                $error = New-Object InvalidOperationException('Synthetic HTTP 400')
                $error.Data['GdtRangeLimitDays'] = 1
                throw $error
            }
            3 { return '{"datas":[{"shdon":"1"}],"total":1,"state":""}' }
            default { return '{"datas":[{"shdon":"2"}],"total":1,"state":""}' }
        }
    }
    $config.FromDate = [datetime]'2026-08-01'; $config.ToDate = [datetime]'2026-08-02'
    $rows = @(Get-GdtInvoiceIndex $config 'sold')
    Assert-Equal 2 $rows.Count 'Partial pages retained without replay duplicates'
    Assert-Equal 0 @(Get-GdtIndexErrors).Count 'Replayed rows still count toward page total'

    $script:PartialRangeCalls = 0
    $savedPartialPage = ${function:Invoke-GdtRequest}
    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, $RequestProfile)
        if ($script:PartialRangeCalls -ge 2) { throw 'Synthetic HTTP 500 after range recovery' }
        & $savedPartialPage -Config $Config -Uri $Uri -RequestProfile $RequestProfile
    }
    $rows = @(Get-GdtInvoiceIndex $config 'sold')
    Assert-Equal 1 $rows.Count 'Previously received rows survive even if all smaller periods fail'
    Assert-Equal 2 @(Get-GdtIndexErrors).Count 'Each unrecovered subperiod reports its own failure'

    # Khong chia lai vo han neu GDT tu choi ca ky mot ngay.
    $script:RangeCaptchas = 0
    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, $RequestProfile)
        $error = New-Object InvalidOperationException('Synthetic HTTP 400')
        $error.Data['GdtRangeLimitDays'] = 1
        throw $error
    }
    $config.ToDate = $config.FromDate
    $null = @(Get-GdtInvoiceIndex $config 'sold')
    Assert-Equal 1 $script:RangeCaptchas 'One-day rejection terminates without CAPTCHA retry'
    Assert-Equal 1 @(Get-GdtIndexErrors).Count 'Unrecoverable one-day rejection remains visible'
}
finally {
    Set-Item Function:\Invoke-GdtRequest -Value $savedRangeRequest
    Set-Item Function:\Get-GdtCaptcha -Value $savedRangeCaptcha
    Set-Item Function:\Get-HttpStatusCode -Value $savedRangeStatus
    if ($null -ne $savedRangeWebRequest) { Set-Item Function:\Invoke-WebRequest -Value $savedRangeWebRequest }
    else { Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue }
}
