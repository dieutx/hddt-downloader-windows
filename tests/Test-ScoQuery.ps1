Set-StrictMode -Version 2.0

# Du lieu gia: SCO gioi han 10 ngay va CAPTCHA rieng cho tung trang.
$scoRanges = @(Get-GdtInvoiceDateRanges -FromDate ([datetime]'2026-08-01') -ToDate ([datetime]'2026-08-31') -Source 'sco-query')
Assert-Equal 4 $scoRanges.Count 'SCO month is split into four ranges'
$nextDay = [datetime]'2026-08-01'
foreach ($range in $scoRanges) {
    Assert-Equal $nextDay $range.From 'SCO ranges have no gap or overlap'
    Assert-Equal $true (($range.To - $range.From).Days -le 9) 'SCO range has at most ten inclusive days'
    $nextDay = $range.To.AddDays(1)
}
Assert-Equal ([datetime]'2026-09-01') $nextDay 'Last day of month is included'
Assert-Equal 1 @(Get-GdtInvoiceDateRanges -FromDate ([datetime]'2026-08-01') -ToDate ([datetime]'2026-08-31') -Source 'query').Count 'Regular invoices retain monthly ranges'
Assert-Equal 2 @(Get-GdtInvoiceDateRanges -FromDate ([datetime]'2026-08-30') -ToDate ([datetime]'2026-09-02') -Source 'sco-query').Count 'SCO ranges preserve month boundaries'

$savedScoRequest = ${function:Invoke-GdtRequest}
$savedScoCaptcha = ${function:Get-GdtCaptcha}
$script:ScoCaptchaHits = 0
$script:ScoRequestUris = New-Object 'Collections.Generic.List[string]'
try {
    Set-Item Function:\Get-GdtCaptcha -Value {
        param($Config)
        $script:ScoCaptchaHits++
        [pscustomobject]@{ Key = ('key ' + $script:ScoCaptchaHits); Code = 'ABC+23' }
    }
    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, $RequestProfile)
        $script:ScoRequestUris.Add($Uri)
        if ($script:ScoRequestUris.Count -eq 1) { return '{"datas":[{"shdon":"1"}],"total":2,"state":"next+page"}' }
        return '{"datas":[{"shdon":"2"}],"total":2,"state":""}'
    }
    $scoConfig = [pscustomobject]@{
        IncludeRegular = $false; IncludeSco = $true
        FromDate = [datetime]'2026-08-01'; ToDate = [datetime]'2026-08-01'
        BaseUrl = 'https://example.invalid/api'; PageSize = 100
    }
    $scoRows = @(Get-GdtInvoiceIndex -Config $scoConfig -Direction 'sold')
    Assert-Equal 2 $scoRows.Count 'Both SCO pages are retained'
    Assert-Equal 2 $script:ScoCaptchaHits 'Each SCO page gets a fresh CAPTCHA'
    Assert-Equal $true ($script:ScoRequestUris[0] -match '[?&]size=50&') 'Page size is capped at fifty to avoid GDT HTTP 500'
    Assert-Equal $true ($script:ScoRequestUris[0] -match 'ckey=key%201&cvalue=ABC%2B23') 'CAPTCHA parameters are encoded'
    Assert-Equal $true ($script:ScoRequestUris[1] -match 'state=next%2Bpage.*ckey=key%202') 'Cursor and fresh CAPTCHA are sent together'
    Assert-Equal 0 @(Get-GdtIndexErrors).Count 'Complete SCO query has no period error'

    $script:ScoCaptchaHits = 0
    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, $RequestProfile)
        if ($script:ScoCaptchaHits -eq 1) { throw 'Synthetic HTTP 400' }
        return '{"datas":[],"total":0,"state":""}'
    }
    $null = Invoke-GdtInvoicePage $scoConfig 'https://example.invalid/api/sco-query/invoices/sold?size=50' 'sco-query'
    Assert-Equal 2 $script:ScoCaptchaHits 'HTTP 400 retries with a new CAPTCHA'

    $script:ScoCaptchaHits = 0
    Set-Item Function:\Invoke-GdtRequest -Value { param($Config, $Uri, $RequestProfile) throw 'Synthetic HTTP 400' }
    $failed = $false
    try { $null = Invoke-GdtInvoicePage $scoConfig 'https://example.invalid/api/sco-query/invoices/sold?size=50' 'sco-query' }
    catch { $failed = $true }
    Assert-Equal $true $failed 'Persistent CAPTCHA failure is reported'
    Assert-Equal 3 $script:ScoCaptchaHits 'CAPTCHA retries are bounded'

    Set-Item Function:\Invoke-GdtRequest -Value { param($Config, $Uri, $RequestProfile) return '{"datas":[{"shdon":"1"}],"total":2,"state":""}' }
    $scoRows = @(Get-GdtInvoiceIndex -Config $scoConfig -Direction 'purchase')
    Assert-Equal 1 $scoRows.Count 'Partial list data is retained'
    Assert-Equal 1 @(Get-GdtIndexErrors).Count 'Missing cursor before total is reached is reported'

    Set-Item Function:\Invoke-GdtRequest -Value { param($Config, $Uri, $RequestProfile) return '{"datas":[{"shdon":"1"}],"total":5,"state":"repeated"}' }
    $null = @(Get-GdtInvoiceIndex -Config $scoConfig -Direction 'sold')
    Assert-Equal 1 @(Get-GdtIndexErrors).Count 'Repeated cursor reports incomplete query'
}
finally {
    Set-Item Function:\Invoke-GdtRequest -Value $savedScoRequest
    Set-Item Function:\Get-GdtCaptcha -Value $savedScoCaptcha
}
