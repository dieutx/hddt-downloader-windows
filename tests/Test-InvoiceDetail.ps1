Set-StrictMode -Version 2.0

# Fixture gia; khong goi GDT hoac su dung thong tin dang nhap that.
$detailJson = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'fixtures/sample-invoice-detail.json'), [Text.Encoding]::UTF8)
$detailData = $detailJson | ConvertFrom-Json
$detailParsed = ConvertFrom-GdtInvoiceDetail -Detail $detailData -Direction purchase
Assert-Equal $detailData.nbten $detailParsed.Summary.SellerName 'Detail preserves Unicode seller'
Assert-Equal 120000 $detailParsed.Summary.AmountBeforeTax 'Detail summary amount'
Assert-Equal 132000 $detailParsed.Summary.TotalAmount 'Detail summary total'
Assert-Equal 1 @($detailParsed.Details).Count 'Detail item count'
Assert-Equal '921' $detailParsed.Details[0].LineNumber 'Detail keeps supplied line number'
Assert-Equal $detailData.hdhhdvu[0].ten $detailParsed.Details[0].Description 'Detail preserves Unicode item'
Assert-Equal $null $detailParsed.Details[0].Quantity 'Null quantity remains blank'
Assert-Equal $null $detailParsed.Details[0].UnitPrice 'Null unit price remains blank'
Assert-Equal 0.1 $detailParsed.Details[0].TaxRate 'Detail tax rate is a fraction'
Assert-Equal 132000 $detailParsed.Details[0].AmountWithTax 'Detail item total from before-tax amount and tax'
Assert-Equal 0 @(New-ExcelXmlRows $detailParsed.Details).Count 'API data never becomes an XML sheet row'
$detailMerged = Merge-GdtInvoiceData -Index ([pscustomobject]@{ nbdchi = 'Known address'; tgtcthue = 1; thttltsuat = @([pscustomobject]@{ tsuat = 0.1 }) }) -Detail $detailData
Assert-Equal 'Known address' $detailMerged.nbdchi 'Null detail does not erase a known index address'
Assert-Equal 120000 $detailMerged.tgtcthue 'Detail amount overrides stale index value'
Assert-Equal 1 @($detailMerged.thttltsuat).Count 'Empty detail array does not erase known tax buckets'
$emptyDetailData = $detailJson | ConvertFrom-Json
$emptyDetailData.hdhhdvu = $null
Assert-Equal 0 @((ConvertFrom-GdtInvoiceDetail $emptyDetailData purchase).Details).Count 'Null goods array keeps summary without invented items'

# PowerShell 5.1 co the doc Content sai charset; uu tien byte UTF-8 goc.
$detailBytes = [Text.Encoding]::UTF8.GetBytes($detailJson)
$detailStream = New-Object IO.MemoryStream (,$detailBytes)
try {
    $detailStream.Position = 3
    $decoded = ConvertFrom-GdtResponseText ([pscustomobject]@{ Content = [Text.Encoding]::GetEncoding(1252).GetString($detailBytes); RawContentStream = $detailStream })
    Assert-Equal $detailJson $decoded 'JSON transport decodes the original UTF-8 bytes'
    Assert-Equal 3 $detailStream.Position 'JSON decode preserves raw stream position'
    Assert-Equal $detailJson (ConvertFrom-GdtResponseText ([pscustomobject]@{ Content = $detailBytes })) 'Byte-only response decodes UTF-8'
    Assert-Equal $detailJson (ConvertFrom-GdtResponseText ([pscustomobject]@{ Content = $detailJson })) 'String-only response remains supported'
}
finally { $detailStream.Dispose() }

$detailInvoice = [pscustomobject]@{
    Direction = 'purchase'; Source = 'query'; SellerTaxCode = '0101111111'
    InvoiceSeries = 'C26TABC'; InvoiceNumber = '123'; InvoiceTemplate = '1'; InvoiceDate = [datetime]'2026-09-20'
}
$originalDetailRequest = ${function:Invoke-GdtRequest}
$script:DetailResponseJson = $detailJson
$script:DetailCapturedRequest = $null
try {
    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, $RequestProfile, $ExtraHeaders)
        $script:DetailCapturedRequest = @{ Uri = $Uri; Profile = $RequestProfile; Headers = $ExtraHeaders }
        return $script:DetailResponseJson
    }
    $detailConfig = [pscustomobject]@{ BaseUrl = 'https://example.invalid/api' }
    $receivedDetail = Get-GdtInvoiceDetail $detailConfig $detailInvoice
    Assert-Equal 'fake-detail-invoice' $receivedDetail.id 'Detail response parsed'
    Assert-Equal 'https://example.invalid/api/query/invoices/detail?nbmst=0101111111&khhdon=C26TABC&shdon=123&khmshdon=1' $script:DetailCapturedRequest.Uri 'Detail uses the verified endpoint and current invoice identity'
    Assert-Equal 'InvoiceDetail' $script:DetailCapturedRequest.Profile 'Detail uses JSON browser profile'
    Assert-Equal $true ($script:DetailCapturedRequest.Headers.Action.EndsWith('mua%20v%C3%A0o)')) 'Purchase detail action'
    $detailInvoice.Direction = 'sold'
    $null = Get-GdtInvoiceDetail $detailConfig $detailInvoice
    Assert-Equal $true ($script:DetailCapturedRequest.Headers.Action.EndsWith('b%C3%A1n%20ra)')) 'Sold detail action'
    $detailInvoice.Direction = 'purchase'
    foreach ($invalidDetail in @('{bad json', '{}', $detailJson.Replace('C26TABC', 'C26WRONG'))) {
        $script:DetailResponseJson = $invalidDetail
        $detailRejected = $false
        try { $null = Get-GdtInvoiceDetail $detailConfig $detailInvoice } catch { $detailRejected = $true }
        Assert-Equal $true $detailRejected 'Malformed or mismatched detail is rejected'
    }
}
finally { Set-Item Function:\Invoke-GdtRequest -Value $originalDetailRequest }

# Di qua HTTP that cua ung dung, chi thay transport bang mock offline.
$originalDetailWebRequest = ${function:Invoke-WebRequest}
$script:DetailTransportHits = 0
$script:DetailTransportHeaders = $null
$script:DetailTransportBytes = $detailBytes
$detailTransportDirectory = Join-Path ([IO.Path]::GetTempPath()) ('hddt-detail-transport-' + [guid]::NewGuid().ToString('N'))
try {
    $detailTransportConfig = New-TestXmlConfig
    $detailTransportConfig.XmlDirectory = $detailTransportDirectory
    $null = New-HddtXmlSharedState $detailTransportConfig
    Set-HddtSharedValue -Key 'Token' -Value 'current-synthetic-session'
    Set-HddtSharedValue -Key 'CooldownFallbackSeconds' -Value 0
    Reset-HddtStopRequest
    Set-Item Function:\Invoke-WebRequest -Value {
        param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction, $Body, $ContentType)
        $script:DetailTransportHits++
        $status = 200
        if ($Uri -like '*export-xml?*') { $status = 500 }
        elseif ($script:DetailTransportHits -eq 2) { $status = 429 }
        if ($status -ne 200) {
            $response = New-Object HddtTest.FakeWebResponse ($status)
            throw (New-Object Net.WebException('Synthetic HTTP error', $null, [Net.WebExceptionStatus]::ProtocolError, $response))
        }
        $script:DetailTransportHeaders = $Headers
        return [pscustomobject]@{
            StatusCode = 200
            Content = [Text.Encoding]::GetEncoding(1252).GetString($script:DetailTransportBytes)
            RawContentStream = (New-Object IO.MemoryStream (,$script:DetailTransportBytes))
        }
    }
    $transportResult = Get-GdtXmlDownloadResult $detailTransportConfig $detailInvoice 0
    Assert-Equal $true $transportResult.Success 'Real HTTP and scheduler recover XML 500 after detail 429'
    Assert-Equal 3 $script:DetailTransportHits 'Real transport requests XML then detail with its retry'
    Assert-Equal $detailData.nbten $transportResult.ApiDetail.nbten 'Real transport decodes Vietnamese JSON correctly'
    Assert-Equal 'Bearer current-synthetic-session' $script:DetailTransportHeaders.Authorization 'Fallback uses the current session token'
    Assert-Equal 1 (Get-GdtXmlThrottleSnapshot).RateLimitCount 'Detail 429 uses the shared scheduler'
    Assert-Equal 0 (Get-GdtXmlThrottleSnapshot).ActiveCount 'Detail fallback releases every scheduler slot'
}
finally {
    if ($null -ne $originalDetailWebRequest) { Set-Item Function:\Invoke-WebRequest -Value $originalDetailWebRequest }
    else { Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue }
    Set-HddtSharedState $null
    Reset-HddtStopRequest
    $resolvedDetailTransportDirectory = [IO.Path]::GetFullPath($detailTransportDirectory)
    $allowedDetailTransportPrefix = Join-Path ([IO.Path]::GetFullPath([IO.Path]::GetTempPath())) 'hddt-detail-transport-'
    if (-not $resolvedDetailTransportDirectory.StartsWith($allowedDetailTransportPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Thu muc kiem thu nam ngoai thu muc tam.' }
    if (Test-Path -LiteralPath $resolvedDetailTransportDirectory) { Remove-Item -LiteralPath $resolvedDetailTransportDirectory -Recurse -Force }
}

$originalDetailSaveXml = ${function:Save-GdtInvoiceXml}
$originalGetDetail = ${function:Get-GdtInvoiceDetail}
$script:DetailFallbackHits = 0
$script:DetailXmlStatus = 500
$script:DetailFallbackFails = $false
$script:DetailFixture = $detailData
try {
    Reset-HddtStopRequest
    Set-Item Function:\Save-GdtInvoiceXml -Value {
        param($Config, $Invoice)
        $script:LastGdtStatusCode = $script:DetailXmlStatus
        $script:LastGdtRequestUri = 'https://example.invalid/api/query/invoices/export-xml'
        if ($script:DetailXmlStatus -eq 200) { return 'fake.xml' }
        throw ('XML HTTP ' + $script:DetailXmlStatus)
    }
    Set-Item Function:\Get-GdtInvoiceDetail -Value {
        param($Config, $Invoice)
        $script:DetailFallbackHits++
        $script:LastGdtStatusCode = 200
        $script:LastGdtRequestUri = 'https://example.invalid/api/query/invoices/detail'
        if ($script:DetailFallbackFails) { throw 'Synthetic detail failure' }
        return $script:DetailFixture
    }
    $fallbackResult = Get-GdtXmlDownloadResult $detailConfig $detailInvoice 0
    Assert-Equal $true $fallbackResult.Success 'XML 500 is recovered with API detail'
    Assert-Equal 1 $script:DetailFallbackHits 'XML 500 requests detail once'
    Assert-Equal 500 $fallbackResult.XmlFailure.StatusCode 'Recovery keeps the original XML failure metadata'
    Assert-Equal 0 @($fallbackResult.XmlFiles).Count 'Recovery never invents an XML file'
    $script:DetailXmlStatus = 504
    Assert-Equal $false (Get-GdtXmlDownloadResult $detailConfig $detailInvoice 0).Success 'Other HTTP errors retain the existing failure behavior'
    $script:DetailXmlStatus = 200
    Assert-Equal $true (Get-GdtXmlDownloadResult $detailConfig $detailInvoice 0).Success 'Normal XML download remains successful'
    Assert-Equal 1 $script:DetailFallbackHits 'Other errors and normal XML never request detail'
    $script:DetailXmlStatus = 500
    $script:DetailFallbackFails = $true
    $failedDetailResult = Get-GdtXmlDownloadResult $detailConfig $detailInvoice 0
    Assert-Equal $false $failedDetailResult.Success 'Detail failure remains a reported failure'
    Assert-Equal 500 $failedDetailResult.XmlFailure.StatusCode 'Detail failure retains XML 500'
    Assert-Equal $true ($failedDetailResult.ErrorMessage -like '*Synthetic detail failure*') 'Detail failure includes the fallback reason'
    Set-HddtStopRequest
    $null = Get-GdtXmlDownloadResult $detailConfig $detailInvoice 0
    Assert-Equal 2 $script:DetailFallbackHits 'Stop request prevents a new fallback request'
}
finally {
    Set-Item Function:\Save-GdtInvoiceXml -Value $originalDetailSaveXml
    Set-Item Function:\Get-GdtInvoiceDetail -Value $originalGetDetail
    Reset-HddtStopRequest
}

# Nap dung cac ham entry point qua AST, khong chay dang nhap/tai du lieu.
$entryTokens = $null; $entryErrors = $null
$entryAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'Invoke-Hddt.ps1'), [ref]$entryTokens, [ref]$entryErrors)
foreach ($entryFunction in $entryAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    . ([scriptblock]::Create($entryFunction.Extent.Text))
}
$script:HddtXmlProcessedCount = 1; $script:HddtXmlInvoiceCount = 1
$detailBinding = [pscustomobject]@{
    Invoice = $detailInvoice; HasResult = $false; DownloadError = $null
    Details = (New-Object 'Collections.Generic.List[object]')
    Summary = [pscustomobject]@{
        Direction = 'purchase'; Source = 'query'; XmlFile = ''; InvoiceSeries = 'C26TABC'; InvoiceNumber = '123'
        GdtIndex = [pscustomobject]@{ nbdchi = 'Known address'; tgtcthue = 1 }
    }
}
Assert-Equal 1 (Add-HddtXmlResultRows $detailBinding $fallbackResult) 'Main entry point adds API goods to the workbook binding'
Assert-Equal 'api-detail' $detailBinding.Summary.DataOrigin 'Recovered summary records its data origin'
Assert-Equal 'Known address' $detailBinding.Summary.SellerAddress 'Main merge retains known data when API has null'
Assert-Equal 'Da lay du lieu tu API detail' $detailBinding.DownloadError.FinalResult 'Error sheet explains successful detail recovery'
$detailSummaryRows = @(New-ExcelSummaryRows @($detailBinding.Summary))
Assert-Equal $detailData.nbten $detailSummaryRows[0].Row.Cells[9].Value 'Summary sheet gets Vietnamese name from detail'
Assert-Equal 120000 $detailSummaryRows[0].Row.Cells[42].Value 'Summary sheet gets the recovered amount'
$detailExcelRows = @(New-ExcelDetailRows $detailBinding.Details.ToArray())
Assert-Equal 1 $detailExcelRows.Count 'Detail sheet gets recovered goods'

$detailWorkbookPath = Join-Path ([IO.Path]::GetTempPath()) ('hddt-detail-' + [guid]::NewGuid().ToString('N') + '.xlsx')
$detailArchive = $null
try {
    $soldDetailParsed = ConvertFrom-GdtInvoiceDetail $detailData sold
    $null = Export-InvoiceWorkbook -Path $detailWorkbookPath -SummaryRows @($detailBinding.Summary, $soldDetailParsed.Summary) -DetailRows @($detailBinding.Details.ToArray() + $soldDetailParsed.Details) -ErrorRows @($detailBinding.DownloadError)
    $detailArchive = [IO.Compression.ZipFile]::OpenRead($detailWorkbookPath)
    foreach ($sheetNumber in @(1, 2, 4, 5)) {
        $reader = New-Object IO.StreamReader ($detailArchive.GetEntry(('xl/worksheets/sheet{0}.xml' -f $sheetNumber)).Open(), [Text.Encoding]::UTF8)
        try { $detailSheetText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        Assert-Equal $true ($detailSheetText.Contains($detailData.nbten)) ('Unicode seller survives actual summary/detail package sheet ' + $sheetNumber)
    }
    foreach ($sheetNumber in @(3, 6)) {
        $reader = New-Object IO.StreamReader ($detailArchive.GetEntry(('xl/worksheets/sheet{0}.xml' -f $sheetNumber)).Open(), [Text.Encoding]::UTF8)
        try { $detailSheetText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        Assert-Equal $false ($detailSheetText.Contains($detailData.hdhhdvu[0].ten)) 'API goods are absent from actual XML sheets'
    }
}
finally {
    if ($null -ne $detailArchive) { $detailArchive.Dispose() }
    if (Test-Path -LiteralPath $detailWorkbookPath) { Remove-Item -LiteralPath $detailWorkbookPath -Force }
}
