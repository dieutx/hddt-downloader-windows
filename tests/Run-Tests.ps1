Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'src\Logging.ps1')
. (Join-Path $root 'src\Config.ps1')
. (Join-Path $root 'src\Http.ps1')
. (Join-Path $root 'src\InvoiceApi.ps1')
. (Join-Path $root 'src\XmlParser.ps1')
. (Join-Path $root 'src\ExcelExporter.ps1')
. (Join-Path $root 'src\Login.ps1')

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    if ($Expected -ne $Actual) {
        throw "FAILED: $Message. Expected=[$Expected], Actual=[$Actual]"
    }
}

$parsed = ConvertFrom-InvoiceXml -Path (Join-Path $PSScriptRoot 'fixtures\sample-invoice.xml') -Direction 'purchase' -Source 'query'
Assert-Equal 'C26TABC' $parsed.Summary.InvoiceSeries 'Invoice series'
Assert-Equal '123' $parsed.Summary.InvoiceNumber 'Invoice number'
Assert-Equal 250000 $parsed.Summary.AmountBeforeTax 'Summary before tax'
Assert-Equal 24000 $parsed.Summary.TaxAmount 'Summary tax'
Assert-Equal 'ABC123' $parsed.Summary.TaxAuthorityCode 'Tax authority code'
Assert-Equal 2 @($parsed.Details).Count 'Detail row count'
Assert-Equal 20000 $parsed.Details[0].TaxAmount 'First item tax'
Assert-Equal 220000 $parsed.Details[0].AmountWithTax 'First item amount with tax fallback'
Assert-Equal 'Sản phẩm B' $parsed.Details[1].Description 'Second item description'

$directionRoot = Join-Path ([IO.Path]::GetTempPath()) ('hddt-direction-' + [guid]::NewGuid().ToString('N'))
$purchaseFolder = Join-Path $directionRoot 'purchase'
$soldFolder = Join-Path $directionRoot 'sold'
New-Item -ItemType Directory -Path $purchaseFolder, $soldFolder -Force | Out-Null
try {
    Assert-Equal 'purchase' (Resolve-LocalInvoiceDirection -Path (Join-Path $purchaseFolder 'invoice.xml') -RootDirectory $directionRoot -ConfiguredDirection 'auto') 'Auto direction from purchase folder'
    Assert-Equal 'sold' (Resolve-LocalInvoiceDirection -Path (Join-Path $soldFolder 'invoice.xml') -RootDirectory $directionRoot -ConfiguredDirection 'auto') 'Auto direction from sold folder'
    Assert-Equal 'purchase' (Resolve-LocalInvoiceDirection -Path (Join-Path $directionRoot 'purchase_query_test.xml') -RootDirectory $directionRoot -ConfiguredDirection 'auto') 'Auto direction from legacy file name'
    Assert-Equal 'sold' (Resolve-LocalInvoiceDirection -Path (Join-Path $purchaseFolder 'invoice.xml') -RootDirectory $directionRoot -ConfiguredDirection 'sold') 'Explicit direction overrides folder'
}
finally {
    Remove-Item -LiteralPath $directionRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$ranges = @(Get-MonthDateRanges -FromDate ([datetime]'2026-08-20') -ToDate ([datetime]'2026-10-02'))
Assert-Equal 3 $ranges.Count 'Monthly date range count'
Assert-Equal ([datetime]'2026-08-31') $ranges[0].To 'First monthly range end'
Assert-Equal ([datetime]'2026-10-01') $ranges[2].From 'Last monthly range start'
Assert-Equal 5 (Get-RetryDelaySeconds -StatusCode 429 -Attempt 1) 'First 429 backoff'
Assert-Equal 20 (Get-RetryDelaySeconds -StatusCode 429 -Attempt 3) 'Third 429 backoff'
Assert-Equal 45 (Get-RetryDelaySeconds -StatusCode 429 -Attempt 1 -RetryAfterSeconds 45) 'Retry-After precedence'

$originalGdtRequest = ${function:Invoke-GdtRequest}
$script:IndexRequestCount = 0
Set-Item Function:\Invoke-GdtRequest -Value {
    param($Config, $Uri, $AsBytes)
    $script:IndexRequestCount++
    if ($script:IndexRequestCount -gt 2) { throw 'Pagination did not stop at repeated state.' }
    return '{"datas":[{"nbmst":"test","khhdon":"C26TABC","shdon":"123","khmshdon":"1","tdlap":"2026-09-01"}],"state":"same-state"}'
}
try {
    $indexConfig = [pscustomobject]@{
        IncludeRegular = $true; IncludeSco = $false
        FromDate = [datetime]'2026-09-01'; ToDate = [datetime]'2026-09-30'
        BaseUrl = 'https://hoadondientu.gdt.gov.vn:30000'; PageSize = 50
    }
    $indexRows = @(Get-GdtInvoiceIndex -Config $indexConfig -Direction 'purchase')
    Assert-Equal 2 $script:IndexRequestCount 'Repeated pagination state stops requests'
    Assert-Equal 2 $indexRows.Count 'Rows from completed pages remain available'
}
finally { Set-Item Function:\Invoke-GdtRequest -Value $originalGdtRequest }

$originalStatusResolver = $null
function Invoke-WebRequest {
    param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction)
    return [pscustomobject]@{
        StatusCode = 200
        RawContentStream = New-Object IO.MemoryStream (,[byte[]](80, 75, 3, 4))
        Content = $null
    }
}
try {
    $httpConfig = [pscustomobject]@{
        BaseUrl = 'https://hoadondientu.gdt.gov.vn:30000'
        Token = 'test-token'
        HttpTimeoutSeconds = 30
        RequestDelayMs = 0
        MaxRetries = 0
    }
    $memoryResult = [byte[]](Invoke-GdtRequest -Config $httpConfig -Uri ($httpConfig.BaseUrl + '/test-memory.zip') -AsBytes)
    Assert-Equal 4 $memoryResult.Length 'In-memory binary request length'

    $originalStatusResolver = ${function:Get-HttpStatusCode}
    Set-Item Function:\Get-HttpStatusCode -Value { param($ErrorRecord) return 500 }
    Set-Item Function:\Invoke-WebRequest -Value {
        param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction)
        throw 'Simulated server error'
    }
    $missingXmlDetected = $false
    try {
        Invoke-GdtRequest -Config $httpConfig -Uri ($httpConfig.BaseUrl + '/missing-xml.zip') -AsBytes | Out-Null
    }
    catch {
        $missingXmlDetected = ($_.Exception.Message -eq 'GDT không có hồ sơ XML gốc cho hóa đơn này (HTTP 500).')
    }
    Assert-Equal $true $missingXmlDetected 'Export XML HTTP 500 is not retried'
    Set-Item Function:\Get-HttpStatusCode -Value $originalStatusResolver
}
finally {
    if ($null -ne $originalStatusResolver) { Set-Item Function:\Get-HttpStatusCode -Value $originalStatusResolver }
    Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue
}

$tempXmlDirectory = Join-Path ([IO.Path]::GetTempPath()) ('hddt-xml-' + [guid]::NewGuid().ToString('N'))
Add-Type -AssemblyName System.IO.Compression
$zipMemory = New-Object IO.MemoryStream
$testArchive = New-Object IO.Compression.ZipArchive($zipMemory, [IO.Compression.ZipArchiveMode]::Create, $true)
try {
    $xmlEntry = $testArchive.CreateEntry('nested/invoice.xml')
    $xmlWriter = New-Object IO.StreamWriter($xmlEntry.Open())
    try { $xmlWriter.Write('<HDon />') } finally { $xmlWriter.Dispose() }
    $ignoredEntry = $testArchive.CreateEntry('readme.txt')
    $ignoredWriter = New-Object IO.StreamWriter($ignoredEntry.Open())
    try { $ignoredWriter.Write('ignored') } finally { $ignoredWriter.Dispose() }
}
finally { $testArchive.Dispose() }
try {
    $extractedXml = @(Expand-InvoiceXmlBytes -Bytes $zipMemory.ToArray() -DestinationDirectory $tempXmlDirectory -FileNamePrefix 'purchase_query_test_123')
    Assert-Equal 1 $extractedXml.Count 'Only XML is written from in-memory archive'
    Assert-Equal $true (Test-Path -LiteralPath (Join-Path $tempXmlDirectory 'purchase_query_test_123.xml')) 'Flat extracted XML exists'
    Assert-Equal $false (Test-Path -LiteralPath (Join-Path $tempXmlDirectory 'readme.txt')) 'Non-XML entry is not written'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $tempXmlDirectory -Directory).Count 'No per-invoice directory is created'
}
finally {
    $zipMemory.Dispose()
    Remove-Item -LiteralPath $tempXmlDirectory -Recurse -Force -ErrorAction SilentlyContinue
}

$tempXlsx = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.xlsx')
try {
    Export-InvoiceWorkbook -Path $tempXlsx -SummaryRows @($parsed.Summary) -DetailRows @($parsed.Details) -ErrorRows @() -Overwrite
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($tempXlsx)
    try {
        Assert-Equal 1 @($zip.Entries | Where-Object FullName -eq 'xl/workbook.xml').Count 'Workbook package entry'
        Assert-Equal 3 @($zip.Entries | Where-Object FullName -like 'xl/worksheets/sheet*.xml').Count 'Worksheet package count'
        Assert-Equal 2 @($zip.Entries | Where-Object FullName -like 'xl/tables/table*.xml').Count 'Table package count'

        for ($sheetNumber = 1; $sheetNumber -le 3; $sheetNumber++) {
            $sheetEntry = $zip.GetEntry("xl/worksheets/sheet$sheetNumber.xml")
            $sheetReader = New-Object IO.StreamReader($sheetEntry.Open())
            try { [xml]$sheetDocument = $sheetReader.ReadToEnd() }
            finally { $sheetReader.Dispose() }
            $sheetNamespaceManager = New-Object Xml.XmlNamespaceManager($sheetDocument.NameTable)
            $sheetNamespaceManager.AddNamespace('x', $script:SpreadsheetNamespace)
            $sheetTableCount = @($sheetDocument.SelectNodes('/x:worksheet/x:tableParts/x:tablePart', $sheetNamespaceManager)).Count
            $sheetFilterCount = @($sheetDocument.SelectNodes('/x:worksheet/x:autoFilter', $sheetNamespaceManager)).Count
            if ($sheetTableCount -gt 0) {
                Assert-Equal 0 $sheetFilterCount "Sheet $sheetNumber does not duplicate table AutoFilter at worksheet level"
            }
            else {
                Assert-Equal 1 $sheetFilterCount "Sheet $sheetNumber keeps worksheet AutoFilter when there is no table"
            }
        }

        $stylesEntry = $zip.GetEntry('xl/styles.xml')
        $stylesReader = New-Object IO.StreamReader($stylesEntry.Open())
        try { [xml]$stylesDocument = $stylesReader.ReadToEnd() }
        finally { $stylesReader.Dispose() }
        $namespaceManager = New-Object Xml.XmlNamespaceManager($stylesDocument.NameTable)
        $namespaceManager.AddNamespace('x', $script:SpreadsheetNamespace)
        $headerFont = $stylesDocument.SelectSingleNode('/x:styleSheet/x:fonts/x:font[2]', $namespaceManager)
        $headerFontOrder = (($headerFont.ChildNodes | ForEach-Object { $_.LocalName }) -join ',')
        Assert-Equal 'b,sz,color,name' $headerFontOrder 'Header font element order follows Open XML schema'
    }
    finally { $zip.Dispose() }
}
finally {
    Remove-Item -LiteralPath $tempXlsx -Force -ErrorAction SilentlyContinue
}

$tempEnv = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.env')
try {
    @'
GDT_TOKEN=Bearer abcdefghijklmnopqrstuvwxyz123456
INVOICE_DIRECTION=purchase
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
OUTPUT_DIR=output
OUTPUT_XLSX=test.xlsx
'@ | Set-Content -LiteralPath $tempEnv -Encoding UTF8
    $config = Get-HddtConfig -EnvFile $tempEnv -RepositoryRoot $root
    Assert-Equal 'abcdefghijklmnopqrstuvwxyz123456' $config.Token 'Bearer normalization'
    Assert-Equal 1 $config.Directions.Count 'Direction count'
    Assert-Equal 'purchase' $config.Directions[0] 'Direction value'
    Assert-Equal 600 $config.RequestDelayMs 'Default global request delay'
    Assert-Equal 'info' $config.LogLevel 'Default log level'
    Assert-Equal 1 $config.ProgressEvery 'Default progress interval'
    Assert-Equal $true $config.LogToFile 'Default file logging'
    Assert-Equal $true $config.AdaptiveThrottle 'Default adaptive throttling'
    Assert-Equal $false $config.RedownloadXml 'Default XML resume mode'
}
finally {
    Remove-Item -LiteralPath $tempEnv -Force -ErrorAction SilentlyContinue
}

# --- Dieu khien dung an toan (Ctrl+C) ---
Reset-HddtStopRequest
Assert-Equal $false (Test-HddtStopRequested) 'Stop flag initially false'
Set-HddtStopRequest
Assert-Equal $true (Test-HddtStopRequested) 'Stop flag set'
Reset-HddtStopRequest
Assert-Equal $false (Test-HddtStopRequested) 'Stop flag reset'

# --- Nhan dien CAPTCHA SVG (port tu modDetectCaptcha.bas) ---
function New-TestCaptchaPath {
    param([int]$KeywordIndex, [double]$Position)
    $keyword = $script:CaptchaKeywords[$KeywordIndex]
    $commands = -join $keyword.ToCharArray()
    $builder = New-Object Text.StringBuilder
    foreach ($command in $commands.ToCharArray()) {
        switch ($command) {
            'M' { [void]$builder.Append(('M{0} 0 ' -f $Position.ToString([Globalization.CultureInfo]::InvariantCulture))) }
            'Q' { [void]$builder.Append('Q1 1 2 2 ') }
            'Z' { [void]$builder.Append('Z ') }
        }
    }
    return $builder.ToString().TrimEnd()
}

function New-TestCaptchaResponse {
    param([int[]]$KeywordIndexes, [double[]]$Positions)
    $paths = @()
    for ($index = 0; $index -lt $KeywordIndexes.Count; $index++) {
        $paths += '<path d="{0}" />' -f (New-TestCaptchaPath -KeywordIndex $KeywordIndexes[$index] -Position $Positions[$index])
    }
    $svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 40">{0}</svg>' -f ($paths -join '')
    return (@{ key = 'captcha-key-1'; content = $svg } | ConvertTo-Json -Compress)
}

$orderedResponse = New-TestCaptchaResponse -KeywordIndexes @(19, 4, 2) -Positions @(30, 10, 20)
$orderedPayload = $orderedResponse | ConvertFrom-Json
Assert-Equal 'ECT' (ConvertFrom-SvgCaptcha -Text (Get-JsonTextValue $orderedPayload 'content')) 'CAPTCHA decodes and sorts by coordinate'

for ($keywordIndex = 0; $keywordIndex -lt $script:CaptchaKeywords.Count; $keywordIndex++) {
    if ([string]::IsNullOrEmpty($script:CaptchaKeywords[$keywordIndex])) { continue }
    $expectedCharacter = if ($keywordIndex -le 26) { [string][char]($keywordIndex + 65) } else { [string]($keywordIndex - 26) }
    $singlePayload = (New-TestCaptchaResponse -KeywordIndexes @($keywordIndex) -Positions @(7)) | ConvertFrom-Json
    $decoded = ConvertFrom-SvgCaptcha -Text (Get-JsonTextValue $singlePayload 'content')
    Assert-Equal $expectedCharacter $decoded ('CAPTCHA keyword round-trip at index ' + $keywordIndex)
}
Assert-Equal '' (ConvertFrom-SvgCaptcha -Text '<svg><path d="M0 0 Z" /></svg>') 'Unknown path keyword is ignored'

# --- Dang nhap voi CAPTCHA tu dong ---
$originalLoginRequest = ${function:Invoke-GdtRequest}
$script:LoginCalls = @()
$script:LoginAuthBodies = @()
Set-Item Function:\Invoke-GdtRequest -Value {
    param($Config, $Uri, [string]$Method, [string]$Body, [switch]$SkipAuthorization, [hashtable]$ExtraHeaders, [switch]$AsBytes, [string]$ContentType)
    $script:LoginCalls += [pscustomobject]@{ Uri = $Uri; Method = $Method; Body = $Body }
    if ($Uri -like '*/captcha') {
        return (New-TestCaptchaResponse -KeywordIndexes @(19, 4, 2) -Positions @(30, 10, 20))
    }
    if ($Uri -like '*/authenticate') {
        $script:LoginAuthBodies += $Body
        if ($script:LoginAuthBodies.Count -eq 1) { return '{"message":"Mã CAPTCHA không đúng."}' }
        return '{"token":"jwt-token-123"}'
    }
    throw ('Unexpected URI in login test: ' + $Uri)
}
try {
    $loginConfig = [pscustomobject]@{
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'
        Username = 'tester'
        Password = 'secret-password'
        Token = ''
        RequestDelayMs = 0
        AdaptiveThrottle = $false
        MaxRetries = 0
        HttpTimeoutSeconds = 30
    }
    Reset-HddtStopRequest
    $loginToken = Invoke-GdtLogin -Config $loginConfig
    Assert-Equal 'jwt-token-123' $loginToken 'Login returns token'
    Assert-Equal 4 $script:LoginCalls.Count 'Login retries once after captcha error'
    $authPayload = $script:LoginAuthBodies[-1] | ConvertFrom-Json
    Assert-Equal 'tester' $authPayload.username 'Login sends username'
    Assert-Equal 'secret-password' $authPayload.password 'Login sends password'
    Assert-Equal 'captcha-key-1' $authPayload.ckey 'Login sends captcha key'
    Assert-Equal 'ECT' $authPayload.cvalue 'Login sends decoded captcha'
}
finally { Set-Item Function:\Invoke-GdtRequest -Value $originalLoginRequest }

# --- Chuoi hoa don lien quan (relative/related) ---
$originalRelationRequest = ${function:Invoke-GdtRequest}
$script:RelationUris = @()
$script:RelationFailRelative = $false
$script:RelativeResponse = '[{"khmshdon":"1","khhdon":"C26TABC","shdon":"100","khmshdgoc":"1","khhdgoc":"C26TOLD","shdgoc":"50","tthai":2}]'
$script:RelatedResponse = '{"mtthdtbssrs":[{"ten":"Thông báo hóa đơn","ngay":"2026-09-02T00:00:00","ldo":"sai sót kinh doanh","loai":"3","kqtnhan":"1"}]}'
Set-Item Function:\Invoke-GdtRequest -Value {
    param($Config, $Uri, [string]$Method, [string]$Body, [switch]$SkipAuthorization, [hashtable]$ExtraHeaders, [switch]$AsBytes, [string]$ContentType)
    $script:RelationUris += $Uri
    if ($Uri -like '*/invoices/relative?*') {
        if ($script:RelationFailRelative) { throw 'Yeu cau GDT that bai (HTTP 500).' }
        return $script:RelativeResponse
    }
    if ($Uri -like '*/invoices/related?*') { return $script:RelatedResponse }
    throw ('Unexpected URI in relation test: ' + $Uri)
}
try {
    $relationConfig = [pscustomobject]@{
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'
        Token = 'test-token'
        RequestDelayMs = 0
        AdaptiveThrottle = $false
        MaxRetries = 0
        HttpTimeoutSeconds = 30
    }
    $relationInvoice = [pscustomobject]@{
        Direction = 'purchase'; Source = 'query'
        SellerTaxCode = '0123456789'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '123'; InvoiceTemplate = '1'
        Status = 3; RelatedChain = 'list-chain'; OriginalInvoiceType = '02'
        OriginalTemplateCode = '1'; OriginalSeries = 'C26TABC'; OriginalNumber = '100'
        OriginalDate = '01/09/2026'; OriginalNote = 'ghi chu'; RelatedInfo = ''
    }

    $relationResult = Get-GdtInvoiceRelation -Config $relationConfig -Invoice $relationInvoice
    Assert-Equal 2 $script:RelationUris.Count 'Status 3 calls relative then related'
    $expectedChain = "1. Hóa đơn có liên quan | 1 | C26TABC | 100 | Thay thế cho hóa đơn có ký hiệu mẫu số 1, ký hiệu hóa đơn C26TOLD, số hóa đơn 50`r`n2. Hóa đơn đang tra cứu | 1 | C26TABC | 123"
    Assert-Equal $expectedChain $relationResult.RelatedChain 'Relative chain text'
    Assert-Equal 'Hóa đơn có Thông báo hóa đơn ngày 02/09/2026. Tính chất Thay thế, lý do sai sót kinh doanh. Cơ quan thuế tiếp nhận.' $relationResult.RelatedInfo 'Related notice summary'
    Assert-Equal '02' $relationResult.OriginalInvoiceType 'Original invoice type copied'

    $script:RelationUris = @()
    $statusSixInvoice = [pscustomobject]@{
        Direction = 'purchase'; Source = 'query'
        SellerTaxCode = '0123456789'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '124'; InvoiceTemplate = '1'
        Status = 6; RelatedChain = 'list-chain'; OriginalInvoiceType = ''
        OriginalTemplateCode = ''; OriginalSeries = ''; OriginalNumber = ''; OriginalDate = ''; OriginalNote = ''
        RelatedInfo = ''
    }
    $statusSixResult = Get-GdtInvoiceRelation -Config $relationConfig -Invoice $statusSixInvoice
    Assert-Equal 1 $script:RelationUris.Count 'Status 6 calls related only'
    Assert-Equal 'Không có thông tin hiển thị' $statusSixResult.RelatedChain 'Status 6 chain text'

    $script:RelationUris = @()
    $statusOneInvoice = [pscustomobject]@{
        Direction = 'purchase'; Source = 'query'
        SellerTaxCode = '0123456789'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '125'; InvoiceTemplate = '1'
        Status = 1; RelatedChain = ''; OriginalInvoiceType = ''
        OriginalTemplateCode = ''; OriginalSeries = ''; OriginalNumber = ''; OriginalDate = ''; OriginalNote = ''
        RelatedInfo = ''
    }
    Assert-Equal $null (Get-GdtInvoiceRelation -Config $relationConfig -Invoice $statusOneInvoice) 'Status 1 has no relation data'
    Assert-Equal 0 $script:RelationUris.Count 'Status 1 makes no relation request'

    $listOnlyResult = Get-GdtInvoiceRelation -Config $relationConfig -Invoice $relationInvoice -ListOnly
    Assert-Equal 0 $script:RelationUris.Count 'ListOnly makes no extra relation request'
    Assert-Equal 'list-chain' $listOnlyResult.RelatedChain 'ListOnly keeps list chain'
    Assert-Equal '' $listOnlyResult.RelatedInfo 'ListOnly has no related info'

    $script:RelationUris = @()
    $script:RelationFailRelative = $true
    $failedResult = Get-GdtInvoiceRelation -Config $relationConfig -Invoice $relationInvoice
    Assert-Equal 2 $script:RelationUris.Count 'Relative failure still calls related'
    Assert-Equal $true ($failedResult.RelatedChain.StartsWith('Lỗi: Không thể lấy chuỗi hóa đơn liên quan')) 'Relative failure writes error into chain cell'
    Assert-Equal $true ($failedResult.RelatedInfo.StartsWith('Hóa đơn có')) 'Related still succeeds after relative failure'
    $script:RelationFailRelative = $false
}
finally {
    Set-Item Function:\Invoke-GdtRequest -Value $originalRelationRequest
    Reset-HddtStopRequest
}

# --- Cac ham dinh dang lien quan ---
Assert-Equal 'Không có thông tin liên quan' (ConvertTo-RelatedInformationText -ResponseText '[]') 'Empty related response'
Assert-Equal 'Không có thông tin liên quan' (ConvertTo-RelatedInformationText -ResponseText '{"mtthdtbssrs":[]}') 'Empty notice container'
$rawRelated = ConvertTo-RelatedInformationText -ResponseText '{"foo":1}'
Assert-Equal $true ($rawRelated -match '"foo"') 'Non notice response keeps JSON text'
Assert-Equal 'Điều chỉnh' (Get-RelatedNoticeNature '2') 'Notice nature maps to Vietnamese'
Assert-Equal '02/09/2026' (ConvertTo-RelatedDate '2026-09-02T00:00:00') 'Related date formats as dd/MM/yyyy'
Assert-Equal $true ((Get-GdtRelationActionHeader -EndpointName 'relative' -Source 'query' -Direction 'purchase').StartsWith('Xem%20h%C3%B3a%20%C4%91%C6%A1n%20li%C3%AAn%20quan%20(')) 'Relative action header'
Assert-Equal ('https://hoadondientu.gdt.gov.vn/api/query/invoices/related?nbmst=0123456789&khmshdon=1&khhdon=C26TABC&shdon=123') (Get-GdtRelationUri -Config $relationConfig -Invoice $relationInvoice -EndpointName 'related') 'Related endpoint uri'

# --- Cau hinh dang nhap ---
$tempLoginEnv = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.env')
$tempEmptyEnv = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.env')
try {
    @'
GDT_TOKEN=
GDT_USERNAME=tester
GDT_PASSWORD=secret-password
INVOICE_DIRECTION=purchase
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
OUTPUT_DIR=output
OUTPUT_XLSX=test.xlsx
'@ | Set-Content -LiteralPath $tempLoginEnv -Encoding UTF8
    $loginModeConfig = Get-HddtConfig -EnvFile $tempLoginEnv -RepositoryRoot $root
    Assert-Equal '' $loginModeConfig.Token 'Login mode leaves token empty'
    Assert-Equal 'tester' $loginModeConfig.Username 'Username parsed from .env'
    Assert-Equal 'secret-password' $loginModeConfig.Password 'Password parsed from .env'
    Assert-Equal $true $loginModeConfig.FetchRelated 'FETCH_RELATED defaults to true'

    @'
GDT_TOKEN=
INVOICE_DIRECTION=purchase
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
OUTPUT_DIR=output
OUTPUT_XLSX=test.xlsx
'@ | Set-Content -LiteralPath $tempEmptyEnv -Encoding UTF8
    $missingAuthRejected = $false
    try { Get-HddtConfig -EnvFile $tempEmptyEnv -RepositoryRoot $root | Out-Null }
    catch { $missingAuthRejected = $true }
    Assert-Equal $true $missingAuthRejected 'Missing token and credentials is rejected'
}
finally {
    Remove-Item -LiteralPath $tempLoginEnv, $tempEmptyEnv -Force -ErrorAction SilentlyContinue
}

# --- Cot chuoi lien quan va style xuong dong ---
$tempWrapXlsx = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.xlsx')
try {
    $wrapSummary = [pscustomobject]@{
        Direction = 'purchase'; Source = 'query'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '123'
        RelatedChain = "dong 1`r`ndong 2"; RelatedInfo = 'thong tin lien quan'
    }
    Export-InvoiceWorkbook -Path $tempWrapXlsx -SummaryRows @($wrapSummary) -DetailRows @() -ErrorRows @() -Overwrite
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $wrapZip = [IO.Compression.ZipFile]::OpenRead($tempWrapXlsx)
    try {
        $wrapStylesEntry = $wrapZip.GetEntry('xl/styles.xml')
        $wrapStylesReader = New-Object IO.StreamReader($wrapStylesEntry.Open())
        try { $wrapStylesText = $wrapStylesReader.ReadToEnd() } finally { $wrapStylesReader.Dispose() }
        Assert-Equal $true ($wrapStylesText -match '<cellXfs count="3">') 'Wrap style adds third cell format'
        Assert-Equal $true ($wrapStylesText -match 'wrapText="1"') 'Wrap style enables text wrapping'

        $wrapSheetEntry = $wrapZip.GetEntry('xl/worksheets/sheet1.xml')
        $wrapSheetReader = New-Object IO.StreamReader($wrapSheetEntry.Open())
        try { $wrapSheetText = $wrapSheetReader.ReadToEnd() } finally { $wrapSheetReader.Dispose() }
        Assert-Equal $true ($wrapSheetText -match '<c r="W2" s="2"') 'RelatedChain cell uses wrap style'
        Assert-Equal $true ($wrapSheetText -match '<c r="AD2" s="2"') 'RelatedInfo cell uses wrap style'
        Assert-Equal $true ($wrapSheetText -match '<c r="A2" t="inlineStr"') 'Normal cell keeps default style'
    }
    finally { $wrapZip.Dispose() }
}
finally {
    Remove-Item -LiteralPath $tempWrapXlsx -Force -ErrorAction SilentlyContinue
}

# --- Kiem tra splat: GET/DELETE khong duoc gui body ---
$script:CapturedSplat = $null
function Invoke-WebRequest {
    param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction, $Body, $ContentType)
    $script:CapturedSplat = $PSBoundParameters
    return [pscustomobject]@{ StatusCode = 200; Content = '{"ok":true}' }
}
try {
    $splatConfig = [pscustomobject]@{
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'
        Token = 'test-token'
        RequestDelayMs = 0
        AdaptiveThrottle = $false
        MaxRetries = 0
        HttpTimeoutSeconds = 30
    }
    Reset-HddtStopRequest

    Invoke-GdtRequest -Config $splatConfig -Uri ($splatConfig.BaseUrl + '/captcha') -SkipAuthorization | Out-Null
    Assert-Equal $false ($script:CapturedSplat.ContainsKey('Body')) 'GET does not send a body'
    Assert-Equal $false ($script:CapturedSplat.ContainsKey('ContentType')) 'GET does not set content type'
    Assert-Equal 'Get' ([string]$script:CapturedSplat['Method']) 'GET verb'
    Assert-Equal $false ($script:CapturedSplat['Headers'].ContainsKey('Authorization')) 'Anonymous GET has no Authorization'

    Invoke-GdtRequest -Config $splatConfig -Uri ($splatConfig.BaseUrl + '/security-taxpayer/authenticate') -Method Post -Body '{"username":"x"}' -SkipAuthorization | Out-Null
    Assert-Equal '{"username":"x"}' ([string]$script:CapturedSplat['Body']) 'POST sends the body'
    Assert-Equal 'Post' ([string]$script:CapturedSplat['Method']) 'POST verb'
    Assert-Equal 'application/json' ([string]$script:CapturedSplat['ContentType']) 'POST content type'

    Invoke-GdtRequest -Config $splatConfig -Uri ($splatConfig.BaseUrl + '/captcha') | Out-Null
    Assert-Equal 'Bearer test-token' ([string]$script:CapturedSplat['Headers']['Authorization']) 'Authorization header uses the token'
}
finally {
    Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue
}

# --- Kiem tra cu phap toan bo script (bat loi encoding/thieu dau) ---
$scriptPaths = @((Join-Path $root 'Invoke-Hddt.ps1'), (Join-Path $root 'Parse-LocalXml.ps1'))
$scriptPaths += @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.ps1' | ForEach-Object { $_.FullName })
foreach ($scriptPath in $scriptPaths) {
    $tokens = $null
    $parseErrors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors) | Out-Null
    Assert-Equal 0 (@($parseErrors).Count) ('No parse error in ' + (Split-Path -Leaf $scriptPath))
}

Write-Host 'All tests passed.' -ForegroundColor Green
