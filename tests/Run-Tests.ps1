Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'src\Logging.ps1')
. (Join-Path $root 'src\Config.ps1')
. (Join-Path $root 'src\Http.ps1')
. (Join-Path $root 'src\InvoiceApi.ps1')
. (Join-Path $root 'src\XmlParser.ps1')
. (Join-Path $root 'src\ExcelExporter.ps1')

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

Write-Host 'All tests passed.' -ForegroundColor Green
