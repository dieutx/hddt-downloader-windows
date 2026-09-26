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
. (Join-Path $root 'src\BrowserProfile.ps1')
. (Join-Path $root 'src\XmlScheduler.ps1')

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

# Dữ liệu XML đôi khi có trường thuế không phải số; exporter không được ép lỗi
# và làm mất toàn bộ workbook.
$badSummary = [pscustomobject]@{
    Direction = 'purchase'; Source = 'query'; InvoiceId = 'bad-1'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '123'
    InvoiceDate = [datetime]'2026-09-20'; Currency = 'VND'; ExchangeRate = 1; SellerName = 'Seller'; SellerTaxCode = '0101111111'
    SellerAddress = ''; SellerSigningTime = $null; TaxAuthorityCode = ''; TaxAuthoritySigningTime = $null
    BuyerName = 'Buyer'; BuyerTaxCode = '0302222222'; BuyerAddress = ''; ProviderTaxCode = ''; TaxAmount = 0
}
$badDetail = [pscustomobject]@{
    Direction = 'purchase'; InvoiceSummary = $badSummary; LineNumber = '1'; Nature = ''; ProductCode = 'P'; Description = 'P'
    Unit = ''; Quantity = 1; UnitPrice = 100; DiscountRate = ''; DiscountAmount = ''; TaxType = ''; TaxRate = 'N/A'
    AmountBeforeTax = 100; TaxAmount = 'N/A'; AmountWithTax = 'N/A'
}
$badDetailRows = @(New-ExcelDetailRows @($badDetail))
Assert-Equal 1 $badDetailRows.Count 'Malformed nonnumeric tax values remain exportable'
Assert-Equal 'N/A' $badDetailRows[0].Row.Cells[27].Value 'Malformed tax text is preserved for review'

$soldXmlSummary = [pscustomobject]@{
    Direction = 'sold'; Source = 'query'; InvoiceId = 'sold-1'; InvoiceSeries = 'C26SOLD'; InvoiceNumber = '9'
    InvoiceDate = [datetime]'2026-09-20'; InvoiceDateText = '2026-09-20'; Currency = 'VND'; ExchangeRate = 1
    SellerName = 'Actual buyer in XML'; SellerTaxCode = '0101111111'; SellerAddress = 'Buyer address'; SellerSigningTime = [datetime]'2026-09-21'
    TaxAuthorityCode = 'CQT-1'; TaxAuthoritySigningTime = [datetime]'2026-09-22'
    BuyerName = 'Queried seller in XML'; BuyerTaxCode = '0302222222'; BuyerAddress = 'Seller address'; ProviderTaxCode = ''; TaxAmount = 10
}
$soldXmlDetail = [pscustomobject]@{ InvoiceSummary = $soldXmlSummary; LineNumber = '1'; ProductCode = 'P'; Description = 'Item'; Unit = ''; Quantity = 1; UnitPrice = 10; DiscountRate = ''; DiscountAmount = ''; TaxRate = '10%'; TaxAmount = 1; AmountBeforeTax = 10; AmountWithTax = 11 }
$soldXmlRows = @(New-ExcelXmlRows @($soldXmlDetail))
Assert-Equal 'Actual buyer in XML' $soldXmlRows[0].Row.Cells[7].Value 'Sold XML sheet keeps NBan in its source-mapped buyer column'
Assert-Equal 'Queried seller in XML' $soldXmlRows[0].Row.Cells[13].Value 'Sold XML sheet keeps NMua in its source-mapped seller column'

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
Assert-Equal 15 (Get-RetryDelaySeconds -StatusCode 429 -Attempt 1) 'First 429 backoff'
Assert-Equal 60 (Get-RetryDelaySeconds -StatusCode 429 -Attempt 3) 'Third 429 backoff'
Assert-Equal 600 (Get-RetryDelaySeconds -StatusCode 429 -Attempt 8) '429 backoff caps at ten minutes'
Assert-Equal 45 (Get-RetryDelaySeconds -StatusCode 429 -Attempt 1 -RetryAfterSeconds 45) 'Retry-After precedence'
Assert-Equal 12 (Get-429RetryLimit) '429 retry limit is independent of MAX_RETRIES'

$originalGdtRequest = ${function:Invoke-GdtRequest}
$script:IndexRequestCount = 0
Set-Item Function:\Invoke-GdtRequest -Value {
    param($Config, $Uri, $AsBytes, [string]$RequestProfile)
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

# Một kỳ lỗi không được làm mất các kỳ còn lại của cùng nguồn.
$script:PeriodRequestCount = 0
Set-Item Function:\Invoke-GdtRequest -Value {
    param($Config, $Uri, [string]$RequestProfile)
    $script:PeriodRequestCount++
    if ($Uri -match '31%2F08%2F2026') { throw 'simulated period failure' }
    return '{"datas":[],"state":""}'
}
try {
    $periodConfig = [pscustomobject]@{
        IncludeRegular = $true; IncludeSco = $false
        FromDate = [datetime]'2026-08-20'; ToDate = [datetime]'2026-10-02'
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'; PageSize = 50
    }
    $periodRows = @(Get-GdtInvoiceIndex -Config $periodConfig -Direction 'purchase')
    $periodErrors = @(Get-GdtIndexErrors)
    Assert-Equal 3 $script:PeriodRequestCount 'Each monthly period is attempted independently'
    Assert-Equal 0 $periodRows.Count 'Failed period does not fabricate rows'
    Assert-Equal 1 $periodErrors.Count 'One period failure creates one structured error'
    Assert-Equal '20/08/2026-31/08/2026' $periodErrors[0].Period 'Period failure retains the failed range'
}
finally { Set-Item Function:\Invoke-GdtRequest -Value $originalGdtRequest }

# Phản hồi HTTP 200 nhưng thiếu datas phải được coi là lỗi, không phải trang rỗng.
Set-Item Function:\Invoke-GdtRequest -Value {
    param($Config, $Uri, [string]$RequestProfile)
    return '{"state":""}'
}
try {
    $badPayloadConfig = [pscustomobject]@{
        IncludeRegular = $true; IncludeSco = $false
        FromDate = [datetime]'2026-09-01'; ToDate = [datetime]'2026-09-01'
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'; PageSize = 50
    }
    Get-GdtInvoiceIndex -Config $badPayloadConfig -Direction 'purchase' | Out-Null
    Assert-Equal 1 @(Get-GdtIndexErrors).Count 'Successful HTTP response without datas is rejected'
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

    $preservedPath = Join-Path $tempXmlDirectory 'preserved.xml'
    Set-Content -LiteralPath $preservedPath -Value '<HDon />' -Encoding UTF8
    $invalidMemory = New-Object IO.MemoryStream
    $invalidArchive = New-Object IO.Compression.ZipArchive($invalidMemory, [IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        $invalidEntry = $invalidArchive.CreateEntry('invoice.xml')
        $invalidWriter = New-Object IO.StreamWriter($invalidEntry.Open())
        try { $invalidWriter.Write('<HDon>') } finally { $invalidWriter.Dispose() }
    }
    finally { $invalidArchive.Dispose() }
    $invalidRejected = $false
    try { Expand-InvoiceXmlBytes -Bytes $invalidMemory.ToArray() -DestinationDirectory $tempXmlDirectory -FileNamePrefix 'preserved' | Out-Null }
    catch { $invalidRejected = $true }
    finally { $invalidMemory.Dispose() }
    Assert-Equal $true $invalidRejected 'Malformed XML download is rejected'
    Assert-Equal '<HDon />' ((Get-Content -LiteralPath $preservedPath -Raw).Trim()) 'Malformed download does not overwrite a good resume file'
    Remove-Item -LiteralPath $preservedPath -Force -ErrorAction SilentlyContinue
}
finally {
    $zipMemory.Dispose()
    Remove-Item -LiteralPath $tempXmlDirectory -Recurse -Force -ErrorAction SilentlyContinue
}

function Read-TestZipEntryText {
    param($Archive, [string]$Name)
    $entry = $Archive.GetEntry($Name)
    if ($null -eq $entry) { throw "Thiếu entry trong workbook: $Name" }
    $reader = New-Object IO.StreamReader($entry.Open())
    try { return $reader.ReadToEnd() }
    finally { $reader.Dispose() }
}

function Get-TestCellNode {
    param($SheetDocument, [string]$Reference)
    $manager = New-Object Xml.XmlNamespaceManager($SheetDocument.NameTable)
    $manager.AddNamespace('x', $script:SpreadsheetNamespace)
    return $SheetDocument.SelectSingleNode("/x:worksheet/x:sheetData/x:row/x:c[@r='$Reference']", $manager)
}

function Get-TestCellValue {
    param($SheetDocument, [string]$Reference)
    $cell = Get-TestCellNode $SheetDocument $Reference
    if ($null -eq $cell) { return '' }
    $manager = New-Object Xml.XmlNamespaceManager($SheetDocument.NameTable)
    $manager.AddNamespace('x', $script:SpreadsheetNamespace)
    $valueNode = $cell.SelectSingleNode('x:is/x:t', $manager)
    if ($null -eq $valueNode) { $valueNode = $cell.SelectSingleNode('x:v', $manager) }
    if ($null -eq $valueNode) { return '' }
    return [string]$valueNode.InnerText
}

$tempXlsx = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.xlsx')
try {
    # Add a known provider lookup to exercise the data-only link projection.
    $fixtureSummary = $parsed.Summary
    $fixtureSummary.ProviderTaxCode = '0100109106'
    $fixtureSummary.InvoiceAdditionalFields = @([pscustomobject]@{ TTRuong = 'MaTraCuu'; DLieu = 'LOOK-001' })
    Add-Member -InputObject $fixtureSummary -NotePropertyName GdtIndex -NotePropertyValue ([pscustomobject]@{
        id = 'INV-LINK-001'; msttcgp = '0100109106'; nbmst = '0101111111'; mhdon = 'ABC123'
        tthai = 2; ttxly = 2; tgtcthue = 250000; tgtkcthue = 0; tgtthue = 24000
        thttltsuat = @([pscustomobject]@{ tsuat = '10'; thtien = '250000'; tthue = '24000'; gttsuat = '0' })
        cttkhac = @([pscustomobject]@{ TTRuong = 'MaTraCuu'; DLieu = 'LOOK-001' })
    }) -Force
    $testError = [pscustomobject]@{
        Direction = 'purchase'; Source = 'query'; SellerTaxCode = '0101111111'
        InvoiceTemplate = '1'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '123'; InvoiceDate = '2026-09-20'
        Stage = 'Tải XML'; Endpoint = 'https://hoadondientu.gdt.gov.vn/api/query/invoices/export-xml'
        StatusCode = 504; Attempts = 2; RetryAfterSeconds = 0
        Error = 'Authorization: Bearer super-secret-token'; FinalResult = 'Không tải được'; Note = 'partial'
    }
    $exportSummary = Export-InvoiceWorkbook -Path $tempXlsx -SummaryRows @($fixtureSummary) -DetailRows @($parsed.Details) -ErrorRows @($testError) -Overwrite
    Assert-Equal 1 $exportSummary.SummaryRows 'Export summary counts summary rows'
    Assert-Equal 1 $exportSummary.LinksResolved 'Export summary counts resolved lookup links'
    Assert-Equal 0 $exportSummary.LinksMissing 'Export summary counts missing lookup links'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($tempXlsx)
    try {
        Assert-Equal 1 @($zip.Entries | Where-Object FullName -eq 'xl/workbook.xml').Count 'Workbook package entry'
        Assert-Equal 1 @($zip.Entries | Where-Object FullName -eq 'xl/styles.xml').Count 'Styles package entry'
        # Excel không mở được workbook khi thiếu '_rels/.rels' (part quan hệ gốc).
        Assert-Equal 1 @($zip.Entries | Where-Object FullName -eq '_rels/.rels').Count 'Root relationship package entry'
        Assert-Equal 0 @($zip.Entries | Where-Object { $_.FullName -match '\\' }).Count 'No package entry keeps a path separator'
        $rootRelationships = Read-TestZipEntryText $zip '_rels/.rels'
        Assert-Equal $true ($rootRelationships -match 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"') 'Root relationship type'
        Assert-Equal $true ($rootRelationships -match 'Target="xl/workbook.xml"') 'Root relationship points at the workbook part'
        Assert-Equal 8 @($zip.Entries | Where-Object FullName -like 'xl/worksheets/sheet*.xml').Count 'Worksheet package count'
        Assert-Equal 0 @($zip.Entries | Where-Object FullName -like 'xl/tables/table*.xml').Count 'Data-only workbook has no table parts'
        Assert-Equal 0 @($zip.Entries | Where-Object FullName -match 'vbaProject|MENU|Thamkhao').Count 'No VBA or menu sheets are packaged'

        [xml]$workbookDocument = Read-TestZipEntryText $zip 'xl/workbook.xml'
        $workbookManager = New-Object Xml.XmlNamespaceManager($workbookDocument.NameTable)
        $workbookManager.AddNamespace('x', $script:SpreadsheetNamespace)
        $sheetNodes = @($workbookDocument.SelectNodes('/x:workbook/x:sheets/x:sheet', $workbookManager))
        $expectedSheetNames = @('TongHopHD_Mua', 'ChiTietHD_Mua', 'ChiTietHD_Mua_XML', 'TongHopHD_Ban', 'ChiTietHD_Ban', 'ChiTietHD_Ban_XML', 'BaoCao_LoiTaiHD', 'LinkTraCuu')
        Assert-Equal ($expectedSheetNames -join '|') (($sheetNodes | ForEach-Object { $_.GetAttribute('name') }) -join '|') 'Source-style sheet order and names'
        Assert-Equal $script:SourceSummaryHeaders.Count $script:SourceSummaryWidths.Count 'Summary headers and widths have equal lengths'
        Assert-Equal $script:SourceDetailHeaders.Count $script:SourceDetailWidths.Count 'Detail headers and widths have equal lengths'
        Assert-Equal $script:SourceXmlHeadersPurchase.Count $script:SourceXmlWidths.Count 'XML headers and widths have equal lengths'
        Assert-Equal $script:SourceXmlHeadersSold.Count $script:SourceXmlWidths.Count 'Sold XML headers and widths have equal lengths'
        Assert-Equal $script:SourceLookupHeaders.Count $script:SourceLookupWidths.Count 'Lookup sheet headers and widths have equal lengths'
        Assert-Equal 'Công đoạn lỗi' $script:SourceErrorHeaders[9] 'Error report uses source stage header'

        $sheetDocuments = @{}
        for ($sheetNumber = 1; $sheetNumber -le 8; $sheetNumber++) {
            [xml]$sheetDocument = Read-TestZipEntryText $zip ("xl/worksheets/sheet{0}.xml" -f $sheetNumber)
            $sheetDocuments[$sheetNumber] = $sheetDocument
            $manager = New-Object Xml.XmlNamespaceManager($sheetDocument.NameTable)
            $manager.AddNamespace('x', $script:SpreadsheetNamespace)
            Assert-Equal 0 @($sheetDocument.SelectNodes('/x:worksheet/x:tableParts/x:tablePart', $manager)).Count ("Sheet $sheetNumber has no table part")
            $expectedFilterCount = if ($sheetNumber -eq 7) { 1 } else { 0 }
            Assert-Equal $expectedFilterCount @($sheetDocument.SelectNodes('/x:worksheet/x:autoFilter', $manager)).Count ("Sheet $sheetNumber AutoFilter contract")
        }

        $summarySheet = $sheetDocuments[1]
        $summaryDate = Get-TestCellNode $summarySheet 'F3'
        Assert-Equal '6' $summaryDate.GetAttribute('s') 'Summary invoice date uses date style'
        Assert-Equal ([datetime]'2026-09-20').ToOADate().ToString('0.###############', [Globalization.CultureInfo]::InvariantCulture) (Get-TestCellValue $summarySheet 'F3') 'Summary invoice date is a numeric Excel date'
        Assert-Equal 'https://tracuuhoadon.vetc.com.vn/' (Get-TestCellValue $summarySheet 'BC3') 'Summary lookup link'
        Assert-Equal 'LOOK-001' (Get-TestCellValue $summarySheet 'BD3') 'Summary lookup code'

        $detailSheet = $sheetDocuments[2]
        $detailRate = Get-TestCellNode $detailSheet 'Y3'
        Assert-Equal '8' $detailRate.GetAttribute('s') 'Detail tax rate uses percentage style'
        Assert-Equal '0.1' (Get-TestCellValue $detailSheet 'Y3') 'Detail tax rate is stored as a fraction'
        Assert-Equal 'https://tracuuhoadon.vetc.com.vn/' (Get-TestCellValue $detailSheet 'AF3') 'Detail lookup link'

        $xmlSheet = $sheetDocuments[3]
        Assert-Equal '2026-09-20' (Get-TestCellValue $xmlSheet 'D3') 'XML sheet keeps source invoice date text'
        Assert-Equal '10%' (Get-TestCellValue $xmlSheet 'Y3') 'XML sheet keeps source tax-rate text'
        Assert-Equal 'https://tracuuhoadon.vetc.com.vn/' (Get-TestCellValue $xmlSheet 'AD3') 'XML lookup link'
        Assert-Equal 1 @($zip.Entries | Where-Object FullName -eq 'xl/worksheets/_rels/sheet1.xml.rels').Count 'Summary hyperlink relationship part'
        Assert-Equal 1 @($zip.Entries | Where-Object FullName -eq 'xl/worksheets/_rels/sheet2.xml.rels').Count 'Detail hyperlink relationship part'
        Assert-Equal 1 @($zip.Entries | Where-Object FullName -eq 'xl/worksheets/_rels/sheet3.xml.rels').Count 'XML hyperlink relationship part'

        $errorSheet = $sheetDocuments[7]
        $errorText = Get-TestCellValue $errorSheet 'M2'
        Assert-Equal $true ($errorText -match '\[REDACTED\]') 'Authorization secrets are redacted from the error sheet'
        Assert-Equal $false ($errorText -match 'super-secret-token') 'Authorization token is absent from the error sheet'

        # Sheet LinkTraCuu phải có trong workbook và giữ nguyên bảng định tuyến
        # của TaiHoaDonDienTu: cột B (MST) + cột D (link) sinh ra link tra cứu.
        $lookupSheet = $sheetDocuments[8]
        Assert-Equal 'Tên tổ chức' (Get-TestCellValue $lookupSheet 'A1') 'Lookup sheet organization header'
        Assert-Equal 'MST' (Get-TestCellValue $lookupSheet 'B1') 'Lookup sheet provider MST header'
        Assert-Equal 'Link tra cứu' (Get-TestCellValue $lookupSheet 'D1') 'Lookup sheet link header'
        Assert-Equal 'Tên trường mã tra cứu' (Get-TestCellValue $lookupSheet 'E1') 'Lookup sheet lookup field header'
        $lookupManager = New-Object Xml.XmlNamespaceManager($lookupSheet.NameTable)
        $lookupManager.AddNamespace('x', $script:SpreadsheetNamespace)
        $lookupRowCount = @($lookupSheet.SelectNodes('/x:worksheet/x:sheetData/x:row', $lookupManager)).Count
        Assert-Equal (1 + @(Get-TaiHoaDonDienTuLookupSheetRows).Count) $lookupRowCount 'Lookup sheet keeps every reference row'
        $lookupSheetText = $lookupSheet.OuterXml
        Assert-Equal $true ($lookupSheetText -match '0100109106') 'Lookup sheet keeps provider MSTs'
        Assert-Equal $true ($lookupSheetText -match 'https://tracuuhoadon.vetc.com.vn/') 'Lookup sheet keeps provider links'
        Assert-Equal 1 @($zip.Entries | Where-Object FullName -eq 'xl/worksheets/_rels/sheet8.xml.rels').Count 'Lookup sheet hyperlink relationship part'

        $stylesDocument = [xml](Read-TestZipEntryText $zip 'xl/styles.xml')
        $stylesManager = New-Object Xml.XmlNamespaceManager($stylesDocument.NameTable)
        $stylesManager.AddNamespace('x', $script:SpreadsheetNamespace)
        $headerFont = $stylesDocument.SelectSingleNode('/x:styleSheet/x:fonts/x:font[2]', $stylesManager)
        $headerFontOrder = (($headerFont.ChildNodes | ForEach-Object { $_.LocalName }) -join ',')
        Assert-Equal 'sz,name,family' $headerFontOrder 'Header font element order follows Open XML schema'
    }
    finally { $zip.Dispose() }
}
finally {
    Remove-Item -LiteralPath $tempXlsx -Force -ErrorAction SilentlyContinue
}

# --- Bang tra cuu: link sinh ra tu sheet LinkTraCuu ---
Assert-Equal 'https://tracuuhoadon.vetc.com.vn/' (Get-TaiHoaDonDienTuLookupLink '0100109106') 'Provider link from the lookup table'
Assert-Equal '' (Get-TaiHoaDonDienTuLookupLink '0999999999') 'Unknown provider has no link'
Assert-Equal 'https://giaothongso.com.vn/tra-cuu-hoa-don-mtc/' (Get-TaiHoaDonDienTuSellerLookupLink '0109266456') 'Seller specific link from column C'
Assert-Equal 'https://gsm-einvoice.hilo.com.vn/' (Get-TaiHoaDonDienTuLookupLink '0110269067') 'Hilo provider link'
Assert-Equal $true (Test-TaiHoaDonDienTuLookupField 'MaTraCuu') 'Lookup field name matches'
Assert-Equal $true (Test-TaiHoaDonDienTuLookupField 'Hilo-SearchKey') 'Hilo lookup field name matches'
Assert-Equal $false (Test-TaiHoaDonDienTuLookupField 'KhongPhaiTenTruong') 'Unknown lookup field name is rejected'
Assert-Equal 'Tất cả' (Get-ExcelStatusLabel 0) 'Status label for 0'
Assert-Equal 'Hóa đơn mới' (Get-ExcelStatusLabel 1) 'Status label for 1'
Assert-Equal 'Tổng cục thuế đã nhận' (Get-ExcelValidationLabel 0) 'Validation label for ttxly 0'
Assert-Equal 'Đã cấp mã hóa đơn' (Get-ExcelValidationLabel 5) 'Validation label for ttxly 5'
Assert-Equal 'Tổng cục thuế đã nhận hóa đơn có mã khởi tạo từ máy tính tiền' (Get-ExcelValidationLabel 8) 'Validation label for ttxly 8'
Assert-Equal '' (Get-ExcelValidationLabel 9) 'Validation label beyond the lookup table is empty'
Assert-Equal 'https://tracuuhoadon.vetc.com.vn/' (Get-ExcelLookupLink -ProviderTaxCode '0100109106' -SellerTaxCode '0101111111' -Mode summary) 'Summary link from provider MST'
# GDT không trả MSTTCGP thì vẫn tra được bằng MST của bên liên quan.
Assert-Equal 'https://tracuuhoadon.vetc.com.vn/' (Get-ExcelLookupLink -ProviderTaxCode '' -SellerTaxCode '0100109106' -CounterPartyTaxCodes @('0101111111', '0100109106') -Mode summary) 'Summary link falls back to the counterparty MST'
Assert-Equal 'https://giaothongso.com.vn/tra-cuu-hoa-don-mtc/' (Get-ExcelLookupLink -ProviderTaxCode '' -SellerTaxCode '0109266456' -CounterPartyTaxCodes @('0109266456') -Mode summary) 'Seller specific link wins for the counterparty MST'
Assert-Equal 'Khong co link tra cuu' (Get-ExcelLookupLink -ProviderTaxCode '' -SellerTaxCode '0101111111' -CounterPartyTaxCodes @('0101111111') -Mode summary) 'Unknown counterparty keeps the source wording'
Assert-Equal 'Khong tim thay link tra cuu' (Get-ExcelLookupLink -ProviderTaxCode '' -SellerTaxCode '0101111111' -CounterPartyTaxCodes @('0101111111') -Mode xml) 'Unknown counterparty keeps the source wording in the XML sheet'
Assert-Equal 'https://gsm-einvoice.hilo.com.vn/' (Get-ExcelLookupLink -ProviderTaxCode '' -SellerTaxCode '0110269067-002' -CounterPartyTaxCodes @('0110269067-002') -Mode summary) 'Hilo seller without provider MST keeps its link'
# Hàng tổng hợp không có MSTTCGP vẫn ghi link tìm được từ MST đối tác.
$counterPartySummary = [pscustomobject]@{
    Direction = 'purchase'; Source = 'query'; InvoiceSeries = 'C26ABC'; InvoiceNumber = '1'
    Status = 1; ValidationStatus = 0; SellerTaxCode = '0100109106'; BuyerTaxCode = '0101111111'
    GdtIndex = [pscustomobject]@{ nbmst = '0100109106'; nmmst = '0101111111'; tthai = 1; ttxly = 0 }
}
$counterPartyRows = @(New-ExcelSummaryRows @($counterPartySummary))
Assert-Equal 'https://tracuuhoadon.vetc.com.vn/' $counterPartyRows[0].Row.Cells[55].Value 'Summary row keeps the counterparty lookup link'
Assert-Equal 'Tổng cục thuế đã nhận' $counterPartyRows[0].Row.Cells[53].Value 'Summary row maps ttxly through the reference labels'

# --- Nap bang tra cuu tu workbook cua nguoi dung (LOOKUP_TABLE_XLSX) ---
$lookupRoundTrip = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.xlsx')
try {
    $null = Export-InvoiceWorkbook -Path $lookupRoundTrip -SummaryRows @() -DetailRows @() -ErrorRows @() -Overwrite
    $builtRows = @(Read-TaiHoaDonDienTuLookupSheet -Path $lookupRoundTrip)
    Assert-Equal 118 $builtRows.Count 'Exported lookup sheet is readable again'
    Assert-Equal 'MST' $builtRows[0]['B'] 'Lookup sheet header round-trips'
    $importedCount = Import-ExcelLookupTable -Path $lookupRoundTrip
    Assert-Equal 117 $importedCount 'Every reference row is imported back'
    Assert-Equal 117 @(Get-TaiHoaDonDienTuLookupSheetRows).Count 'Re-importing the same table does not duplicate rows'
    Assert-Equal 'https://tracuuhoadon.vetc.com.vn/' (Get-TaiHoaDonDienTuLookupLink '0100109106') 'Provider link survives the round-trip'
    Assert-Equal 'https://giaothongso.com.vn/tra-cuu-hoa-don-mtc/' (Get-TaiHoaDonDienTuSellerLookupLink '0109266456') 'Seller link survives the round-trip'

    $missingLookupTableRejected = $false
    try { Import-ExcelLookupTable -Path (Join-Path ([IO.Path]::GetTempPath()) 'hddt-khong-ton-tai.xlsx') }
    catch { $missingLookupTableRejected = $true }
    Assert-Equal $true $missingLookupTableRejected 'Missing lookup table is reported'
}
finally {
    Remove-Item -LiteralPath $lookupRoundTrip -Force -ErrorAction SilentlyContinue
}

# Gói thiếu thành phần bắt buộc phải bị chặn trước khi ghi đè workbook cũ,
# thay vì tạo ra file mà Excel không mở được.
$incompletePackage = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.xlsx')
try {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $incompleteStream = [IO.File]::Open($incompletePackage, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $incompleteArchive = New-Object IO.Compression.ZipArchive($incompleteStream, [IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            $incompleteEntry = $incompleteArchive.CreateEntry('[Content_Types].xml')
            $incompleteWriter = New-Object IO.StreamWriter($incompleteEntry.Open())
            try { $incompleteWriter.Write('<Types />') } finally { $incompleteWriter.Dispose() }
        }
        finally { $incompleteArchive.Dispose() }
    }
    finally { $incompleteStream.Dispose() }
    $incompletePackageRejected = $false
    try { Assert-ExcelPackage -Path $incompletePackage -Parts @() }
    catch { $incompletePackageRejected = $true }
    Assert-Equal $true $incompletePackageRejected 'Package without the root relationship part is rejected'
}
finally {
    Remove-Item -LiteralPath $incompletePackage -Force -ErrorAction SilentlyContinue
}

$tempEnv = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.env')
try {
    @'
GDT_USERNAME=tester
GDT_PASSWORD=secret-password
INVOICE_DIRECTION=purchase ; inline comment
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
OUTPUT_DIR=output
OUTPUT_XLSX=test.xlsx
'@ | Set-Content -LiteralPath $tempEnv -Encoding UTF8
    $config = Get-HddtConfig -EnvFile $tempEnv -RepositoryRoot $root
    Assert-Equal '' $config.Token 'Runtime token starts empty until automatic login'
    Assert-Equal 'tester' $config.Username 'Username parsed from .env'
    Assert-Equal 'secret-password' $config.Password 'Password parsed from .env'
    Assert-Equal 1 $config.Directions.Count 'Direction count'
    Assert-Equal 'purchase' $config.Directions[0] 'Direction value'
    Assert-Equal 600 $config.RequestDelayMs 'Default global request delay'
    Assert-Equal 'info' $config.LogLevel 'Default log level'
    Assert-Equal 1 $config.ProgressEvery 'Default progress interval'
    Assert-Equal $true $config.LogToFile 'Default file logging'
    Assert-Equal $true $config.AdaptiveThrottle 'Default adaptive throttling'
    Assert-Equal $false $config.RedownloadXml 'Default XML resume mode'
    Assert-Equal $null $config.ProxyUri 'No proxy is configured by default'
}
finally {
    Remove-Item -LiteralPath $tempEnv -Force -ErrorAction SilentlyContinue
}

# --- Proxy config: compact form, validation va credentials ---
$tempProxyEnv = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.env')
try {
    @'
GDT_USERNAME=tester
GDT_PASSWORD=secret-password
INVOICE_DIRECTION=purchase
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
OUTPUT_DIR=output
OUTPUT_XLSX=test.xlsx
PROXY_URL=proxy.example:8080:proxy-user:proxy-pass
'@ | Set-Content -LiteralPath $tempProxyEnv -Encoding UTF8
    $proxyConfig = Get-HddtConfig -EnvFile $tempProxyEnv -RepositoryRoot $root
    Assert-Equal 'http://proxy.example:8080/' $proxyConfig.ProxyUri.AbsoluteUri 'Compact proxy host:port is normalized'
    Assert-Equal 'proxy-user' $proxyConfig.ProxyUsername 'Compact proxy username is extracted'
    Assert-Equal 'proxy-pass' $proxyConfig.ProxyPassword 'Compact proxy password is extracted'

    @'
GDT_USERNAME=tester
GDT_PASSWORD=secret-password
INVOICE_DIRECTION=purchase
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
OUTPUT_DIR=output
OUTPUT_XLSX=test.xlsx
PROXY_URL=http://user:password@proxy.example:8080
'@ | Set-Content -LiteralPath $tempProxyEnv -Encoding UTF8
    $embeddedProxyRejected = $false
    try { Get-HddtConfig -EnvFile $tempProxyEnv -RepositoryRoot $root | Out-Null }
    catch { $embeddedProxyRejected = $true }
    Assert-Equal $true $embeddedProxyRejected 'Proxy URI userinfo is rejected'

    @'
GDT_USERNAME=tester
GDT_PASSWORD=secret-password
INVOICE_DIRECTION=purchase
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
OUTPUT_DIR=output
OUTPUT_XLSX=test.xlsx
PROXY_URL=http://proxy.example:8080
PROXY_PASSWORD=orphan-password
'@ | Set-Content -LiteralPath $tempProxyEnv -Encoding UTF8
    $orphanProxyPasswordRejected = $false
    try { Get-HddtConfig -EnvFile $tempProxyEnv -RepositoryRoot $root | Out-Null }
    catch { $orphanProxyPasswordRejected = $true }
    Assert-Equal $true $orphanProxyPasswordRejected 'Proxy password without username is rejected'
}
finally {
    Remove-Item -LiteralPath $tempProxyEnv -Force -ErrorAction SilentlyContinue
}

# --- Interactive config: secure username/password prompts and no secret echo ---
$tempInteractivePath = Join-Path ([IO.Path]::GetTempPath()) ('hddt-interactive-' + [guid]::NewGuid().ToString('N') + '.env')
$script:InteractiveAnswers = [System.Collections.Queue]::new()
foreach ($answer in @('interactive-user', 'interactive-pass', 'purchase', '01/09/2026', '30/09/2026', 'output', 'interactive.xlsx', 'true', 'false', 'false', 'false', 'true', 'http://proxy.example:8080', 'proxy-user', 'proxy-pass')) {
    $script:InteractiveAnswers.Enqueue([string]$answer)
}
$script:InteractivePrompts = New-Object System.Collections.Generic.List[string]
function Read-Host {
    param([string]$Prompt, [switch]$AsSecureString)
    $script:InteractivePrompts.Add([string]$Prompt)
    $answer = [string]$script:InteractiveAnswers.Dequeue()
    if ($AsSecureString) {
        $secure = New-Object Security.SecureString
        foreach ($character in $answer.ToCharArray()) { $secure.AppendChar($character) }
        $secure.MakeReadOnly()
        return $secure
    }
    return $answer
}
try {
    $interactiveConfig = Get-HddtConfig -EnvFile $tempInteractivePath -RepositoryRoot $root -Interactive
    Assert-Equal '' $interactiveConfig.Token 'Interactive config leaves runtime token empty'
    Assert-Equal 'interactive-user' $interactiveConfig.Username 'Interactive username is retained in memory'
    Assert-Equal 'interactive-pass' $interactiveConfig.Password 'Interactive password is retained in memory'
    Assert-Equal 'http://proxy.example:8080/' $interactiveConfig.ProxyUri.AbsoluteUri 'Interactive proxy URL is parsed'
    Assert-Equal $false (($script:InteractivePrompts -join '|') -match 'interactive-pass') 'Interactive password is not echoed in a prompt'
    Assert-Equal $false (($script:InteractivePrompts -join '|') -match 'proxy-pass') 'Interactive proxy password is not echoed in a prompt'
}
finally {
    Remove-Item Function:\Read-Host -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $tempInteractivePath -Force -ErrorAction SilentlyContinue
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
    param($Config, $Uri, [string]$Method, [string]$Body, [switch]$SkipAuthorization, [hashtable]$ExtraHeaders, [switch]$AsBytes, [string]$ContentType, [string]$RequestProfile)
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

    $script:LoginCalls = @()
    $script:LoginAuthBodies = @()
    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, [string]$Method, [string]$Body, [switch]$SkipAuthorization, [hashtable]$ExtraHeaders, [switch]$AsBytes, [string]$ContentType, [string]$RequestProfile)
        $script:LoginCalls += [pscustomobject]@{ Uri = $Uri; Method = $Method; Body = $Body }
        if ($Uri -like '*/captcha') { return (New-TestCaptchaResponse -KeywordIndexes @(19, 4, 2) -Positions @(30, 10, 20)) }
        throw 'Đăng nhập GDT không thành công (HTTP 401). Kiểm tra GDT_USERNAME/GDT_PASSWORD hoặc quyền truy cập tài khoản.'
    }
    $login401Error = ''
    try { Invoke-GdtLogin -Config $loginConfig | Out-Null }
    catch { $login401Error = $_.Exception.Message }
    Assert-Equal $true ($login401Error -match 'HTTP 401') 'Authentication 401 has a clear login error'
    Assert-Equal 1 @($script:LoginCalls | Where-Object Uri -like '*/authenticate').Count 'Authentication 401 is not retried three times'
}
finally { Set-Item Function:\Invoke-GdtRequest -Value $originalLoginRequest }

# --- Chuoi hoa don lien quan (relative/related) ---
$originalRelationRequest = ${function:Invoke-GdtRequest}
$script:RelationUris = @()
$script:RelationFailRelative = $false
$script:RelativeResponse = '[{"khmshdon":"1","khhdon":"C26TABC","shdon":"100","khmshdgoc":"1","khhdgoc":"C26TOLD","shdgoc":"50","tthai":2}]'
$script:RelatedResponse = '{"mtthdtbssrs":[{"ten":"Thông báo hóa đơn","ngay":"2026-09-02T00:00:00","ldo":"sai sót kinh doanh","loai":"3","kqtnhan":"1"}]}'
Set-Item Function:\Invoke-GdtRequest -Value {
    param($Config, $Uri, [string]$Method, [string]$Body, [switch]$SkipAuthorization, [hashtable]$ExtraHeaders, [switch]$AsBytes, [string]$ContentType, [string]$RequestProfile)
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
Assert-Equal '01/09/2026' (ConvertTo-RelatedDate '01/09/2026') 'Related dd/MM/yyyy date keeps Vietnamese order'
Assert-Equal $true ((Get-GdtRelationActionHeader -EndpointName 'relative' -Source 'query' -Direction 'purchase').StartsWith('Xem%20h%C3%B3a%20%C4%91%C6%A1n%20li%C3%AAn%20quan%20(')) 'Relative action header'
Assert-Equal ('https://hoadondientu.gdt.gov.vn/api/query/invoices/related?nbmst=0123456789&khmshdon=1&khhdon=C26TABC&shdon=123') (Get-GdtRelationUri -Config $relationConfig -Invoice $relationInvoice -EndpointName 'related') 'Related endpoint uri'

# --- Cau hinh dang nhap ---
$tempLoginEnv = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.env')
$tempEmptyEnv = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.env')
try {
    @'
GDT_USERNAME=tester
GDT_PASSWORD=secret-password
INVOICE_DIRECTION=purchase
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
OUTPUT_DIR=output
OUTPUT_XLSX=test.xlsx
'@ | Set-Content -LiteralPath $tempLoginEnv -Encoding UTF8
    $loginModeConfig = Get-HddtConfig -EnvFile $tempLoginEnv -RepositoryRoot $root
    Assert-Equal '' $loginModeConfig.Token 'Account config leaves runtime token empty'
    Assert-Equal 'tester' $loginModeConfig.Username 'Username parsed from .env'
    Assert-Equal 'secret-password' $loginModeConfig.Password 'Password parsed from .env'
    Assert-Equal $true $loginModeConfig.FetchRelated 'FETCH_RELATED defaults to true'

    @'
INVOICE_DIRECTION=purchase
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
OUTPUT_DIR=output
OUTPUT_XLSX=test.xlsx
'@ | Set-Content -LiteralPath $tempEmptyEnv -Encoding UTF8
    $missingAuthRejected = $false
    try { Get-HddtConfig -EnvFile $tempEmptyEnv -RepositoryRoot $root | Out-Null }
    catch { $missingAuthRejected = $true }
    Assert-Equal $true $missingAuthRejected 'Missing username and password is rejected'
}
finally {
    Remove-Item -LiteralPath $tempLoginEnv, $tempEmptyEnv -Force -ErrorAction SilentlyContinue
}

# --- Cot chuoi lien quan va style xuong dong ---
$tempWrapXlsx = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.xlsx')
try {
    $wrapSummary = [pscustomobject]@{
        Direction = 'purchase'; Source = 'query'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '123'
        Status = 2; RelatedChain = "dong 1`r`ndong 2"; RelatedInfo = 'thong tin lien quan'
    }
    $null = Export-InvoiceWorkbook -Path $tempWrapXlsx -SummaryRows @($wrapSummary) -DetailRows @() -ErrorRows @() -Overwrite
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $wrapZip = [IO.Compression.ZipFile]::OpenRead($tempWrapXlsx)
    try {
        $wrapStylesEntry = $wrapZip.GetEntry('xl/styles.xml')
        $wrapStylesReader = New-Object IO.StreamReader($wrapStylesEntry.Open())
        try { $wrapStylesText = $wrapStylesReader.ReadToEnd() } finally { $wrapStylesReader.Dispose() }
        Assert-Equal $true ($wrapStylesText -match '<cellXfs count="17">') 'Workbook contains the complete data-only style set'
        Assert-Equal $true ($wrapStylesText -match 'wrapText="1"') 'Wrap style enables text wrapping'

        $wrapSheetEntry = $wrapZip.GetEntry('xl/worksheets/sheet1.xml')
        $wrapSheetReader = New-Object IO.StreamReader($wrapSheetEntry.Open())
        try { $wrapSheetText = $wrapSheetReader.ReadToEnd() } finally { $wrapSheetReader.Dispose() }
        Assert-Equal $true ($wrapSheetText -match '<c r="BE3" s="9"') 'RelatedChain cell uses wrap style'
        Assert-Equal $true ($wrapSheetText -match '<c r="BL3" s="9"') 'RelatedInfo cell uses wrap style'
        Assert-Equal $true ($wrapSheetText -match '<c r="D3"[^>]*t="inlineStr"') 'Normal cell keeps default style'
    }
    finally { $wrapZip.Dispose() }
}
finally {
    Remove-Item -LiteralPath $tempWrapXlsx -Force -ErrorAction SilentlyContinue
}

# --- Kiem tra splat: GET/DELETE khong duoc gui body ---
$script:CapturedSplat = $null
function Invoke-WebRequest {
    param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction, $Body, $ContentType, [uri]$Proxy, [pscredential]$ProxyCredential)
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

    $proxyConfig = [pscustomobject]@{
        BaseUrl = $splatConfig.BaseUrl; Token = 'test-token'; RequestDelayMs = 0
        AdaptiveThrottle = $false; MaxRetries = 0; HttpTimeoutSeconds = 30
        ProxyUri = [uri]'http://proxy.example:8080'; ProxyUsername = 'proxy-user'; ProxyPassword = 'proxy-pass'
    }
    Invoke-GdtRequest -Config $proxyConfig -Uri ($proxyConfig.BaseUrl + '/captcha') | Out-Null
    Assert-Equal 'http://proxy.example:8080/' ([string]$script:CapturedSplat['Proxy']) 'Proxy URI is passed to Invoke-WebRequest'
    Assert-Equal 'proxy-user' $script:CapturedSplat['ProxyCredential'].UserName 'Proxy credential username is passed'
    Assert-Equal 'proxy-pass' $script:CapturedSplat['ProxyCredential'].GetNetworkCredential().Password 'Proxy credential password is passed'
}
finally {
    Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue
}

# --- Kiem tra retry 429: phai thu lai nhieu lan thay vi dung ---
# Dung WebException voi Response gia (gia lap StatusCode = 429) de
# Get-HttpStatusCode doc duoc; Start-Sleep duoc thay the de khong cho that.
Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Net;
namespace HddtTest {
    public class FakeWebResponse : System.Net.WebResponse {
        private readonly Uri _uri;
        private readonly Hashtable _headers;
        public FakeWebResponse(long statusCode) {
            _uri = new Uri("https://hoadondientu.gdt.gov.vn/api/test");
            _headers = new Hashtable();
            _headers["StatusCode"] = (int)statusCode;
        }
        public override Uri ResponseUri { get { return _uri; } }
        public Hashtable HeadersTable { get { return _headers; } }
        public int StatusCode { get { return (int)_headers["StatusCode"]; } }
    }
}
'@
$script:RateLimitHits = 0
$script:RateLimitSleptSeconds = @()
function Invoke-WebRequest {
    param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction, $Body, $ContentType)
    $script:RateLimitHits++
    if ($script:RateLimitHits -le 6) {
        $fakeResponse = New-Object HddtTest.FakeWebResponse (429)
        $exception = New-Object System.Net.WebException('Simulated HTTP 429', $null, [System.Net.WebExceptionStatus]::ProtocolError, $fakeResponse)
        $errorRecord = New-Object System.Management.Automation.ErrorRecord (
            $exception, 'Http429', [System.Management.Automation.ErrorCategory]::InvalidOperation, $Uri
        )
        throw $errorRecord
    }
    return [pscustomobject]@{ StatusCode = 200; Content = '{"ok":true}' }
}
try {
    $rlConfig = [pscustomobject]@{
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'
        Token = 'test-token'
        RequestDelayMs = 0
        AdaptiveThrottle = $false
        MaxRetries = 4
        HttpTimeoutSeconds = 30
    }
    Reset-HddtStopRequest
    $script:OriginalSleep = ${function:Start-Sleep}
    Set-Item Function:\Start-Sleep -Value { param($Seconds) $script:RateLimitSleptSeconds += [int]$Seconds }
    $result429 = Invoke-GdtRequest -Config $rlConfig -Uri ($rlConfig.BaseUrl + '/invoices/query')
    Assert-Equal 7 $script:RateLimitHits '429 retries continue until success'
    Assert-Equal '{"ok":true}' $result429 'Request succeeds after rate-limit window'
    Assert-Equal 6 @($script:RateLimitSleptSeconds).Count 'Each 429 retry waits with backoff'
}
finally {
    Set-Item Function:\Start-Sleep -Value $script:OriginalSleep
    Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue
}

# =============================================================================
# Nang cap: tai XML song song co kiem soat + HTTP profile theo trinh duyet
# =============================================================================

# --- Cau hinh moi: XML_CONCURRENCY, XML_MAX_CONCURRENCY, XML_REQUEST_INTERVAL_MS ---
$baseEnvLines = @(
    'GDT_USERNAME=tester'
    'GDT_PASSWORD=secret-password'
    'INVOICE_DIRECTION=purchase'
    'FROM_DATE=01/09/2026'
    'TO_DATE=30/09/2026'
    'OUTPUT_DIR=output'
    'OUTPUT_XLSX=test.xlsx'
)
$tempNewEnvFiles = New-Object System.Collections.Generic.List[string]
function New-TestEnvFile {
    param([string[]]$ExtraLines = @())
    $path = Join-Path ([IO.Path]::GetTempPath()) ('hddt-test-' + [guid]::NewGuid().ToString('N') + '.env')
    $lines = @($baseEnvLines) + @($ExtraLines)
    [IO.File]::WriteAllLines($path, [string[]]$lines, (New-Object Text.UTF8Encoding($false)))
    $script:tempNewEnvFiles.Add($path)
    return $path
}
function Test-TestEnvRejected {
    param([string[]]$ExtraLines = @())
    $path = New-TestEnvFile -ExtraLines $ExtraLines
    $rejected = $false
    try { Get-HddtConfig -EnvFile $path -RepositoryRoot $root | Out-Null }
    catch { $rejected = $true }
    return $rejected
}
try {
    $defaultNewConfig = Get-HddtConfig -EnvFile (New-TestEnvFile) -RepositoryRoot $root
    Assert-Equal 4 $defaultNewConfig.XmlConcurrency 'XML_CONCURRENCY defaults to 4'
    Assert-Equal 4 $defaultNewConfig.XmlMaxConcurrency 'XML_MAX_CONCURRENCY defaults to 4'
    Assert-Equal 800 $defaultNewConfig.XmlRequestIntervalMs 'XML_REQUEST_INTERVAL_MS defaults to 800'
    Assert-Equal '' $defaultNewConfig.BrowserUserAgent 'BROWSER_USER_AGENT defaults to empty'
    Assert-Equal $false $defaultNewConfig.LogHttpProfile 'LOG_HTTP_PROFILE defaults to false'

    $customNewConfig = Get-HddtConfig -EnvFile (New-TestEnvFile -ExtraLines @(
        'XML_CONCURRENCY=4'
        'XML_MAX_CONCURRENCY=6'
        'XML_REQUEST_INTERVAL_MS=250'
        'BROWSER_USER_AGENT=TestAgent/1.0'
        'LOG_HTTP_PROFILE=true'
    )) -RepositoryRoot $root
    Assert-Equal 4 $customNewConfig.XmlConcurrency 'XML_CONCURRENCY parsed from .env'
    Assert-Equal 6 $customNewConfig.XmlMaxConcurrency 'XML_MAX_CONCURRENCY parsed from .env'
    Assert-Equal 250 $customNewConfig.XmlRequestIntervalMs 'XML_REQUEST_INTERVAL_MS parsed from .env'
    Assert-Equal 'TestAgent/1.0' $customNewConfig.BrowserUserAgent 'BROWSER_USER_AGENT parsed from .env'
    Assert-Equal $true $customNewConfig.LogHttpProfile 'LOG_HTTP_PROFILE parsed from .env'

    Assert-Equal $true (Test-TestEnvRejected @('XML_CONCURRENCY=5', 'XML_MAX_CONCURRENCY=3')) 'XML_CONCURRENCY above XML_MAX_CONCURRENCY is rejected'
    Assert-Equal $true (Test-TestEnvRejected @('XML_CONCURRENCY=0')) 'XML_CONCURRENCY below 1 is rejected'
    Assert-Equal $true (Test-TestEnvRejected @('XML_CONCURRENCY=11')) 'XML_CONCURRENCY above 10 is rejected'
    Assert-Equal $true (Test-TestEnvRejected @('XML_MAX_CONCURRENCY=0')) 'XML_MAX_CONCURRENCY below 1 is rejected'
    Assert-Equal $true (Test-TestEnvRejected @('XML_REQUEST_INTERVAL_MS=60001')) 'XML_REQUEST_INTERVAL_MS range is enforced'
    Assert-Equal $true (Test-TestEnvRejected @('LOG_HTTP_PROFILE=maybe')) 'LOG_HTTP_PROFILE must be a boolean'
}
finally {
    foreach ($envPath in $tempNewEnvFiles) { Remove-Item -LiteralPath $envPath -Force -ErrorAction SilentlyContinue }
}

# --- Get-HddtConfigValue doc an toan object cau hinh thieu truong ---
$stubConfigObject = [pscustomobject]@{ Token = 'abc'; EmptyText = ''; NullText = $null }
Assert-Equal 'abc' (Get-HddtConfigValue $stubConfigObject 'Token' 'def') 'Existing config value is read'
Assert-Equal 'def' (Get-HddtConfigValue $stubConfigObject 'MissingText' 'def') 'Missing config property returns default'
Assert-Equal '' (Get-HddtConfigValue $stubConfigObject 'EmptyText' 'def') 'Empty config value is kept'
Assert-Equal 'def' (Get-HddtConfigValue $stubConfigObject 'NullText' 'def') 'Null config value returns default'
Assert-Equal 'def' (Get-HddtConfigValue $null 'Token' 'def') 'Null config object returns default'

# --- Browser profile: mot UA, client hint dong bo, khong random ---
Reset-GdtBrowserProfile
Assert-Equal $true ((Get-GdtBrowserUserAgent -Config ([pscustomobject]@{})) -like 'Mozilla/5.0*') 'Default user agent is a browser UA'
Assert-Equal $true ((Get-GdtBrowserUserAgent -Config $null) -like 'Mozilla/5.0*') 'Missing config falls back to default UA'
Assert-Equal 'ConfigAgent/1.0' (Get-GdtBrowserUserAgent -Config ([pscustomobject]@{ BrowserUserAgent = 'ConfigAgent/1.0' })) 'Config overrides the user agent'

Reset-GdtBrowserProfile
$profileA = Get-GdtBrowserProfile -Config ([pscustomobject]@{ BrowserUserAgent = 'AgentA/1.0' })
$profileB = Get-GdtBrowserProfile -Config ([pscustomobject]@{ BrowserUserAgent = 'AgentA/1.0' })
Assert-Equal $true ([object]::ReferenceEquals($profileA, $profileB)) 'Same UA reuses one profile object per run'
$profileC = Get-GdtBrowserProfile -Config ([pscustomobject]@{ BrowserUserAgent = 'AgentB/1.0' })
Assert-Equal $false ([object]::ReferenceEquals($profileA, $profileC)) 'Different UA creates a new profile'

$edgeProfile = New-GdtBrowserProfile -UserAgent 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0'
Assert-Equal $true ($edgeProfile.SecChUa -like '*"Microsoft Edge";v="126"*') 'Edge client hint matches the UA major version'
$chromeProfile = New-GdtBrowserProfile -UserAgent 'Mozilla/5.0 (Windows NT 10.0) Chrome/130.0.0.0 Safari/537.36'
Assert-Equal $true ($chromeProfile.SecChUa -like '*"Google Chrome";v="130"*') 'Chrome client hint matches the UA major version'
$genericProfile = New-GdtBrowserProfile -UserAgent 'Custom/1.0'
Assert-Equal $true ($genericProfile.SecChUa -like '*Chromium*') 'Unknown UA still produces a usable client hint'
Assert-Equal '"Windows"' $genericProfile.SecChUaPlatform 'Client hint platform is Windows'
Assert-Equal '?0' $genericProfile.SecChUaMobile 'Client hint mobile flag is 0'
Assert-Equal $true ($genericProfile.AcceptLanguage -like 'vi-VN*') 'Locale prefers Vietnamese'
Assert-Equal 'https://hoadondientu.gdt.gov.vn/' $genericProfile.RootReferer 'Root referer is the GDT home page'
Assert-Equal 'https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don' $genericProfile.LookupReferer 'Lookup referer is the tra cuu page'

# --- Suy ra request profile tu URI ---
Assert-Equal 'Captcha' (Resolve-GdtRequestProfile -Uri 'https://hoadondientu.gdt.gov.vn/api/captcha') 'Captcha profile from URI'
Assert-Equal 'Login' (Resolve-GdtRequestProfile -Uri 'https://hoadondientu.gdt.gov.vn/api/security-taxpayer/authenticate') 'Login profile from URI'
Assert-Equal 'ExportXml' (Resolve-GdtRequestProfile -Uri 'https://hoadondientu.gdt.gov.vn/api/query/invoices/export-xml?nbmst=1') 'ExportXml profile from URI'
Assert-Equal 'InvoiceRelation' (Resolve-GdtRequestProfile -Uri 'https://hoadondientu.gdt.gov.vn/api/query/invoices/relative?x=1') 'Relative profile from URI'
Assert-Equal 'InvoiceRelation' (Resolve-GdtRequestProfile -Uri 'https://hoadondientu.gdt.gov.vn/api/query/invoices/related?x=1') 'Related profile from URI'
Assert-Equal 'InvoiceQuery' (Resolve-GdtRequestProfile -Uri 'https://hoadondientu.gdt.gov.vn/api/query/invoices?size=50') 'Query profile is the default'
Assert-Equal 'InvoiceQuery' (Resolve-GdtRequestProfile -Uri 'https://hoadondientu.gdt.gov.vn/api/sco-query/invoices?size=50') 'SCO query keeps the JSON profile'
Assert-Equal 'ExportXml' (Resolve-GdtRequestProfile -Uri 'https://hoadondientu.gdt.gov.vn/api/query/invoices?x=1' -RequestProfile 'ExportXml') 'Explicit profile wins over URI inference'

# --- Header theo tung endpoint ---
Reset-GdtBrowserProfile
$headerConfig = [pscustomobject]@{ BrowserUserAgent = 'UnitTestAgent/1.0'; LogHttpProfile = $false }
$captchaHeaders = Get-GdtRequestHeaders -RequestProfile 'Captcha' -Config $headerConfig
Assert-Equal 'UnitTestAgent/1.0' $captchaHeaders['User-Agent'] 'User-Agent comes from the browser profile'
Assert-Equal $true ($captchaHeaders['Accept'] -like 'application/json*') 'Captcha profile accepts JSON'
Assert-Equal 'https://hoadondientu.gdt.gov.vn/' $captchaHeaders['Referer'] 'Captcha refers to the root page'
Assert-Equal $false ($captchaHeaders.ContainsKey('Origin')) 'Captcha sends no Origin'
Assert-Equal $false ($captchaHeaders.ContainsKey('Authorization')) 'Empty token sends no Authorization header'

$loginHeaders = Get-GdtRequestHeaders -RequestProfile 'Login' -Config $headerConfig
Assert-Equal 'https://hoadondientu.gdt.gov.vn' $loginHeaders['Origin'] 'Login sends Origin'
Assert-Equal 'https://hoadondientu.gdt.gov.vn/' $loginHeaders['Referer'] 'Login refers to the root page'

$queryHeaders = Get-GdtRequestHeaders -RequestProfile 'InvoiceQuery' -Config $headerConfig -AuthorizationToken 'tok-1'
Assert-Equal 'https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don' $queryHeaders['Referer'] 'Query refers to the tra cuu page'
Assert-Equal 'Bearer tok-1' $queryHeaders['Authorization'] 'Bearer token is attached'
Assert-Equal 'cors' $queryHeaders['Sec-Fetch-Mode'] 'Query uses cors fetch mode'
Assert-Equal $true ($queryHeaders.ContainsKey('sec-ch-ua')) 'Client hint is present'
Assert-Equal $true ($queryHeaders.ContainsKey('Request-Id') -eq $false) 'Request-Id is added per request, not per profile'

$exportHeaders = Get-GdtRequestHeaders -RequestProfile 'ExportXml' -Config $headerConfig -AuthorizationToken 'tok-1'
Assert-Equal $true ($exportHeaders['Accept'] -like '*application/zip*') 'Export profile accepts zip downloads'
Assert-Equal 'https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don' $exportHeaders['Referer'] 'Export refers to the tra cuu page'

$relationHeaders = Get-GdtRequestHeaders -RequestProfile 'InvoiceRelation' -Config $headerConfig `
    -AdditionalHeaders (Get-GdtRelationHeaders -EndpointName 'relative' -Source 'query' -Direction 'purchase')
Assert-Equal $true ($relationHeaders.ContainsKey('Action')) 'Relation request keeps the Action header'
Assert-Equal '/tra-cuu/tra-cuu-hoa-don' $relationHeaders['End-Point'] 'Relation request keeps the End-Point header'
Assert-Equal 'vi' $relationHeaders['Accept-Language'] 'Endpoint Accept-Language overrides the profile'
Assert-Equal $true ($relationHeaders['Referer'] -like 'https://hoadondientu.gdt.gov.vn*') 'Relation Referer stays on the GDT origin'
Assert-Equal 'UnitTestAgent/1.0' $relationHeaders['User-Agent'] 'Endpoint headers keep the shared user agent'

# --- Log profile: chi ghi header an toan, khong bao gio log Authorization ---
$originalProfileLog = ${function:Write-HddtLog}
$script:CapturedProfileLogLines = @()
try {
    Set-Item Function:\Write-HddtLog -Value { param($Level = 'INFO', $Message) $script:CapturedProfileLogLines += [string]$Message }
    Reset-GdtBrowserProfile
    $logEnabledConfig = [pscustomobject]@{ BrowserUserAgent = 'LogAgent/1.0'; LogHttpProfile = $true }
    $null = Get-GdtRequestHeaders -RequestProfile 'InvoiceQuery' -Config $logEnabledConfig -AuthorizationToken 'secret-token-value'
    Assert-Equal $true (@($script:CapturedProfileLogLines | Where-Object { $_ -match 'profile=InvoiceQuery' }).Count -eq 1) 'Header log emitted when LOG_HTTP_PROFILE=true'
    Assert-Equal $true (@($script:CapturedProfileLogLines | Where-Object { $_ -match 'LogAgent/1.0' }).Count -ge 1) 'User-Agent is logged for diagnosis'
    Assert-Equal 0 (@($script:CapturedProfileLogLines | Where-Object { $_ -match 'secret-token-value' }).Count) 'Authorization value never reaches the log'
    Assert-Equal 0 (@($script:CapturedProfileLogLines | Where-Object { $_ -match '(?i)authorization' }).Count) 'Authorization header name never reaches the log'
    Assert-Equal 0 (@($script:CapturedProfileLogLines | Where-Object { $_ -match '(?i)cookie' }).Count) 'Cookie never reaches the log'

    $script:CapturedProfileLogLines = @()
    $logDisabledConfig = [pscustomobject]@{ BrowserUserAgent = 'QuietAgent/1.0'; LogHttpProfile = $false }
    $null = Get-GdtRequestHeaders -RequestProfile 'InvoiceQuery' -Config $logDisabledConfig
    Assert-Equal 0 $script:CapturedProfileLogLines.Count 'No header log when LOG_HTTP_PROFILE=false'
}
finally {
    Set-Item Function:\Write-HddtLog -Value $originalProfileLog
    Reset-GdtBrowserProfile
}

# --- Khong con UA legacy / Accept-Encoding / Connection trong src ---
foreach ($srcFile in @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.ps1')) {
    $srcContent = [IO.File]::ReadAllText($srcFile.FullName)
    Assert-Equal 0 ([regex]::Matches($srcContent, 'HDDT-Windows-PowerShell').Count) ('Legacy user agent removed from ' + $srcFile.Name)
    Assert-Equal 0 ([regex]::Matches($srcContent, "Accept-Encoding").Count) ('Accept-Encoding left to the transport in ' + $srcFile.Name)
    Assert-Equal 0 ([regex]::Matches($srcContent, "'Connection'").Count) ('Connection header not forced in ' + $srcFile.Name)
}

# --- Bo dieu tiet tai XML: slot, khoang cach, cooldown, phuc hoi ---
function New-TestXmlConfig {
    param([int]$Concurrency = 2, [int]$MaxConcurrency = 3, [int]$IntervalMs = 0)
    return [pscustomobject]@{
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'
        Token = 'test-token'
        Username = 'tester'
        Password = 'secret-password'
        RequestDelayMs = 0
        AdaptiveThrottle = $false
        MaxRetries = 2
        HttpTimeoutSeconds = 30
        ProgressEvery = 1
        LogLevel = 'info'
        XmlConcurrency = $Concurrency
        XmlMaxConcurrency = $MaxConcurrency
        XmlRequestIntervalMs = $IntervalMs
        BrowserUserAgent = ''
        LogHttpProfile = $false
        RedownloadXml = $false
        XmlDirectory = ([IO.Path]::GetTempPath())
    }
}

$throttleConfig = New-TestXmlConfig -Concurrency 2 -MaxConcurrency 3 -IntervalMs 0
$originalThrottleLog = ${function:Write-HddtLog}
$script:CapturedThrottleLogs = @()
try {
    Set-Item Function:\Write-HddtLog -Value { param($Level = 'INFO', $Message) $script:CapturedThrottleLogs += [string]$Message }
    Reset-HddtStopRequest
    $null = New-HddtXmlSharedState -Config $throttleConfig

    $snapshot = Get-GdtXmlThrottleSnapshot
    Assert-Equal 2 $snapshot.BaseConcurrency 'Base concurrency from XML_CONCURRENCY'
    Assert-Equal 3 $snapshot.MaxConcurrency 'Max concurrency from XML_MAX_CONCURRENCY'
    Assert-Equal 2 $snapshot.CurrentConcurrency 'Initial concurrency equals XML_CONCURRENCY'
    Assert-Equal 0 $snapshot.ActiveCount 'No slot held initially'
    Assert-Equal 0 $snapshot.PeakConcurrency 'Peak concurrency starts at zero'

    # Giu dung so ket noi cho phep; phia sau bi chan ngay khi khong choi.
    Assert-Equal $true (Enter-GdtXmlRequestSlot -Config $throttleConfig) 'First slot acquired'
    Assert-Equal $true (Enter-GdtXmlRequestSlot -Config $throttleConfig) 'Second slot acquired up to XML_CONCURRENCY'
    Assert-Equal $false (Enter-GdtXmlRequestSlot -Config $throttleConfig -TryOnly) 'Third slot refused at the initial cap'
    Assert-Equal 2 (Get-GdtXmlThrottleSnapshot).ActiveCount 'Two slots are held'
    Exit-GdtXmlRequestSlot -ResponseTimeMs 120
    Exit-GdtXmlRequestSlot
    $snapshot = Get-GdtXmlThrottleSnapshot
    Assert-Equal 0 $snapshot.ActiveCount 'Both slots released'
    Assert-Equal 2 $snapshot.PeakConcurrency 'Peak concurrency recorded'
    Assert-Equal 1 $snapshot.CompletedCount 'Only a successful response time is counted'
    Assert-Equal 120 $snapshot.TotalResponseTimeMs 'Response time total accumulates'

    # Phuc hoi day concurrency len muc toi da; khong vuot qua XML_MAX_CONCURRENCY.
    Set-HddtSharedValue -Key 'XmlCurrentConcurrency' -Value 3
    Enter-GdtXmlRequestSlot -Config $throttleConfig | Out-Null
    Enter-GdtXmlRequestSlot -Config $throttleConfig | Out-Null
    Enter-GdtXmlRequestSlot -Config $throttleConfig | Out-Null
    Assert-Equal $false (Enter-GdtXmlRequestSlot -Config $throttleConfig -TryOnly) 'Fourth slot refused at XML_MAX_CONCURRENCY'
    Exit-GdtXmlRequestSlot; Exit-GdtXmlRequestSlot; Exit-GdtXmlRequestSlot
    Assert-Equal 0 (Get-GdtXmlThrottleSnapshot).ActiveCount 'All cap-test slots released'

    # Gian cach giua hai request XML duoc ton trong.
    Set-HddtSharedValue -Key 'XmlCurrentIntervalMs' -Value 150
    Set-HddtSharedValue -Key 'XmlLastRequestStartUtc' -Value ([datetime]::MinValue)
    $intervalStart = [datetime]::UtcNow
    $null = Enter-GdtXmlRequestSlot -Config $throttleConfig
    Exit-GdtXmlRequestSlot
    $null = Enter-GdtXmlRequestSlot -Config $throttleConfig
    Exit-GdtXmlRequestSlot
    $intervalElapsedMs = (([datetime]::UtcNow) - $intervalStart).TotalMilliseconds
    Assert-Equal $true ($intervalElapsedMs -ge 130) ('XML interval spaces request starts (measured {0} ms)' -f [int]$intervalElapsedMs)

    # HTTP 429: cooldown theo Retry-After, giam mot ket noi, tang gian cach 1.5x.
    Set-HddtSharedValue -Key 'XmlCurrentIntervalMs' -Value 800
    Set-HddtSharedValue -Key 'XmlCurrentConcurrency' -Value 2
    $script:CapturedThrottleLogs = @()
    $cooldownStart = [datetime]::UtcNow
    Register-GdtXmlRateLimit -RetryAfterSeconds 45
    $snapshot = Get-GdtXmlThrottleSnapshot
    Assert-Equal 1 $snapshot.RateLimitCount 'Rate limit counted once'
    Assert-Equal 1 $snapshot.CurrentConcurrency 'Concurrency drops by one on 429'
    Assert-Equal 1200 $snapshot.CurrentIntervalMs 'Interval grows by 1.5x on 429'
    $lastRateLimitUtc = [datetime](Get-HddtSharedValue -Key 'XmlLastRateLimitUtc')
    Assert-Equal $true (($lastRateLimitUtc - $cooldownStart).TotalSeconds -ge 0 -and ($lastRateLimitUtc - $cooldownStart).TotalSeconds -le 5) '429 records the last rate-limit time for recovery'
    $cooldownUntil = [datetime](Get-HddtSharedValue -Key 'XmlGlobalCooldownUntilUtc')
    Assert-Equal $true (($cooldownUntil - $cooldownStart).TotalSeconds -ge 44) 'Cooldown follows Retry-After'
    Assert-Equal $true (($cooldownUntil - $cooldownStart).TotalSeconds -le 47) 'No jitter when Retry-After is present'
    Assert-Equal $false (Enter-GdtXmlRequestSlot -Config $throttleConfig -TryOnly) 'Slot refused during cooldown'
    Assert-Equal 1 (@($script:CapturedThrottleLogs | Where-Object { $_ -match '\[XML THROTTLE\] HTTP 429 \| cooldown 45s \| concurrency 2→1 \| interval 800→1200 ms' })).Count 'Throttle log reports cooldown, concurrency and interval'

    # Hai worker nhan 429 gan nhau: phan hoi co Retry-After ngan hon khong
    # duoc rut ngan cooldown dai hon da duoc ghi truoc do.
    Register-GdtXmlRateLimit -RetryAfterSeconds 120
    $longCooldownUntil = [datetime](Get-HddtSharedValue -Key 'XmlGlobalCooldownUntilUtc')
    Register-GdtXmlRateLimit -RetryAfterSeconds 30
    $mergedCooldownUntil = [datetime](Get-HddtSharedValue -Key 'XmlGlobalCooldownUntilUtc')
    Assert-Equal $true ($mergedCooldownUntil -ge $longCooldownUntil) 'Concurrent 429 cooldown keeps the latest expiry'
    Assert-Equal $false (Enter-GdtXmlRequestSlot -Config $throttleConfig -TryOnly) 'Longer cooldown remains enforced after a shorter Retry-After'

    # Khong co Retry-After va fallback = 0: khong cho phep cho, van giam ket noi.
    Set-HddtSharedValue -Key 'CooldownFallbackSeconds' -Value 0
    Set-HddtSharedValue -Key 'XmlCurrentIntervalMs' -Value 0
    Set-HddtSharedValue -Key 'XmlGlobalCooldownUntilUtc' -Value ([datetime]::MinValue)
    Register-GdtXmlRateLimit -RetryAfterSeconds 0
    $cooldownUntil = [datetime](Get-HddtSharedValue -Key 'XmlGlobalCooldownUntilUtc')
    Assert-Equal $true ($cooldownUntil -le ([datetime]::UtcNow).AddMilliseconds(50)) 'Zero fallback means no cooldown wait'
    Assert-Equal $true (Enter-GdtXmlRequestSlot -Config $throttleConfig -TryOnly) 'Request retries immediately with zero cooldown'
    Exit-GdtXmlRequestSlot
    Assert-Equal 1 (Get-GdtXmlThrottleSnapshot).CurrentConcurrency 'Concurrency decreased on the second 429'
    Register-GdtXmlRateLimit -RetryAfterSeconds 0
    Assert-Equal 1 (Get-GdtXmlThrottleSnapshot).CurrentConcurrency 'Concurrency never drops below one'

    # Phuc hoi theo thoi gian: im 429 du lau thi chia doi gian cach ve muc cau
    # hinh truoc, sau do moi tang lai ket noi. Khong doi chuoi request thanh
    # cong lien tiep nen khong bi ket o 1 ket noi sau mot dot rate-limit.
    Set-HddtSharedValue -Key 'XmlBaseIntervalMs' -Value 800
    Set-HddtSharedValue -Key 'XmlCurrentIntervalMs' -Value 4000
    Set-HddtSharedValue -Key 'XmlCurrentConcurrency' -Value 1
    Set-HddtSharedValue -Key 'XmlLastRateLimitUtc' -Value ([datetime]::UtcNow)
    Register-GdtXmlSuccess
    Assert-Equal 4000 (Get-GdtXmlThrottleSnapshot).CurrentIntervalMs 'No recovery before the quiet window'

    Set-HddtSharedValue -Key 'XmlLastRateLimitUtc' -Value ([datetime]::UtcNow).AddSeconds(-30)
    $script:CapturedThrottleLogs = @()
    Register-GdtXmlSuccess
    $snapshot = Get-GdtXmlThrottleSnapshot
    Assert-Equal 2000 $snapshot.CurrentIntervalMs 'Interval is halved after a quiet window'
    Assert-Equal 1 (@($script:CapturedThrottleLogs | Where-Object { $_ -match 'interval 4000→2000 ms' })).Count 'Interval recovery logged'

    Set-HddtSharedValue -Key 'XmlCurrentIntervalMs' -Value 800
    Set-HddtSharedValue -Key 'XmlLastRateLimitUtc' -Value ([datetime]::UtcNow).AddSeconds(-30)
    $script:CapturedThrottleLogs = @()
    Register-GdtXmlSuccess
    $snapshot = Get-GdtXmlThrottleSnapshot
    Assert-Equal 2 $snapshot.CurrentConcurrency 'Concurrency raised by one after the interval is back to base'
    Assert-Equal 1 (@($script:CapturedThrottleLogs | Where-Object { $_ -match 'concurrency 1→2' })).Count 'Concurrency recovery logged'

    # Hoi phuc hoan toan: ngung theo doi de khong kiem tra/log lap moi request.
    Set-HddtSharedValue -Key 'XmlCurrentConcurrency' -Value 3
    Set-HddtSharedValue -Key 'XmlLastRateLimitUtc' -Value ([datetime]::UtcNow).AddSeconds(-30)
    $script:CapturedThrottleLogs = @()
    Register-GdtXmlSuccess
    Assert-Equal 3 (Get-GdtXmlThrottleSnapshot).CurrentConcurrency 'Concurrency never exceeds XML_MAX_CONCURRENCY'
    Assert-Equal ([datetime]::MinValue) (Get-HddtSharedValue -Key 'XmlLastRateLimitUtc') 'Fully recovered throttle stops tracking the last 429'
    Assert-Equal 0 $script:CapturedThrottleLogs.Count 'No log while fully recovered'

    # Dung an toan: khong gui them request nao sau yeu cau dung.
    Set-HddtSharedValue -Key 'StopRequested' -Value $true
    $stopRejected = $false
    try { $null = Enter-GdtXmlRequestSlot -Config $throttleConfig -TryOnly }
    catch { $stopRejected = $true }
    Assert-Equal $true $stopRejected 'Slot entry refuses new requests after stop'
    Set-HddtSharedValue -Key 'StopRequested' -Value $false
}
finally {
    Set-Item Function:\Write-HddtLog -Value $originalThrottleLog
    Set-HddtSharedState -Shared $null
    Reset-HddtStopRequest
}

# --- Co dung trang thai dung chung cho worker thay cờ dung khong ---
$stopConfig = New-TestXmlConfig
try {
    $null = New-HddtXmlSharedState -Config $stopConfig
    Assert-Equal $false (Test-HddtStopRequested) 'Stop flag is clear initially'
    Set-HddtStopRequest
    Assert-Equal $true (Get-HddtSharedValue -Key 'StopRequested') 'Stop flag is visible to worker runspaces'
    Assert-Equal $true (Test-HddtStopRequested) 'Main thread sees the stop flag'
    Reset-HddtStopRequest
    Assert-Equal $false (Get-HddtSharedValue -Key 'StopRequested') 'Reset clears the shared stop flag'
}
finally {
    Set-HddtSharedState -Shared $null
    Reset-HddtStopRequest
}

# --- Get-GdtXmlDownloadResult voi Save-GdtInvoiceXml gia ---
$originalSaveXml = ${function:Save-GdtInvoiceXml}
$script:SaveXmlMockCalls = 0
try {
    Set-Item Function:\Save-GdtInvoiceXml -Value {
        param($Config, $Invoice)
        $script:SaveXmlMockCalls++
        if ($Invoice.InvoiceNumber -eq '2') { throw 'mo phong loi tai XML' }
        return @((Join-Path $Config.XmlDirectory ('fake-{0}.xml' -f $Invoice.InvoiceNumber)))
    }
    $downloadConfig = New-TestXmlConfig
    $okInvoice = [pscustomobject]@{ Direction = 'purchase'; Source = 'query'; SellerTaxCode = '0101111111'; InvoiceTemplate = '1'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '1' }
    $okResult = Get-GdtXmlDownloadResult -Config $downloadConfig -Invoice $okInvoice -PipelineIndex 4 -WorkerIndex 2 -WorkerCount 3
    Assert-Equal $true $okResult.Success 'Successful download is reported as success'
    Assert-Equal 4 $okResult.PipelineIndex 'Pipeline index is preserved'
    Assert-Equal 2 $okResult.WorkerIndex 'Result records which worker handled the invoice'
    Assert-Equal 3 $okResult.WorkerCount 'Result records the worker pool size'
    Assert-Equal 1 @($okResult.XmlFiles).Count 'XML file list is returned'
    Assert-Equal 'purchase/C26TABC/1' $okResult.Label 'Label is built for progress logs'
    Assert-Equal '' $okResult.ErrorMessage 'Success has no error message'
    Assert-Equal 1 $script:SaveXmlMockCalls 'Download helper calls Save-GdtInvoiceXml once'

    $failInvoice = [pscustomobject]@{ Direction = 'purchase'; Source = 'query'; SellerTaxCode = '0101111111'; InvoiceTemplate = '1'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '2' }
    $failResult = Get-GdtXmlDownloadResult -Config $downloadConfig -Invoice $failInvoice -PipelineIndex 5
    Assert-Equal $false $failResult.Success 'Failed download is reported without throwing'
    Assert-Equal 'Tải XML' $failResult.Stage 'Failure stage is the download stage'
    Assert-Equal 'mo phong loi tai XML' $failResult.ErrorMessage 'Error message is captured for the workbook'
    Assert-Equal 5 $failResult.PipelineIndex 'Failed result keeps its index'
}
finally { Set-Item Function:\Save-GdtInvoiceXml -Value $originalSaveXml }

# --- REDOWNLOAD_XML: dung lai XML da co hoac tai lai ---
$originalRedownloadRequest = ${function:Invoke-GdtRequest}
$script:RedownloadRequestCalls = 0
$script:RedownloadZipBytes = $null
$tempRedownloadRoot = Join-Path ([IO.Path]::GetTempPath()) ('hddt-redownload-' + [guid]::NewGuid().ToString('N'))
try {
    $redownloadZipStream = New-Object IO.MemoryStream
    $redownloadArchive = New-Object IO.Compression.ZipArchive($redownloadZipStream, [IO.Compression.ZipArchiveMode]::Create, $true)
    $redownloadEntry = $redownloadArchive.CreateEntry('nested/invoice.xml')
    $redownloadWriter = New-Object IO.StreamWriter($redownloadEntry.Open())
    try { $redownloadWriter.Write('<HDon />') } finally { $redownloadWriter.Dispose() }
    $redownloadArchive.Dispose()
    $script:RedownloadZipBytes = $redownloadZipStream.ToArray()
    $redownloadZipStream.Dispose()

    $purchaseDownloadDir = Join-Path $tempRedownloadRoot 'purchase'
    New-Item -ItemType Directory -Path $purchaseDownloadDir -Force | Out-Null
    $existingXmlPath = Join-Path $purchaseDownloadDir 'purchase_query_0101111111_1_C26TABC_123.xml'
    [IO.File]::WriteAllText($existingXmlPath, '<HDon />')

    Set-Item Function:\Invoke-GdtRequest -Value {
        param($Config, $Uri, [switch]$AsBytes, [string]$RequestProfile)
        $script:RedownloadRequestCalls++
        return $script:RedownloadZipBytes
    }
    $redownloadConfig = [pscustomobject]@{
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'
        RedownloadXml = $false
        XmlDirectory = $tempRedownloadRoot
    }
    $redownloadInvoice = [pscustomobject]@{ Direction = 'purchase'; Source = 'query'; SellerTaxCode = '0101111111'; InvoiceTemplate = '1'; InvoiceSeries = 'C26TABC'; InvoiceNumber = '123' }

    $reusedFiles = @(Save-GdtInvoiceXml -Config $redownloadConfig -Invoice $redownloadInvoice)
    Assert-Equal 0 $script:RedownloadRequestCalls 'Existing XML is reused without any request'
    Assert-Equal 1 $reusedFiles.Count 'Reuse returns the existing file'
    Assert-Equal $existingXmlPath $reusedFiles[0] 'Reuse returns the existing file path'

    $redownloadConfig.RedownloadXml = $true
    $freshFiles = @(Save-GdtInvoiceXml -Config $redownloadConfig -Invoice $redownloadInvoice)
    Assert-Equal 1 $script:RedownloadRequestCalls 'REDOWNLOAD_XML=true downloads again'
    Assert-Equal 1 $freshFiles.Count 'Redownload writes the XML again'
    Assert-Equal $true (Test-Path -LiteralPath $freshFiles[0]) 'Redownloaded XML file exists'
}
finally {
    Set-Item Function:\Invoke-GdtRequest -Value $originalRedownloadRequest
    Remove-Item -LiteralPath $tempRedownloadRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# --- Token single-flight: mot lan CAPTCHA/dang nhap cho ca phien ---
$originalTokenRequest = ${function:Request-GdtSessionToken}
$script:TokenRefreshMockCalls = 0
try {
    Set-Item Function:\Request-GdtSessionToken -Value {
        param($Config)
        $script:TokenRefreshMockCalls++
        Start-Sleep -Milliseconds 50
        return ('token-{0}' -f $script:TokenRefreshMockCalls)
    }
    $tokenConfig = New-TestXmlConfig
    $tokenConfig.Token = 'old-token'
    $null = New-HddtXmlSharedState -Config $tokenConfig
    Assert-Equal 'old-token' (Get-GdtAuthToken -Config $tokenConfig) 'Shared token is the source of truth'

    $refreshedToken = Invoke-GdtTokenRefresh -Config $tokenConfig -FailedToken 'old-token'
    Assert-Equal 'token-1' $refreshedToken 'Refresh returns the new token'
    Assert-Equal 1 $script:TokenRefreshMockCalls 'Exactly one login for one refresh'
    Assert-Equal 1 (Get-GdtXmlThrottleSnapshot).AuthRefreshCount 'Refresh is counted in shared metrics'
    Set-GdtAuthToken -Config $tokenConfig -Token $refreshedToken
    Assert-Equal 'token-1' $tokenConfig.Token 'Set-GdtAuthToken writes the token back to config'

    $staleToken = Invoke-GdtTokenRefresh -Config $tokenConfig -FailedToken 'old-token'
    Assert-Equal 'token-1' $staleToken 'A stale token sees the refreshed one'
    Assert-Equal 1 $script:TokenRefreshMockCalls 'No extra login when the token is already refreshed'

    $sameFailedToken = Invoke-GdtTokenRefresh -Config $tokenConfig -FailedToken 'token-1'
    Assert-Equal 'token-2' $sameFailedToken 'The current failed token still triggers a refresh'
    Assert-Equal 2 $script:TokenRefreshMockCalls 'Second refresh logged in once'
}
finally {
    Set-Item Function:\Request-GdtSessionToken -Value $originalTokenRequest
    Set-HddtSharedState -Shared $null
}

# --- Single-flight that su: 3 runspace cung goi, chi mot lan dang nhap ---
$singleFlightShared = $null
try {
    $singleFlightConfig = New-TestXmlConfig
    $singleFlightShared = New-HddtXmlSharedState -Config $singleFlightConfig
    Set-HddtSharedValue -Key 'Token' -Value 'stale-token'
    Set-HddtSharedValue -Key 'AuthRefreshCount' -Value 0

    $singleFlightScript = @'
param($Root, $Shared)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
. (Join-Path $Root 'src/Logging.ps1')
. (Join-Path $Root 'src/Config.ps1')
. (Join-Path $Root 'src/Http.ps1')
. (Join-Path $Root 'src/Login.ps1')
Set-HddtSharedState -Shared $Shared
function Request-GdtSessionToken {
    param($Config)
    Start-Sleep -Milliseconds 600
    return 'single-flight-token'
}
$flightConfig = [pscustomobject]@{
    BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'
    Token = 'stale-token'
    Username = 'tester'
    Password = 'secret-password'
    RequestDelayMs = 0
    AdaptiveThrottle = $false
    MaxRetries = 0
    HttpTimeoutSeconds = 30
}
Invoke-GdtTokenRefresh -Config $flightConfig -FailedToken 'stale-token'
'@
    $flightWorkers = @()
    for ($flightIndex = 0; $flightIndex -lt 3; $flightIndex++) {
        $flightRunspace = [runspacefactory]::CreateRunspace()
        $flightRunspace.Open()
        $flightPowerShell = [powershell]::Create()
        $flightPowerShell.Runspace = $flightRunspace
        $null = $flightPowerShell.AddScript($singleFlightScript).AddArgument($root).AddArgument($singleFlightShared)
        $flightWorkers += [pscustomobject]@{
            PowerShell = $flightPowerShell
            Runspace = $flightRunspace
            Handle = $flightPowerShell.BeginInvoke()
        }
    }
    $flightResults = @()
    foreach ($flightWorker in $flightWorkers) {
        $flightResults += @($flightWorker.PowerShell.EndInvoke($flightWorker.Handle) | ForEach-Object { [string]$_ })
        try { $flightWorker.PowerShell.Dispose() } catch { }
        try { $flightWorker.Runspace.Close(); $flightWorker.Runspace.Dispose() } catch { }
    }
    Assert-Equal 3 $flightResults.Count 'Every worker receives a refreshed token'
    Assert-Equal 0 (@($flightResults | Where-Object { $_ -ne 'single-flight-token' }).Count) 'All workers see the same refreshed token'
    Assert-Equal 1 (Get-HddtSharedValue -Key 'AuthRefreshCount') 'Only one worker performs the login'
    Assert-Equal 'single-flight-token' (Get-HddtSharedValue -Key 'Token') 'Shared token updated once'
}
finally {
    if ($null -ne $singleFlightShared) { Set-HddtSharedState -Shared $null }
    Reset-HddtStopRequest
}

# --- Pipeline tai XML that su: round-robin, ket qua, log forward, dung sach ---
$stubRoot = Join-Path ([IO.Path]::GetTempPath()) ('hddt-pipeline-' + [guid]::NewGuid().ToString('N'))
$pipelineShared = $null
try {
    New-Item -ItemType Directory -Path (Join-Path $stubRoot 'src') -Force | Out-Null
    $stubEncoding = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText((Join-Path $stubRoot 'src/Logging.ps1'), @'
Set-StrictMode -Version 2.0
$script:HddtSharedState = $null
$script:HddtLogForwardToShared = $false
$script:HddtLogLevel = 'INFO'
function Set-HddtSharedState { param($Shared) $script:HddtSharedState = $Shared }
function Get-HddtSharedState { return $script:HddtSharedState }
function Get-HddtSharedValue {
    param([string]$Key, $Default = $null)
    $shared = $script:HddtSharedState
    if ($null -eq $shared) { return $Default }
    if (-not $shared.ContainsKey($Key)) { return $Default }
    return $shared[$Key]
}
function Set-HddtSharedValue {
    param([string]$Key, $Value)
    $shared = $script:HddtSharedState
    if ($null -eq $shared) { return }
    [Threading.Monitor]::Enter($shared.SyncRoot)
    try { $shared[$Key] = $Value }
    finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
}
function Write-HddtLog {
    param($Level = 'INFO', $Message)
    if ($script:HddtLogForwardToShared -and $null -ne $script:HddtSharedState) {
        $shared = $script:HddtSharedState
        [Threading.Monitor]::Enter($shared.SyncRoot)
        try { $shared.LogQueue.Add([pscustomobject]@{ Level = ([string]$Level).ToUpperInvariant(); Message = [string]$Message }) }
        finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
    }
}
'@, $stubEncoding)
    [IO.File]::WriteAllText((Join-Path $stubRoot 'src/Config.ps1'), @'
Set-StrictMode -Version 2.0
function Get-HddtConfigValue {
    param($Config, [string]$Name, $Default = $null)
    if ($null -eq $Config) { return $Default }
    $property = $Config.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}
'@, $stubEncoding)
    [IO.File]::WriteAllText((Join-Path $stubRoot 'src/Http.ps1'), @'
Set-StrictMode -Version 2.0
$script:HddtStopRequested = $false
$script:LastGdtRequestUri = ''
$script:LastGdtStatusCode = 0
$script:LastGdtRequestAttempts = 0
$script:LastGdtRetryAfterSeconds = 0
function Test-HddtStopRequested {
    if ([bool]$script:HddtStopRequested) { return $true }
    return [bool](Get-HddtSharedValue -Key 'StopRequested' -Default $false)
}
function Set-HddtStopRequest {
    $script:HddtStopRequested = $true
    Set-HddtSharedValue -Key 'StopRequested' -Value $true
}
function Reset-HddtStopRequest {
    $script:HddtStopRequested = $false
    Set-HddtSharedValue -Key 'StopRequested' -Value $false
}
function Get-GdtLastRequestUri { return [string]$script:LastGdtRequestUri }
function Get-GdtLastStatusCode { return [int]$script:LastGdtStatusCode }
function Get-GdtLastRequestAttempts { return [int]$script:LastGdtRequestAttempts }
function Get-GdtLastRetryAfterSeconds { return [int]$script:LastGdtRetryAfterSeconds }
function Get-GdtAuthToken { param($Config) return [string](Get-HddtSharedValue -Key 'Token' -Default '') }
function Set-GdtAuthToken { param($Config, $Token) Set-HddtSharedValue -Key 'Token' -Value $Token }
'@, $stubEncoding)
    [IO.File]::WriteAllText((Join-Path $stubRoot 'src/InvoiceApi.ps1'), @'
Set-StrictMode -Version 2.0
function Get-InvoiceLabel {
    param($Invoice)
    return '{0}/{1}/{2}' -f $Invoice.Direction, $Invoice.InvoiceSeries, $Invoice.InvoiceNumber
}
function Save-GdtInvoiceXml {
    param($Config, $Invoice)
    Start-Sleep -Milliseconds 20
    if ([int]$Invoice.InvoiceNumber % 2 -eq 0) { throw ('mo phong loi tai {0}' -f $Invoice.InvoiceNumber) }
    return @((Join-Path $Config.XmlDirectory ('fake-{0}.xml' -f $Invoice.InvoiceNumber)))
}
'@, $stubEncoding)
    [IO.File]::WriteAllText((Join-Path $stubRoot 'src/Login.ps1'), @'
Set-StrictMode -Version 2.0
function Request-GdtSessionToken { param($Config) return 'stub-token' }
'@, $stubEncoding)
    [IO.File]::WriteAllText((Join-Path $stubRoot 'src/BrowserProfile.ps1'), "Set-StrictMode -Version 2.0`n", $stubEncoding)
    Copy-Item -LiteralPath (Join-Path $root 'src/XmlScheduler.ps1') -Destination (Join-Path $stubRoot 'src/XmlScheduler.ps1') -Force

    $pipelineConfig = New-TestXmlConfig -Concurrency 2 -MaxConcurrency 2 -IntervalMs 0
    $pipelineShared = New-HddtXmlSharedState -Config $pipelineConfig
    $pipelineInvoices = @(1..6 | ForEach-Object {
        [pscustomobject]@{
            Direction = 'purchase'
            Source = 'query'
            SellerTaxCode = '0123456789'
            InvoiceTemplate = '1'
            InvoiceSeries = 'C26TABC'
            InvoiceNumber = ([string]$_)
        }
    })

    $originalPipelineLog = ${function:Write-HddtLog}
    $script:CapturedPipelineLogs = @()
    $pipeline = $null
    $pipelineResults = New-Object System.Collections.Generic.List[object]
    try {
        # Bat log truoc khi mo pipeline de bat ca dong thong bao so worker.
        Set-Item Function:\Write-HddtLog -Value { param($Level = 'INFO', $Message) $script:CapturedPipelineLogs += [string]$Message }
        $pipeline = Start-HddtXmlPipeline -Config $pipelineConfig -Invoices $pipelineInvoices -Root $stubRoot
        $pipelineDeadline = [datetime]::UtcNow.AddSeconds(30)
        while (-not (Test-HddtXmlPipelineCompleted -Pipeline $pipeline) -and [datetime]::UtcNow -lt $pipelineDeadline) {
            foreach ($pipelineResult in @(Receive-HddtXmlPipeline -Pipeline $pipeline)) { $pipelineResults.Add($pipelineResult) }
            Start-Sleep -Milliseconds 50
        }
        foreach ($pipelineResult in @(Receive-HddtXmlPipeline -Pipeline $pipeline)) { $pipelineResults.Add($pipelineResult) }
    }
    finally { Set-Item Function:\Write-HddtLog -Value $originalPipelineLog }

    Assert-Equal 6 $pipelineResults.Count 'Every invoice produces exactly one result'
    $sortedIndexes = @($pipelineResults | ForEach-Object { $_.PipelineIndex } | Sort-Object)
    Assert-Equal '0 1 2 3 4 5' ($sortedIndexes -join ' ') 'All pipeline indexes are covered once'
    $successfulResults = @($pipelineResults | Where-Object { $_.Success })
    $failedResults = @($pipelineResults | Where-Object { -not $_.Success })
    Assert-Equal 3 $successfulResults.Count 'Even invoices succeed'
    Assert-Equal 3 $failedResults.Count 'Odd invoices fail'
    foreach ($pipelineResult in $pipelineResults) {
        $expectedNumber = [string](($pipelineResult.PipelineIndex + 1))
        Assert-Equal ('purchase/C26TABC/' + $expectedNumber) $pipelineResult.Label 'Result maps back to its invoice'
        if ($pipelineResult.Success) {
            Assert-Equal 1 @($pipelineResult.XmlFiles).Count 'Successful result carries its XML file'
        }
        else {
            Assert-Equal ('mo phong loi tai ' + $expectedNumber) $pipelineResult.ErrorMessage 'Failed result carries the error text'
            Assert-Equal 'Tải XML' $pipelineResult.Stage 'Failed result keeps the download stage'
        }
    }
    Assert-Equal $true (@($script:CapturedPipelineLogs | Where-Object { $_ -match 'Bắt đầu hóa đơn' }).Count -ge 6) 'Worker logs are forwarded to the main thread'
    Assert-Equal $true (@($script:CapturedPipelineLogs | Where-Object { $_ -match '\[TẢI XML\] Khởi động 2 worker' }).Count -ge 1) 'Pipeline log states how many workers start'
    Assert-Equal $true (@($script:CapturedPipelineLogs | Where-Object { $_ -match 'Worker \d+/2 bắt đầu làm việc' }).Count -ge 2) 'Every worker logs that it starts working'
    Assert-Equal $true (@($script:CapturedPipelineLogs | Where-Object { $_ -match 'Worker \d+/2 kết thúc' }).Count -ge 2) 'Every worker logs that it finishes'
    Assert-Equal 2 (Get-HddtSharedValue -Key 'WorkerCount') 'Shared state exposes the worker count for progress logs'
    Assert-Equal $true (@($pipelineResults | Where-Object { $_.WorkerIndex -lt 1 -or $_.WorkerIndex -gt 2 }).Count -eq 0) 'Result worker index stays inside the worker pool'
    Assert-Equal 2 (Complete-HddtSharedState -Config $pipelineConfig).WorkerCount 'Metrics report the worker count'

    # Pipeline da xong: stop khong duoc danh dau ngu dung (khong mat cac buoc sau).
    $stopResults = @(Stop-HddtXmlPipeline -Pipeline $pipeline)
    Assert-Equal 0 $stopResults.Count 'Completed pipeline has no leftover results'
    Assert-Equal $false (Get-HddtSharedValue -Key 'StopRequested') 'Stopping a completed pipeline does not raise the stop flag'
    $pipelineMetrics = Complete-HddtSharedState -Config $pipelineConfig
    Assert-Equal $true ($null -ne $pipelineMetrics) 'Shared state closes with metrics'
    Assert-Equal 0 $pipelineMetrics.AuthRefreshCount 'No token refresh happened in the pipeline test'
}
finally {
    Set-HddtSharedState -Shared $null
    Reset-HddtStopRequest
    Remove-Item -LiteralPath $stubRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# --- HTTP 429 voi request XML: khong ngu backoff local, dieu tiep chung ---
$originalXml429Log = ${function:Write-HddtLog}
$originalXml429Sleep = ${function:Start-Sleep}
$originalXml429WebRequest = ${function:Invoke-WebRequest}
$script:Xml429RequestHits = 0
$script:Xml429SecondsSlept = 0
$script:Xml429MillisSlept = 0
$script:Xml429LogLines = @()
try {
    Set-Item Function:\Write-HddtLog -Value { param($Level = 'INFO', $Message) $script:Xml429LogLines += [string]$Message }
    Set-Item Function:\Start-Sleep -Value {
        param([int]$Seconds, [int]$Milliseconds)
        if ($PSBoundParameters.ContainsKey('Seconds')) { $script:Xml429SecondsSlept++ }
        if ($PSBoundParameters.ContainsKey('Milliseconds')) { $script:Xml429MillisSlept++ }
    }
    Set-Item Function:\Invoke-WebRequest -Value {
        param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction, $Body, $ContentType)
        $script:Xml429RequestHits++
        if ($script:Xml429RequestHits -le 2) {
            $fakeResponse = New-Object HddtTest.FakeWebResponse (429)
            $exception = New-Object System.Net.WebException('Simulated HTTP 429', $null, [System.Net.WebExceptionStatus]::ProtocolError, $fakeResponse)
            $errorRecord = New-Object System.Management.Automation.ErrorRecord($exception, 'Http429', [System.Management.Automation.ErrorCategory]::InvalidOperation, $Uri)
            throw $errorRecord
        }
        return [pscustomobject]@{
            StatusCode = 200
            Content = '{"ok":true}'
            RawContentStream = (New-Object IO.MemoryStream (,[byte[]](75, 76)))
        }
    }

    $xml429Config = New-TestXmlConfig -Concurrency 2 -MaxConcurrency 3 -IntervalMs 0
    $null = New-HddtXmlSharedState -Config $xml429Config
    Set-HddtSharedValue -Key 'CooldownFallbackSeconds' -Value 0
    Reset-HddtStopRequest

    $xml429Bytes = [byte[]](Invoke-GdtRequest -Config $xml429Config -Uri ($xml429Config.BaseUrl + '/query/invoices/export-xml?nbmst=0101111111') -AsBytes)
    Assert-Equal 3 $script:Xml429RequestHits 'XML request retries after two 429 responses'
    Assert-Equal 2 $xml429Bytes.Length 'XML request succeeds after the rate-limit window'
    Assert-Equal 0 $script:Xml429SecondsSlept 'XML 429 does not sleep a local backoff'
    Assert-Equal 0 $script:Xml429MillisSlept 'Scheduler waits nothing with zero cooldown and interval'
    $snapshot = Get-GdtXmlThrottleSnapshot
    Assert-Equal 2 $snapshot.RateLimitCount 'Both 429 responses register an XML rate limit'
    Assert-Equal 1 $snapshot.CurrentConcurrency 'Concurrency reduced and floored at one'
    Assert-Equal 1 $snapshot.CompletedCount 'Completed request counted once'
    Assert-Equal 2 (@($script:Xml429LogLines | Where-Object { $_ -match '\[XML THROTTLE\] HTTP 429' })).Count 'One throttle log per 429 response'
    Assert-Equal 0 (@($script:Xml429LogLines | Where-Object { $_ -match '(?i)secret-password' })).Count 'Password never reaches the log'
}
finally {
    Set-Item Function:\Write-HddtLog -Value $originalXml429Log
    Set-Item Function:\Start-Sleep -Value $originalXml429Sleep
    if ($null -ne $originalXml429WebRequest) {
        Set-Item Function:\Invoke-WebRequest -Value $originalXml429WebRequest
    }
    else {
        Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue
    }
    Set-HddtSharedState -Shared $null
    Reset-HddtStopRequest
}

# --- Request that mang header profile (khong con UA legacy) ---
$script:CapturedProfileRequestHeaders = $null
$originalProfileWebRequest = ${function:Invoke-WebRequest}
try {
    Set-Item Function:\Invoke-WebRequest -Value {
        param($Uri, $Method, $Headers, $TimeoutSec, $UseBasicParsing, $WebSession, $ErrorAction, $Body, $ContentType, [uri]$Proxy, [pscredential]$ProxyCredential)
        $script:CapturedProfileRequestHeaders = $Headers
        return [pscustomobject]@{
            StatusCode = 200
            Content = '{"ok":true}'
            RawContentStream = (New-Object IO.MemoryStream (,[byte[]](1, 2, 3)))
        }
    }
    Reset-GdtBrowserProfile
    $profileRequestConfig = New-TestXmlConfig
    $profileRequestConfig.BrowserUserAgent = 'IntegrationAgent/9.9'
    Reset-HddtStopRequest

    Invoke-GdtRequest -Config $profileRequestConfig -Uri ($profileRequestConfig.BaseUrl + '/query/invoices?size=50') | Out-Null
    Assert-Equal 'IntegrationAgent/9.9' $script:CapturedProfileRequestHeaders['User-Agent'] 'Real request carries the browser profile UA'
    Assert-Equal $true ($script:CapturedProfileRequestHeaders['Referer'] -like 'https://hoadondientu.gdt.gov.vn*') 'Real request carries the profile Referer'
    Assert-Equal $true ($script:CapturedProfileRequestHeaders.ContainsKey('Request-Id')) 'Request-Id header is kept per request'
    Assert-Equal $true ($script:CapturedProfileRequestHeaders['Accept'] -like 'application/json*') 'Real request accepts JSON'
    Assert-Equal $false ($script:CapturedProfileRequestHeaders.ContainsKey('Accept-Encoding')) 'Accept-Encoding is left to the transport'
    Assert-Equal $false ($script:CapturedProfileRequestHeaders.ContainsKey('Connection')) 'Connection header is not forced'
    Assert-Equal 'Bearer test-token' $script:CapturedProfileRequestHeaders['Authorization'] 'Authorization uses the configured token'

    Invoke-GdtRequest -Config $profileRequestConfig -Uri ($profileRequestConfig.BaseUrl + '/query/invoices/export-xml?x=1') -AsBytes | Out-Null
    Assert-Equal $true ($script:CapturedProfileRequestHeaders['Accept'] -like '*application/zip*') 'Export request accepts zip downloads'
    Assert-Equal 'IntegrationAgent/9.9' $script:CapturedProfileRequestHeaders['User-Agent'] 'Export request uses the same browser profile'
}
finally {
    if ($null -ne $originalProfileWebRequest) {
        Set-Item Function:\Invoke-WebRequest -Value $originalProfileWebRequest
    }
    else {
        Remove-Item Function:\Invoke-WebRequest -ErrorAction SilentlyContinue
    }
    Reset-GdtBrowserProfile
    Reset-HddtStopRequest
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
