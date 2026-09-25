Set-StrictMode -Version 2.0

function Get-ObjectValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function ConvertTo-QueryValue {
    param([string]$Value)
    return [uri]::EscapeDataString($Value)
}

# Lấy giá trị text của một trường JSON (mảng/đối tượng trả về rỗng), tương ứng
# SafeJsonText trong VBA.
function Get-JsonTextValue {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return '' }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return '' }
    $value = $property.Value
    if ($value -is [string]) { return $value }
    if ($value -is [bool]) { return [string]$value }
    if ($value -is [ValueType]) { return [string]$value }
    return ''
}

# Trạng thái hóa đơn tthai: 2-5 có chuỗi liên quan, 6 chỉ có thông tin liên quan.
function ConvertTo-InvoiceStatus {
    param($Value)
    if ($null -eq $Value) { return 0 }
    $status = 0
    if ([int]::TryParse([string]$Value, [ref]$status)) { return $status }
    return 0
}

# Định dạng ngày ISO (tdlhdgoc, ngay) thành dd/MM/yyyy, tương ứng ISODATE.
function ConvertTo-RelatedDate {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $text = $Value.Trim()
    $parsed = [datetime]::MinValue
    if (-not [datetime]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsed)) {
        return $Value
    }
    if ($text.EndsWith('Z') -or $parsed.Kind -eq [DateTimeKind]::Utc) { $parsed = $parsed.ToLocalTime() }
    return $parsed.ToString('dd/MM/yyyy')
}

function Get-GdtFamilies {
    param($Config)
    $families = New-Object System.Collections.Generic.List[string]
    if ($Config.IncludeRegular) { $families.Add('query') }
    if ($Config.IncludeSco) { $families.Add('sco-query') }
    return $families.ToArray()
}

function Get-MonthDateRanges {
    param(
        [Parameter(Mandatory = $true)][datetime]$FromDate,
        [Parameter(Mandatory = $true)][datetime]$ToDate
    )

    $ranges = New-Object System.Collections.Generic.List[object]
    $current = $FromDate.Date
    while ($current -le $ToDate.Date) {
        $monthStart = New-Object datetime $current.Year, $current.Month, 1
        $monthEnd = $monthStart.AddMonths(1).AddDays(-1)
        $periodEnd = if ($monthEnd -lt $ToDate.Date) { $monthEnd } else { $ToDate.Date }
        $ranges.Add([pscustomobject]@{ From = $current; To = $periodEnd })
        $current = $periodEnd.AddDays(1)
    }
    return $ranges.ToArray()
}

function Get-GdtInvoiceIndex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][ValidateSet('purchase', 'sold')][string]$Direction
    )

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($family in Get-GdtFamilies $Config) {
        $familyStartCount = $results.Count
        foreach ($period in Get-MonthDateRanges -FromDate $Config.FromDate -ToDate $Config.ToDate) {
            $pageNumber = 0
            $seenStates = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
            Write-HddtLog INFO ('[DANH SÁCH] {0}/{1}, kỳ {2:dd/MM/yyyy}-{3:dd/MM/yyyy}.' -f $Direction, $family, $period.From, $period.To)
            $search = 'tdlap=ge={0}T00:00:00;tdlap=le={1}T23:59:59' -f $period.From.ToString('dd/MM/yyyy'), $period.To.ToString('dd/MM/yyyy')
            $state = $null
            do {
                if (Test-HddtStopRequested) {
                    Write-HddtLog WARN ('[DANH SÁCH] Đã dừng giữa chừng; giữ lại {0} hóa đơn {1}/{2} đã nhận.' -f $results.Count, $Direction, $family)
                    return $results.ToArray()
                }
                $pageNumber++
                $uri = '{0}/{1}/invoices/{2}?sort=tdlap%3Adesc&size={3}&search={4}' -f $Config.BaseUrl, $family, $Direction, $Config.PageSize, (ConvertTo-QueryValue $search)
                if (-not [string]::IsNullOrWhiteSpace([string]$state)) {
                    $uri += '&state=' + (ConvertTo-QueryValue ([string]$state))
                }

                $pageWatch = [Diagnostics.Stopwatch]::StartNew()
                $payload = (Invoke-GdtRequest -Config $Config -Uri $uri) | ConvertFrom-Json
                $pageWatch.Stop()
                $datas = Get-ObjectValue $payload 'datas' @()
                $pageItems = @($datas)
                foreach ($data in $pageItems) {
                    $status = ConvertTo-InvoiceStatus (Get-ObjectValue $data 'tthai' $null)
                    $results.Add([pscustomobject]@{
                        Direction = $Direction
                        Source = $family
                        SellerTaxCode = [string](Get-ObjectValue $data 'nbmst' '')
                        InvoiceSeries = [string](Get-ObjectValue $data 'khhdon' '')
                        InvoiceNumber = [string](Get-ObjectValue $data 'shdon' '')
                        InvoiceTemplate = [string](Get-ObjectValue $data 'khmshdon' '')
                        InvoiceDate = [string](Get-ObjectValue $data 'tdlap' '')
                        Status = $status
                        # Thông tin hóa đơn gốc lấy ngay từ danh sách (không tốn request),
                        # tương ứng WriteRelatedInvoiceInfo trong VBA.
                        OriginalInvoiceType = Get-JsonTextValue $data 'lhdgoc'
                        OriginalTemplateCode = Get-JsonTextValue $data 'khmshdgoc'
                        OriginalSeries = Get-JsonTextValue $data 'khhdgoc'
                        OriginalNumber = Get-JsonTextValue $data 'shdgoc'
                        OriginalDate = ConvertTo-RelatedDate (Get-JsonTextValue $data 'tdlhdgoc')
                        OriginalNote = Get-JsonTextValue $data 'gchdgoc'
                        RelatedChain = Build-RelatedInvoiceChain -Item $data
                        RelatedInfo = ''
                    })
                }
                $state = [string](Get-ObjectValue $payload 'state' '')
                if (-not [string]::IsNullOrWhiteSpace($state)) {
                    $state = $state.Trim()
                    if (-not $seenStates.Add($state)) {
                        Write-HddtLog WARN ('[DANH SÁCH] {0}/{1}, kỳ {2:dd/MM/yyyy}-{3:dd/MM/yyyy}: API trả lại state cũ; dừng phân trang.' -f $Direction, $family, $period.From, $period.To)
                        $state = ''
                    }
                }
                $nextText = if ([string]::IsNullOrWhiteSpace([string]$state)) { 'hết trang' } else { 'còn trang' }
                Write-HddtLog INFO ('[DANH SÁCH] Trang {0}: HTTP 200 | +{1} hóa đơn | lũy kế {2} | {3} | {4} ms.' -f $pageNumber, $pageItems.Count, ($results.Count - $familyStartCount), $nextText, $pageWatch.ElapsedMilliseconds)
            } while (-not [string]::IsNullOrWhiteSpace([string]$state))
        }
        Write-HddtLog INFO ('[DANH SÁCH] Hoàn tất {0}/{1}: {2} hóa đơn.' -f $Direction, $family, ($results.Count - $familyStartCount))
    }
    return $results.ToArray()
}

# ---------------------------------------------------------------------------
# Chuỗi hóa đơn liên quan (tương ứng ProcessRelatedInvoiceApis trong VBA):
# - status 2-5: gọi API relative (chuỗi thay thế/điều chỉnh) và related.
# - status 6: không có chuỗi, chỉ gọi API related.
# - status khác: không gọi API nào.
# ---------------------------------------------------------------------------

function Get-RelatedNoticeNature {
    param([string]$Code)
    if ([string]::IsNullOrWhiteSpace($Code)) { return '' }
    $value = 0
    if (-not [int]::TryParse($Code, [ref]$value)) { return $Code }
    switch ($value) {
        1 { return 'Hủy' }
        2 { return 'Điều chỉnh' }
        3 { return 'Thay thế' }
        4 { return 'Giải trình' }
        default { return $Code }
    }
}

function Get-OriginalInvoiceDescription {
    param($Item)
    $template = Get-JsonTextValue $Item 'khmshdgoc'
    $series = Get-JsonTextValue $Item 'khhdgoc'
    $number = Get-JsonTextValue $Item 'shdgoc'
    if (($template + $series + $number) -eq '') { return '' }
    $description = '{0}, ký hiệu hóa đơn {1}, số hóa đơn {2}' -f $template, $series, $number
    $status = ConvertTo-InvoiceStatus (Get-JsonTextValue $Item 'tthai')
    switch ($status) {
        2 { return 'Thay thế cho hóa đơn có ký hiệu mẫu số ' + $description }
        4 { return 'Thay thế cho hóa đơn có ký hiệu mẫu số ' + $description }
        3 { return 'Điều chỉnh cho hóa đơn có ký hiệu mẫu số ' + $description }
        5 { return 'Điều chỉnh cho hóa đơn có ký hiệu mẫu số ' + $description }
        6 { return 'Hủy/liên quan đến hóa đơn có ký hiệu mẫu số ' + $description }
        default { return 'Liên quan đến hóa đơn có ký hiệu mẫu số ' + $description }
    }
}

function Build-RelatedInvoiceKey {
    param($Item, [bool]$UseOriginalFields)
    $template = ''
    $series = ''
    $number = ''
    if ($UseOriginalFields) {
        $template = Get-JsonTextValue $Item 'khmshdgoc'
        $series = Get-JsonTextValue $Item 'khhdgoc'
        $number = Get-JsonTextValue $Item 'shdgoc'
    }
    else {
        $template = Get-JsonTextValue $Item 'khmshdon'
        $series = Get-JsonTextValue $Item 'khhdon'
        $number = Get-JsonTextValue $Item 'shdon'
        if ([string]::IsNullOrWhiteSpace($template)) { $template = Get-JsonTextValue $Item 'khmshdgoc' }
        if ([string]::IsNullOrWhiteSpace($series)) { $series = Get-JsonTextValue $Item 'khhdgoc' }
        if ([string]::IsNullOrWhiteSpace($number)) { $number = Get-JsonTextValue $Item 'shdgoc' }
    }
    if (($template + $series + $number) -eq '') { return '' }
    return '{0} | {1} | {2}' -f $template, $series, $number
}

# Chuỗi liên quan suy ra từ dữ liệu danh sách (không tốn request),
# tương ứng phần chuỗi trong WriteRelatedInvoiceInfo.
function Build-RelatedInvoiceChain {
    param($Item)
    $summary = Build-RelatedInvoiceKey -Item $Item -UseOriginalFields $true
    foreach ($relatedItem in @(Get-ObjectValue $Item 'hdonLquans' @())) {
        if ($null -eq $relatedItem) { continue }
        $key = Build-RelatedInvoiceKey -Item $relatedItem -UseOriginalFields $false
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if ($summary.IndexOf($key, [StringComparison]::OrdinalIgnoreCase) -ge 0) { continue }
        if ([string]::IsNullOrWhiteSpace($summary)) { $summary = $key } else { $summary = $summary + "`r`n" + $key }
    }
    return $summary
}

function Get-GdtRelationActionHeader {
    param([string]$EndpointName, [string]$Source, [string]$Direction)
    $prefix = if ($EndpointName -eq 'relative') {
        'Xem%20h%C3%B3a%20%C4%91%C6%A1n%20li%C3%AAn%20quan%20'
    }
    else {
        'Xem%20th%C3%B4ng%20tin%20li%C3%AAn%20quan%20'
    }
    $sourceText = if ($Source -eq 'sco-query') { 'h%C3%B3a%20%C4%91%C6%A1n%20m%C3%A1y%20t%C3%ADnh%20ti%E1%BB%81n' } else { 'h%C3%B3a%20%C4%91%C6%A1n' }
    $directionText = if ($Direction -eq 'purchase') { 'mua%20v%C3%A0o' } else { 'b%C3%A1n%20ra' }
    return '{0}({1} {2})' -f $prefix, $sourceText, $directionText
}

function Get-GdtRelationUri {
    param($Config, $Invoice, [string]$EndpointName)
    $query = 'nbmst={0}&khmshdon={1}&khhdon={2}&shdon={3}' -f `
        (ConvertTo-QueryValue $Invoice.SellerTaxCode), `
        (ConvertTo-QueryValue $Invoice.InvoiceTemplate), `
        (ConvertTo-QueryValue $Invoice.InvoiceSeries), `
        (ConvertTo-QueryValue $Invoice.InvoiceNumber)
    return '{0}/{1}/invoices/{2}?{3}' -f $Config.BaseUrl, $Invoice.Source, $EndpointName, $query
}

function Get-GdtRelationHeaders {
    param([string]$EndpointName, [string]$Source, [string]$Direction)
    return @{
        Accept = 'application/json, text/plain, */*'
        'Accept-Language' = 'vi'
        'Action' = Get-GdtRelationActionHeader -EndpointName $EndpointName -Source $Source -Direction $Direction
        'End-Point' = '/tra-cuu/tra-cuu-hoa-don'
        'Referer' = 'https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don'
    }
}

# Thông báo lỗi ngay tại cột kết quả khi đã hết lượt thử,
# tương ứng WriteRelationRequestError trong VBA.
function Get-GdtRelationErrorText {
    param([string]$EndpointName, [System.Management.Automation.ErrorRecord]$ErrorRecord)
    $text = if ($EndpointName -eq 'relative') {
        'Lỗi: Không thể lấy chuỗi hóa đơn liên quan'
    }
    else {
        'Lỗi: Không thể lấy thông tin liên quan'
    }
    $attempts = Get-GdtLastRequestAttempts
    if ($attempts -gt 0) { $text = $text + (' sau {0} lần thử' -f $attempts) }
    if ($null -ne $ErrorRecord -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.Exception.Message)) {
        $text = $text + '. ' + $ErrorRecord.Exception.Message
    }
    if ($text.Length -gt 32000) { $text = $text.Substring(0, 32000) }
    return $text
}

# Biến thể mảng JSON từ API relative: hỗ trợ mảng gốc hoặc bọc trong "datas".
function ConvertFrom-GdtJsonArray {
    param([string]$ResponseText)
    if ([string]::IsNullOrWhiteSpace($ResponseText)) { throw 'Phản hồi API không hợp lệ' }
    $parsed = $null
    try { $parsed = $ResponseText | ConvertFrom-Json }
    catch { throw 'Phản hồi API không hợp lệ' }
    $dataProperty = $null
    if ($null -ne $parsed) { $dataProperty = $parsed.PSObject.Properties['datas'] }
    if ($null -ne $dataProperty -and $null -ne $dataProperty.Value) { return @($dataProperty.Value) }
    if ($parsed -is [Array]) { return @($parsed) }
    throw 'Phản hồi API không hợp lệ'
}

# Bố cục chuỗi liên quan từ API relative: từ hóa đơn mới nhất ngược về hóa
# đơn đang tra cứu, tương ứng WriteRelativeInvoiceData.
function ConvertTo-RelativeChainText {
    param([string]$ResponseText, $Invoice)
    if ($ResponseText -match '(?i)error') { throw 'Phản hồi API chứa lỗi' }
    # @() để giữ mảng khi PowerShell trả về scalar với mảng một phần tử.
    $items = @(ConvertFrom-GdtJsonArray -ResponseText $ResponseText)
    $lines = New-Object System.Collections.Generic.List[string]
    $index = 0
    for ($itemIndex = $items.Count - 1; $itemIndex -ge 0; $itemIndex--) {
        $item = $items[$itemIndex]
        if ($null -eq $item) { continue }
        $index++
        $line = '{0}. Hóa đơn có liên quan | {1} | {2} | {3}' -f `
            $index, `
            (Get-JsonTextValue $item 'khmshdon'), `
            (Get-JsonTextValue $item 'khhdon'), `
            (Get-JsonTextValue $item 'shdon')
        $description = Get-OriginalInvoiceDescription -Item $item
        if (-not [string]::IsNullOrWhiteSpace($description)) { $line = $line + ' | ' + $description }
        $lines.Add($line)
    }
    $index++
    $lines.Add(('{0}. Hóa đơn đang tra cứu | {1} | {2} | {3}' -f $index, $Invoice.InvoiceTemplate, $Invoice.InvoiceSeries, $Invoice.InvoiceNumber))
    return ($lines -join "`r`n")
}

function Build-RelatedNoticeLine {
    param($Notice)
    $name = Get-JsonTextValue $Notice 'ten'
    $date = Get-JsonTextValue $Notice 'ngay'
    $reason = Get-JsonTextValue $Notice 'ldo'
    $nature = Get-RelatedNoticeNature (Get-JsonTextValue $Notice 'loai')
    $receiveCode = Get-JsonTextValue $Notice 'kqtnhan'
    if (-not [string]::IsNullOrWhiteSpace($date)) { $date = ConvertTo-RelatedDate $date }
    $receiveResult = ''
    if (-not [string]::IsNullOrWhiteSpace($receiveCode)) {
        $code = 0
        [void][int]::TryParse($receiveCode, [ref]$code)
        $receiveResult = if ($code -eq 1) { 'Cơ quan thuế tiếp nhận.' } else { 'Cơ quan thuế không tiếp nhận.' }
    }
    $line = 'Hóa đơn có ' + $name
    if (-not [string]::IsNullOrWhiteSpace($date)) { $line = $line + ' ngày ' + $date }
    if (-not [string]::IsNullOrWhiteSpace($nature)) { $line = $line + '. Tính chất ' + $nature }
    if (-not [string]::IsNullOrWhiteSpace($reason)) { $line = $line + ', lý do ' + $reason }
    if (-not $line.EndsWith('.')) { $line = $line + '.' }
    if (-not [string]::IsNullOrWhiteSpace($receiveResult)) { $line = $line + ' ' + $receiveResult }
    return $line
}

function Build-RelatedNoticeSummary {
    param($Response)
    $empty = [pscustomobject]@{ Text = ''; HasContainer = $false }
    if ($null -eq $Response) { return $empty }
    $containerProperty = $Response.PSObject.Properties['mtthdtbssrs']
    if ($null -eq $containerProperty) { $containerProperty = $Response.PSObject.Properties['hdtbssrses'] }
    if ($null -eq $containerProperty) { return $empty }
    if ($null -eq $containerProperty.Value) { return [pscustomobject]@{ Text = ''; HasContainer = $true } }
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($notice in @($containerProperty.Value)) {
        if ($null -eq $notice) { continue }
        $line = Build-RelatedNoticeLine -Notice $notice
        if (-not [string]::IsNullOrWhiteSpace($line)) { $lines.Add($line) }
    }
    return [pscustomobject]@{ Text = ($lines -join "`r`n"); HasContainer = $true }
}

# Định dạng phản hồi API related, tương ứng WriteRelatedInformationResponse.
function ConvertTo-RelatedInformationText {
    param([string]$ResponseText)
    $normalized = [string]$ResponseText
    if ($null -eq $ResponseText) { $normalized = '' }
    $normalized = $normalized.Trim()
    if ($normalized -eq '' -or $normalized -eq '[]' -or $normalized -eq '{}' -or $normalized.ToLowerInvariant() -eq 'null') {
        return 'Không có thông tin liên quan'
    }
    try {
        $parsed = $normalized | ConvertFrom-Json
        $noticeSummary = Build-RelatedNoticeSummary -Response $parsed
        if (-not [string]::IsNullOrWhiteSpace($noticeSummary.Text)) {
            $normalized = $noticeSummary.Text
        }
        elseif ($noticeSummary.HasContainer) {
            $normalized = 'Không có thông tin liên quan'
        }
        else {
            $normalized = ConvertTo-Json -InputObject $parsed -Depth 10
        }
    }
    catch {
        # Không parse được thì giữ nguyên văn bản phản hồi như VBA.
    }
    if ($normalized.Length -gt 32000) { $normalized = $normalized.Substring(0, 32000) }
    return $normalized
}

function Test-GdtRelatedInvoice {
    param($Invoice)
    $status = ConvertTo-InvoiceStatus (Get-ObjectValue $Invoice 'Status' 0)
    return ($status -ge 2 -and $status -le 6)
}

# Lấy dữ liệu liên quan cho một hóa đơn. Không ném lỗi ra ngoài: mọi lỗi đều
# được ghi ngay vào ô kết quả tương ứng như VBA.
function Get-GdtInvoiceRelation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Invoice,
        # Chỉ dùng dữ liệu đã có từ danh sách, không gọi API relative/related.
        [switch]$ListOnly
    )

    $status = ConvertTo-InvoiceStatus (Get-ObjectValue $Invoice 'Status' 0)
    $relation = [pscustomobject]@{
        RelatedChain = [string](Get-ObjectValue $Invoice 'RelatedChain' '')
        OriginalInvoiceType = [string](Get-ObjectValue $Invoice 'OriginalInvoiceType' '')
        OriginalTemplateCode = [string](Get-ObjectValue $Invoice 'OriginalTemplateCode' '')
        OriginalSeries = [string](Get-ObjectValue $Invoice 'OriginalSeries' '')
        OriginalNumber = [string](Get-ObjectValue $Invoice 'OriginalNumber' '')
        OriginalDate = [string](Get-ObjectValue $Invoice 'OriginalDate' '')
        OriginalNote = [string](Get-ObjectValue $Invoice 'OriginalNote' '')
        RelatedInfo = ''
    }
    if ($status -lt 2 -or $status -gt 6) { return $null }

    if ($status -eq 6) {
        # Hóa đơn hủy: không có chuỗi liên quan để lấy (WriteNoRelativeInvoiceData).
        $relation.RelatedChain = 'Không có thông tin hiển thị'
    }
    elseif ($status -le 5 -and -not $ListOnly) {
        $relativeUri = Get-GdtRelationUri -Config $Config -Invoice $Invoice -EndpointName 'relative'
        try {
            $relativeText = Invoke-GdtRequest -Config $Config -Uri $relativeUri `
                -ExtraHeaders (Get-GdtRelationHeaders -EndpointName 'relative' -Source $Invoice.Source -Direction $Invoice.Direction)
            $relation.RelatedChain = ConvertTo-RelativeChainText -ResponseText $relativeText -Invoice $Invoice
            Write-HddtLog DEBUG ('Đã lấy chuỗi hóa đơn liên quan: {0}' -f (Get-InvoiceLabel $Invoice))
        }
        catch {
            if (Test-HddtStopRequested) { throw }
            $relation.RelatedChain = Get-GdtRelationErrorText -EndpointName 'relative' -ErrorRecord $_
            Write-HddtLog WARN ('[LIÊN QUAN] Không lấy được chuỗi hóa đơn liên quan cho {0}: {1}' -f (Get-InvoiceLabel $Invoice), $_.Exception.Message)
        }
    }

    if (-not $ListOnly) {
        $relatedUri = Get-GdtRelationUri -Config $Config -Invoice $Invoice -EndpointName 'related'
        try {
            $relatedText = Invoke-GdtRequest -Config $Config -Uri $relatedUri `
                -ExtraHeaders (Get-GdtRelationHeaders -EndpointName 'related' -Source $Invoice.Source -Direction $Invoice.Direction)
            $relation.RelatedInfo = ConvertTo-RelatedInformationText -ResponseText $relatedText
            Write-HddtLog DEBUG ('Đã lấy thông tin liên quan: {0}' -f (Get-InvoiceLabel $Invoice))
        }
        catch {
            if (Test-HddtStopRequested) { throw }
            $relation.RelatedInfo = Get-GdtRelationErrorText -EndpointName 'related' -ErrorRecord $_
            Write-HddtLog WARN ('[LIÊN QUAN] Không lấy được thông tin liên quan cho {0}: {1}' -f (Get-InvoiceLabel $Invoice), $_.Exception.Message)
        }
    }

    return $relation
}

function ConvertTo-SafeFileName {
    param([string]$Value)
    $safe = $Value
    foreach ($character in [IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$character, '_')
    }
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'unknown' }
    return $safe
}

function Get-InvoiceLabel {
    param([Parameter(Mandatory = $true)]$Invoice)
    return '{0}/{1}/{2}' -f $Invoice.Direction, $Invoice.InvoiceSeries, $Invoice.InvoiceNumber
}

function Expand-InvoiceXmlBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [Parameter(Mandatory = $true)][string]$DestinationDirectory,
        [Parameter(Mandatory = $true)][string]$FileNamePrefix
    )

    Add-Type -AssemblyName System.IO.Compression
    New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null
    $memoryStream = New-Object IO.MemoryStream (,$Bytes)
    $archive = $null
    $xmlFiles = New-Object System.Collections.Generic.List[string]
    try {
        $archive = New-Object IO.Compression.ZipArchive($memoryStream, [IO.Compression.ZipArchiveMode]::Read, $false)
        $xmlEntries = @($archive.Entries | Where-Object { [IO.Path]::GetExtension($_.Name).ToLowerInvariant() -eq '.xml' })
        for ($index = 0; $index -lt $xmlEntries.Count; $index++) {
            $entry = $xmlEntries[$index]
            $suffix = if ($xmlEntries.Count -eq 1) { '' } else { '_{0}' -f ($index + 1) }
            $targetPath = Join-Path $DestinationDirectory ($FileNamePrefix + $suffix + '.xml')
            $inputStream = $entry.Open()
            $outputStream = [IO.File]::Create($targetPath)
            try { $inputStream.CopyTo($outputStream) }
            finally {
                $outputStream.Dispose()
                $inputStream.Dispose()
            }
            $xmlFiles.Add($targetPath)
        }
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
        $memoryStream.Dispose()
    }
    return $xmlFiles.ToArray()
}

function Save-GdtInvoiceXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Invoice
    )

    foreach ($required in 'SellerTaxCode', 'InvoiceSeries', 'InvoiceNumber', 'InvoiceTemplate') {
        if ([string]::IsNullOrWhiteSpace([string](Get-ObjectValue $Invoice $required ''))) {
            throw "Thiếu trường bắt buộc $required trong dữ liệu hóa đơn."
        }
    }

    $sellerTaxCode = ConvertTo-QueryValue $Invoice.SellerTaxCode
    $invoiceSeries = ConvertTo-QueryValue $Invoice.InvoiceSeries
    $invoiceNumber = ConvertTo-QueryValue $Invoice.InvoiceNumber
    $invoiceTemplate = ConvertTo-QueryValue $Invoice.InvoiceTemplate
    $payload = 'nbmst={0}&khhdon={1}&shdon={2}&khmshdon={3}' -f $sellerTaxCode, $invoiceSeries, $invoiceNumber, $invoiceTemplate
    $uri = '{0}/{1}/invoices/export-xml?{2}' -f $Config.BaseUrl, $Invoice.Source, $payload

    $directionDirectory = Join-Path $Config.XmlDirectory $Invoice.Direction
    New-Item -ItemType Directory -Path $directionDirectory -Force | Out-Null

    $baseName = ConvertTo-SafeFileName ('{0}_{1}_{2}_{3}_{4}_{5}' -f $Invoice.Direction, $Invoice.Source, $Invoice.SellerTaxCode, $Invoice.InvoiceTemplate, $Invoice.InvoiceSeries, $Invoice.InvoiceNumber)
    $namePattern = '^{0}(?:_\d+)?$' -f [regex]::Escape($baseName)
    $existingXmlFiles = @(Get-ChildItem -LiteralPath $directionDirectory -Filter '*.xml' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match $namePattern } |
        Select-Object -ExpandProperty FullName)
    if (-not $Config.RedownloadXml -and $existingXmlFiles.Count -gt 0) {
        Write-HddtLog DEBUG ('Tái sử dụng {0} XML đã tải trong thư mục {1}.' -f $existingXmlFiles.Count, $directionDirectory)
        return $existingXmlFiles
    }
    if ($Config.RedownloadXml) {
        foreach ($existingXmlFile in $existingXmlFiles) {
            Remove-Item -LiteralPath $existingXmlFile -Force
        }
    }

    $responseBytes = [byte[]](Invoke-GdtRequest -Config $Config -Uri $uri -AsBytes)
    try {
        $xmlFiles = @(Expand-InvoiceXmlBytes -Bytes $responseBytes -DestinationDirectory $directionDirectory -FileNamePrefix $baseName)
    }
    catch {
        throw "Không đọc được XML trả về cho $(Get-InvoiceLabel $Invoice): $($_.Exception.Message)"
    }

    if ($xmlFiles.Count -eq 0) { throw 'Phản hồi không chứa file XML.' }
    Write-HddtLog DEBUG ('Đã ghi {0} XML vào thư mục {1}; không lưu file ZIP.' -f $xmlFiles.Count, $directionDirectory)
    return $xmlFiles
}
