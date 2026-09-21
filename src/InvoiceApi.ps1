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
            Write-HddtLog INFO ('Danh sách {0}/{1}, kỳ {2:dd/MM/yyyy}-{3:dd/MM/yyyy}.' -f $Direction, $family, $period.From, $period.To)
            $search = 'tdlap=ge={0}T00:00:00;tdlap=le={1}T23:59:59' -f $period.From.ToString('dd/MM/yyyy'), $period.To.ToString('dd/MM/yyyy')
            $state = $null
            do {
                $pageNumber++
                $uri = '{0}/{1}/invoices/{2}?sort=tdlap%3Adesc&size={3}&search={4}' -f $Config.BaseUrl, $family, $Direction, $Config.PageSize, (ConvertTo-QueryValue $search)
                if (-not [string]::IsNullOrWhiteSpace([string]$state)) {
                    $uri += '&state=' + (ConvertTo-QueryValue ([string]$state))
                }

                $payload = (Invoke-GdtRequest -Config $Config -Uri $uri) | ConvertFrom-Json
                $datas = Get-ObjectValue $payload 'datas' @()
                $pageItems = @($datas)
                foreach ($data in $pageItems) {
                    $results.Add([pscustomobject]@{
                        Direction = $Direction
                        Source = $family
                        SellerTaxCode = [string](Get-ObjectValue $data 'nbmst' '')
                        InvoiceSeries = [string](Get-ObjectValue $data 'khhdon' '')
                        InvoiceNumber = [string](Get-ObjectValue $data 'shdon' '')
                        InvoiceTemplate = [string](Get-ObjectValue $data 'khmshdon' '')
                        InvoiceDate = [string](Get-ObjectValue $data 'tdlap' '')
                    })
                }
                $state = Get-ObjectValue $payload 'state' $null
                $nextText = if ([string]::IsNullOrWhiteSpace([string]$state)) { 'hết trang' } else { 'còn trang tiếp' }
                Write-HddtLog INFO ('  Trang {0}: nhận {1} hóa đơn; lũy kế nguồn {2}; {3}.' -f $pageNumber, $pageItems.Count, ($results.Count - $familyStartCount), $nextText)
            } while (-not [string]::IsNullOrWhiteSpace([string]$state))
        }
        Write-HddtLog INFO ('Hoàn tất nguồn {0}/{1}: {2} hóa đơn.' -f $Direction, $family, ($results.Count - $familyStartCount))
    }
    return $results.ToArray()
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
