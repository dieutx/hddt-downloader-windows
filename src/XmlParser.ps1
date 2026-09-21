Set-StrictMode -Version 2.0

function Resolve-LocalInvoiceDirection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$RootDirectory,
        [Parameter(Mandatory = $true)][ValidateSet('auto', 'purchase', 'sold')][string]$ConfiguredDirection
    )

    if ($ConfiguredDirection -ne 'auto') { return $ConfiguredDirection }

    $root = [IO.Path]::GetFullPath($RootDirectory).TrimEnd([char[]]@('\', '/'))
    $fullPath = [IO.Path]::GetFullPath($Path)
    $rootName = [IO.Path]::GetFileName($root)
    if ($rootName -in @('purchase', 'sold')) { return $rootName.ToLowerInvariant() }

    $relativePath = $fullPath
    $rootPrefix = $root + [IO.Path]::DirectorySeparatorChar
    if ($fullPath.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        $relativePath = $fullPath.Substring($rootPrefix.Length)
    }

    foreach ($segment in ($relativePath -split '[\\/]')) {
        $normalizedSegment = $segment.ToLowerInvariant()
        if ($normalizedSegment -in @('purchase', 'sold')) { return $normalizedSegment }
    }

    $baseName = [IO.Path]::GetFileNameWithoutExtension($fullPath)
    if ($baseName -match '(?i)^(purchase|sold)_') { return $Matches[1].ToLowerInvariant() }

    throw "Không xác định được XML mua vào hay bán ra: $fullPath. Hãy đặt file trong thư mục purchase/sold hoặc đặt LOCAL_DIRECTION rõ ràng."
}

function Get-XmlText {
    param($Node, [string]$XPath, [string]$Default = '')
    if ($null -eq $Node) { return $Default }
    $selected = $Node.SelectSingleNode($XPath)
    if ($null -eq $selected) { return $Default }
    return [string]$selected.InnerText
}

function Get-XmlAttribute {
    param($Node, [string]$Name)
    if ($null -eq $Node -or $null -eq $Node.Attributes[$Name]) { return '' }
    return [string]$Node.Attributes[$Name].Value
}

function ConvertFrom-XmlNumber {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $number = 0.0
    if ([double]::TryParse($Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    if ([double]::TryParse($Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::GetCultureInfo('vi-VN'), [ref]$number)) {
        return $number
    }
    return $Value
}

function ConvertFrom-XmlDate {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$date)) {
        return $date
    }
    return $Value
}

function Get-AdditionalXmlValue {
    param($Node, [string[]]$Names)
    if ($null -eq $Node) { return '' }
    $normalizedNames = @($Names | ForEach-Object { ($_ -replace '[^A-Za-z0-9]', '').ToLowerInvariant() })
    foreach ($info in $Node.SelectNodes(".//*[local-name()='TTin']")) {
        $fieldName = Get-XmlText $info "./*[local-name()='TTruong']"
        $normalized = ($fieldName -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
        if ($normalizedNames -contains $normalized) {
            return Get-XmlText $info "./*[local-name()='DLieu']"
        }
    }
    return ''
}

function ConvertFrom-InvoiceXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Direction,
        [Parameter(Mandatory = $true)][string]$Source
    )

    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $reader = [System.Xml.XmlReader]::Create($Path, $settings)
    try {
        $document = New-Object System.Xml.XmlDocument
        $document.PreserveWhitespace = $false
        $document.XmlResolver = $null
        $document.Load($reader)
    }
    finally {
        $reader.Dispose()
    }

    $invoice = $document.SelectSingleNode("//*[local-name()='HDon']")
    if ($null -eq $invoice) { throw "Không tìm thấy node HDon trong $Path" }
    $data = $invoice.SelectSingleNode("./*[local-name()='DLHDon']")
    if ($null -eq $data) { throw "Không tìm thấy node DLHDon trong $Path" }

    $general = $data.SelectSingleNode("./*[local-name()='TTChung']")
    $content = $data.SelectSingleNode("./*[local-name()='NDHDon']")
    $seller = $content.SelectSingleNode("./*[local-name()='NBan']")
    $buyer = $content.SelectSingleNode("./*[local-name()='NMua']")
    $totals = $content.SelectSingleNode("./*[local-name()='TToan']")

    $summary = [pscustomobject]@{
        Direction = $Direction
        Source = $Source
        XmlFile = [IO.Path]::GetFileName($Path)
        InvoiceId = Get-XmlAttribute $data 'Id'
        TemplateCode = Get-XmlText $general "./*[local-name()='KHMSHDon']"
        InvoiceSeries = Get-XmlText $general "./*[local-name()='KHHDon']"
        InvoiceNumber = Get-XmlText $general "./*[local-name()='SHDon']"
        InvoiceDate = ConvertFrom-XmlDate (Get-XmlText $general "./*[local-name()='NLap']")
        Currency = Get-XmlText $general "./*[local-name()='DVTTe']"
        ExchangeRate = ConvertFrom-XmlNumber (Get-XmlText $general "./*[local-name()='TGia']")
        SellerName = Get-XmlText $seller "./*[local-name()='Ten']"
        SellerTaxCode = Get-XmlText $seller "./*[local-name()='MST']"
        SellerAddress = Get-XmlText $seller "./*[local-name()='DChi']"
        BuyerName = Get-XmlText $buyer "./*[local-name()='Ten']"
        BuyerTaxCode = Get-XmlText $buyer "./*[local-name()='MST']"
        BuyerAddress = Get-XmlText $buyer "./*[local-name()='DChi']"
        TaxAuthorityCode = Get-XmlText $invoice ".//*[local-name()='MCCQT']"
        AmountBeforeTax = ConvertFrom-XmlNumber (Get-XmlText $totals "./*[local-name()='TgTCThue']")
        TaxAmount = ConvertFrom-XmlNumber (Get-XmlText $totals "./*[local-name()='TgTThue']")
        TotalAmount = ConvertFrom-XmlNumber (Get-XmlText $totals "./*[local-name()='TgTTTBSo']")
        TotalAmountText = Get-XmlText $totals "./*[local-name()='TgTTTBChu']"
        ProviderTaxCode = Get-XmlText $general "./*[local-name()='MSTTCGP']"
    }

    $details = New-Object System.Collections.Generic.List[object]
    $itemNodes = $content.SelectNodes("./*[local-name()='DSHHDVu']/*[local-name()='HHDVu']")
    foreach ($item in $itemNodes) {
        $taxAmountText = Get-XmlText $item "./*[local-name()='TThue']"
        if ([string]::IsNullOrWhiteSpace($taxAmountText)) {
            $taxAmountText = Get-AdditionalXmlValue $item @('TThue', 'TienThueGTGT', 'ThueGTGT')
        }
        $amountWithTaxText = Get-AdditionalXmlValue $item @('THTienCoVAT', 'ThanhTienCoVAT', 'TongTienCoThue')
        $amountBeforeTax = ConvertFrom-XmlNumber (Get-XmlText $item "./*[local-name()='ThTien']")
        $taxAmount = ConvertFrom-XmlNumber $taxAmountText
        $amountWithTax = ConvertFrom-XmlNumber $amountWithTaxText
        if ($null -eq $amountWithTax -and $amountBeforeTax -is [ValueType] -and $taxAmount -is [ValueType]) {
            $amountWithTax = [double]$amountBeforeTax + [double]$taxAmount
        }

        $details.Add([pscustomobject]@{
            Direction = $Direction
            Source = $Source
            XmlFile = [IO.Path]::GetFileName($Path)
            InvoiceSeries = $summary.InvoiceSeries
            InvoiceNumber = $summary.InvoiceNumber
            InvoiceDate = $summary.InvoiceDate
            SellerTaxCode = $summary.SellerTaxCode
            BuyerTaxCode = $summary.BuyerTaxCode
            LineNumber = Get-XmlText $item "./*[local-name()='STT']"
            Nature = Get-XmlText $item "./*[local-name()='TChat']"
            ProductCode = Get-XmlText $item "./*[local-name()='MHHDVu']"
            Description = (Get-XmlText $item "./*[local-name()='THHDVu']" (Get-XmlText $item "./*[local-name()='Ten']"))
            Unit = Get-XmlText $item "./*[local-name()='DVTinh']"
            Quantity = ConvertFrom-XmlNumber (Get-XmlText $item "./*[local-name()='SLuong']")
            UnitPrice = ConvertFrom-XmlNumber (Get-XmlText $item "./*[local-name()='DGia']")
            DiscountRate = ConvertFrom-XmlNumber (Get-XmlText $item "./*[local-name()='TLCKhau']")
            DiscountAmount = ConvertFrom-XmlNumber (Get-XmlText $item "./*[local-name()='STCKhau']")
            TaxRate = Get-XmlText $item "./*[local-name()='TSuat']"
            AmountBeforeTax = $amountBeforeTax
            TaxAmount = $taxAmount
            AmountWithTax = $amountWithTax
        })
    }

    return [pscustomobject]@{ Summary = $summary; Details = $details.ToArray() }
}
