Set-StrictMode -Version 2.0

# Giu du lieu danh sach khi detail tra ve null/rong; detail co gia tri thi uu tien.
function Merge-GdtInvoiceData {
    param($Index, [Parameter(Mandatory = $true)]$Detail)
    $values = [ordered]@{}
    if ($null -ne $Index) {
        foreach ($property in $Index.PSObject.Properties) { $values[$property.Name] = $property.Value }
    }
    foreach ($property in $Detail.PSObject.Properties) {
        $value = $property.Value
        $empty = ($null -eq $value -or ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) -or ($value -is [array] -and $value.Count -eq 0))
        if ($empty -and $values.Contains($property.Name)) { continue }
        $values[$property.Name] = $value
    }
    return [pscustomobject]$values
}

# Chuyen JSON detail sang cung Summary/Details nhu XML, khong tao XML gia.
function ConvertFrom-GdtInvoiceDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Detail,
        [Parameter(Mandatory = $true)][ValidateSet('purchase', 'sold')][string]$Direction
    )
    $summary = [pscustomobject]@{
        Direction = $Direction; Source = 'api-detail'; XmlFile = ''; DataOrigin = 'api-detail'
        GdtIndex = $Detail
        InvoiceId = Get-JsonTextValue $Detail 'id'
        InvoiceType = Get-JsonTextValue $Detail 'tlhdon'
        TemplateCode = Get-JsonTextValue $Detail 'khmshdon'
        InvoiceSeries = Get-JsonTextValue $Detail 'khhdon'
        InvoiceNumber = Get-JsonTextValue $Detail 'shdon'
        InvoiceDateText = Get-JsonTextValue $Detail 'tdlap'
        InvoiceDate = ConvertFrom-XmlDate (Get-JsonTextValue $Detail 'tdlap')
        Currency = Get-JsonTextValue $Detail 'dvtte'
        ExchangeRate = ConvertFrom-XmlNumber (Get-JsonTextValue $Detail 'tgia')
        SellerName = Get-JsonTextValue $Detail 'nbten'
        SellerTaxCode = Get-JsonTextValue $Detail 'nbmst'
        SellerAddress = Get-JsonTextValue $Detail 'nbdchi'
        SellerSigningTime = ConvertFrom-XmlDate (Get-JsonTextValue $Detail 'nky')
        BuyerName = Get-JsonTextValue $Detail 'nmten'
        BuyerTaxCode = Get-JsonTextValue $Detail 'nmmst'
        BuyerAddress = Get-JsonTextValue $Detail 'nmdchi'
        TaxAuthorityCode = Get-JsonTextValue $Detail 'mhdon'
        TaxAuthoritySigningTime = ConvertFrom-XmlDate (Get-JsonTextValue $Detail 'ncma')
        AmountBeforeTax = ConvertFrom-XmlNumber (Get-JsonTextValue $Detail 'tgtcthue')
        TaxAmount = ConvertFrom-XmlNumber (Get-JsonTextValue $Detail 'tgtthue')
        TotalAmount = ConvertFrom-XmlNumber (Get-JsonTextValue $Detail 'tgtttbso')
        TotalAmountText = Get-JsonTextValue $Detail 'tgtttbchu'
        ProviderTaxCode = Get-JsonTextValue $Detail 'msttcgp'
        InvoiceLookupFields = @(Get-ObjectValue $Detail 'cttkhac' @())
        InvoiceAdditionalFields = @(Get-ObjectValue $Detail 'ttkhac' @())
    }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($item in @(Get-ObjectValue $Detail 'hdhhdvu' @())) {
        if ($null -eq $item) { continue }
        $amount = Get-ObjectValue $item 'thtcthue' $null
        if ($null -eq $amount) { $amount = Get-ObjectValue $item 'thtien' $null }
        $amount = ConvertFrom-XmlNumber ([string]$amount)
        $tax = ConvertFrom-XmlNumber (Get-JsonTextValue $item 'tthue')
        $withTax = $null
        if ($amount -is [ValueType] -and $tax -is [ValueType]) { $withTax = [double]$amount + [double]$tax }
        $rate = Get-ObjectValue $item 'tsuat' $null
        if ($null -eq $rate) { $rate = Get-JsonTextValue $item 'ltsuat' }
        $rows.Add([pscustomobject]@{
            Direction = $Direction; Source = 'api-detail'; XmlFile = ''
            InvoiceSeries = $summary.InvoiceSeries; InvoiceNumber = $summary.InvoiceNumber
            InvoiceDate = $summary.InvoiceDate; SellerTaxCode = $summary.SellerTaxCode; BuyerTaxCode = $summary.BuyerTaxCode
            LineNumber = Get-JsonTextValue $item 'stt'
            Nature = Get-JsonTextValue $item 'tchat'
            ProductCode = Get-JsonTextValue $item 'mhhdvu'
            Description = Get-JsonTextValue $item 'ten'
            Unit = Get-JsonTextValue $item 'dvtinh'
            Quantity = ConvertFrom-XmlNumber (Get-JsonTextValue $item 'sluong')
            UnitPrice = ConvertFrom-XmlNumber (Get-JsonTextValue $item 'dgia')
            DiscountRate = ConvertFrom-XmlNumber (Get-JsonTextValue $item 'tlckhau')
            DiscountAmount = ConvertFrom-XmlNumber (Get-JsonTextValue $item 'stckhau')
            TaxRate = $rate
            TaxType = Get-JsonTextValue $item 'ltsuat'
            ItemAdditionalFields = @(Get-ObjectValue $item 'ttkhac' @())
            InvoiceSummary = $summary
            AmountBeforeTax = $amount; TaxAmount = $tax; AmountWithTax = $withTax
        })
    }
    return [pscustomobject]@{ Summary = $summary; Details = $rows.ToArray() }
}
