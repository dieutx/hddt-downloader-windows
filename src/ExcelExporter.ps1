Set-StrictMode -Version 2.0

$script:SpreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:RelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:PackageRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
$script:ContentTypeNamespace = 'http://schemas.openxmlformats.org/package/2006/content-types'

function Get-ExcelColumnName {
    param([Parameter(Mandatory = $true)][int]$Index)
    if ($Index -lt 1) { throw 'Column index must be at least 1.' }
    $name = ''
    while ($Index -gt 0) {
        $Index--
        $name = [char](65 + ($Index % 26)) + $name
        $Index = [Math]::Floor($Index / 26)
    }
    return $name
}

function New-Utf8XmlWriter {
    param([Parameter(Mandatory = $true)][string]$Path)
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    return [System.Xml.XmlWriter]::Create($Path, $settings)
}

function Write-XmlStartElement {
    param($Writer, [string]$Name, [string]$Namespace = $script:SpreadsheetNamespace)
    [void]$Writer.WriteStartElement($Name, $Namespace)
}

function Write-ExcelCell {
    param(
        [Parameter(Mandatory = $true)]$Writer,
        [Parameter(Mandatory = $true)][string]$Reference,
        $Value,
        [int]$Style = 0
    )

    Write-XmlStartElement $Writer 'c'
    $Writer.WriteAttributeString('r', $Reference)
    if ($Style -gt 0) { $Writer.WriteAttributeString('s', [string]$Style) }

    if ($null -eq $Value) {
        $Writer.WriteEndElement()
        return
    }

    if ($Value -is [bool]) {
        $Writer.WriteAttributeString('t', 'b')
        Write-XmlStartElement $Writer 'v'
        $Writer.WriteString($(if ($Value) { '1' } else { '0' }))
        $Writer.WriteEndElement()
    }
    elseif ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        $Writer.WriteAttributeString('t', 'n')
        Write-XmlStartElement $Writer 'v'
        $Writer.WriteString([Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture))
        $Writer.WriteEndElement()
    }
    else {
        $text = if ($Value -is [datetime]) { $Value.ToString('yyyy-MM-dd HH:mm:ss') } else { [string]$Value }
        $Writer.WriteAttributeString('t', 'inlineStr')
        Write-XmlStartElement $Writer 'is'
        Write-XmlStartElement $Writer 't'
        if ($text.Length -gt 0 -and ($text[0] -eq ' ' -or $text[$text.Length - 1] -eq ' ')) {
            $Writer.WriteAttributeString('xml', 'space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
        }
        $Writer.WriteString($text)
        $Writer.WriteEndElement()
        $Writer.WriteEndElement()
    }
    $Writer.WriteEndElement()
}

function Get-ColumnWidths {
    param([string[]]$Columns, $Rows)
    $widths = New-Object double[] $Columns.Count
    for ($column = 0; $column -lt $Columns.Count; $column++) {
        $widths[$column] = [Math]::Min(60, [Math]::Max(10, $Columns[$column].Length + 2))
    }
    foreach ($row in @($Rows)) {
        for ($column = 0; $column -lt $Columns.Count; $column++) {
            $property = $row.PSObject.Properties[$Columns[$column]]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $length = ([string]$property.Value).Length + 2
            if ($length -gt $widths[$column]) { $widths[$column] = [Math]::Min(60, $length) }
        }
    }
    return $widths
}

function Write-WorksheetXml {
    param([string]$Path, [string[]]$Columns, $Rows, [int]$TableId)

    $rowItems = @($Rows)
    $lastColumn = Get-ExcelColumnName $Columns.Count
    $lastRow = $rowItems.Count + 1
    $tableReference = 'A1:{0}{1}' -f $lastColumn, $lastRow
    $writer = New-Utf8XmlWriter $Path
    try {
        Write-XmlStartElement $writer 'worksheet'
        $writer.WriteAttributeString('xmlns', 'r', $null, $script:RelationshipNamespace)

        Write-XmlStartElement $writer 'sheetViews'
        Write-XmlStartElement $writer 'sheetView'
        $writer.WriteAttributeString('workbookViewId', '0')
        Write-XmlStartElement $writer 'pane'
        $writer.WriteAttributeString('ySplit', '1')
        $writer.WriteAttributeString('topLeftCell', 'A2')
        $writer.WriteAttributeString('activePane', 'bottomLeft')
        $writer.WriteAttributeString('state', 'frozen')
        $writer.WriteEndElement()
        $writer.WriteEndElement()
        $writer.WriteEndElement()

        Write-XmlStartElement $writer 'cols'
        $widths = Get-ColumnWidths -Columns $Columns -Rows $rowItems
        for ($column = 0; $column -lt $Columns.Count; $column++) {
            Write-XmlStartElement $writer 'col'
            $writer.WriteAttributeString('min', [string]($column + 1))
            $writer.WriteAttributeString('max', [string]($column + 1))
            $writer.WriteAttributeString('width', [Convert]::ToString($widths[$column], [Globalization.CultureInfo]::InvariantCulture))
            $writer.WriteAttributeString('customWidth', '1')
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        Write-XmlStartElement $writer 'sheetData'
        Write-XmlStartElement $writer 'row'
        $writer.WriteAttributeString('r', '1')
        for ($column = 0; $column -lt $Columns.Count; $column++) {
            Write-ExcelCell -Writer $writer -Reference ((Get-ExcelColumnName ($column + 1)) + '1') -Value $Columns[$column] -Style 1
        }
        $writer.WriteEndElement()

        for ($row = 0; $row -lt $rowItems.Count; $row++) {
            $excelRow = $row + 2
            Write-XmlStartElement $writer 'row'
            $writer.WriteAttributeString('r', [string]$excelRow)
            for ($column = 0; $column -lt $Columns.Count; $column++) {
                $property = $rowItems[$row].PSObject.Properties[$Columns[$column]]
                $value = if ($null -eq $property) { $null } else { $property.Value }
                $reference = (Get-ExcelColumnName ($column + 1)) + $excelRow
                Write-ExcelCell -Writer $writer -Reference $reference -Value $value
            }
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        Write-XmlStartElement $writer 'autoFilter'
        $writer.WriteAttributeString('ref', $tableReference)
        $writer.WriteEndElement()

        if ($rowItems.Count -gt 0) {
            Write-XmlStartElement $writer 'tableParts'
            $writer.WriteAttributeString('count', '1')
            Write-XmlStartElement $writer 'tablePart'
            $writer.WriteAttributeString('r', 'id', $script:RelationshipNamespace, 'rId1')
            $writer.WriteEndElement()
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()
    }
    finally { $writer.Dispose() }

    return [pscustomobject]@{ Reference = $tableReference; HasTable = ($rowItems.Count -gt 0); TableId = $TableId }
}

function Write-TableXml {
    param([string]$Path, [int]$Id, [string]$Name, [string]$Reference, [string[]]$Columns)
    $writer = New-Utf8XmlWriter $Path
    try {
        Write-XmlStartElement $writer 'table'
        $writer.WriteAttributeString('id', [string]$Id)
        $writer.WriteAttributeString('name', $Name)
        $writer.WriteAttributeString('displayName', $Name)
        $writer.WriteAttributeString('ref', $Reference)
        $writer.WriteAttributeString('totalsRowShown', '0')
        Write-XmlStartElement $writer 'autoFilter'
        $writer.WriteAttributeString('ref', $Reference)
        $writer.WriteEndElement()
        Write-XmlStartElement $writer 'tableColumns'
        $writer.WriteAttributeString('count', [string]$Columns.Count)
        for ($column = 0; $column -lt $Columns.Count; $column++) {
            Write-XmlStartElement $writer 'tableColumn'
            $writer.WriteAttributeString('id', [string]($column + 1))
            $writer.WriteAttributeString('name', $Columns[$column])
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()
        Write-XmlStartElement $writer 'tableStyleInfo'
        $writer.WriteAttributeString('name', 'TableStyleMedium2')
        $writer.WriteAttributeString('showFirstColumn', '0')
        $writer.WriteAttributeString('showLastColumn', '0')
        $writer.WriteAttributeString('showRowStripes', '1')
        $writer.WriteAttributeString('showColumnStripes', '0')
        $writer.WriteEndElement()
        $writer.WriteEndElement()
    }
    finally { $writer.Dispose() }
}

function Write-TextFileUtf8NoBom {
    param([string]$Path, [string]$Content)
    [IO.File]::WriteAllText($Path, $Content, (New-Object Text.UTF8Encoding($false)))
}

function Export-InvoiceWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$SummaryRows,
        [Parameter(Mandatory = $true)]$DetailRows,
        [Parameter(Mandatory = $true)]$ErrorRows,
        [switch]$Overwrite
    )

    if (Test-Path -LiteralPath $Path) {
        if (-not $Overwrite) { throw "File Excel đã tồn tại: $Path. Đặt OVERWRITE_OUTPUT=true để ghi đè." }
        Remove-Item -LiteralPath $Path -Force
    }
    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }

    $summaryColumns = @('Direction','Source','XmlFile','InvoiceId','TemplateCode','InvoiceSeries','InvoiceNumber','InvoiceDate','Currency','ExchangeRate','SellerName','SellerTaxCode','SellerAddress','BuyerName','BuyerTaxCode','BuyerAddress','TaxAuthorityCode','AmountBeforeTax','TaxAmount','TotalAmount','TotalAmountText','ProviderTaxCode')
    $detailColumns = @('Direction','Source','XmlFile','InvoiceSeries','InvoiceNumber','InvoiceDate','SellerTaxCode','BuyerTaxCode','LineNumber','Nature','ProductCode','Description','Unit','Quantity','UnitPrice','DiscountRate','DiscountAmount','TaxRate','AmountBeforeTax','TaxAmount','AmountWithTax')
    $errorColumns = @('Direction','Source','Invoice','Error')
    $sheets = @(
        [pscustomobject]@{ Name='TongHop'; Columns=$summaryColumns; Rows=@($SummaryRows); TableName='tblTongHop' },
        [pscustomobject]@{ Name='ChiTiet'; Columns=$detailColumns; Rows=@($DetailRows); TableName='tblChiTiet' },
        [pscustomobject]@{ Name='Loi'; Columns=$errorColumns; Rows=@($ErrorRows); TableName='tblLoi' }
    )

    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('hddt-xlsx-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path (Join-Path $tempRoot '_rels') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'xl\_rels') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'xl\worksheets\_rels') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'xl\tables') -Force | Out-Null

    try {
        $tableResults = @()
        for ($index = 0; $index -lt $sheets.Count; $index++) {
            $sheetNumber = $index + 1
            $sheetPath = Join-Path $tempRoot ('xl\worksheets\sheet{0}.xml' -f $sheetNumber)
            $tableResult = Write-WorksheetXml -Path $sheetPath -Columns $sheets[$index].Columns -Rows $sheets[$index].Rows -TableId $sheetNumber
            $tableResults += $tableResult
            if ($tableResult.HasTable) {
                Write-TableXml -Path (Join-Path $tempRoot ('xl\tables\table{0}.xml' -f $sheetNumber)) -Id $sheetNumber -Name $sheets[$index].TableName -Reference $tableResult.Reference -Columns $sheets[$index].Columns
                $sheetRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="$script:PackageRelationshipNamespace"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table$sheetNumber.xml"/></Relationships>
"@
                Write-TextFileUtf8NoBom -Path (Join-Path $tempRoot ('xl\worksheets\_rels\sheet{0}.xml.rels' -f $sheetNumber)) -Content $sheetRelationships
            }
        }

        $sheetOverrides = for ($index = 1; $index -le $sheets.Count; $index++) {
            '<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $index
        }
        $tableOverrides = for ($index = 0; $index -lt $sheets.Count; $index++) {
            if ($tableResults[$index].HasTable) { '<Override PartName="/xl/tables/table{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' -f ($index + 1) }
        }
        $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="$script:ContentTypeNamespace"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>$($sheetOverrides -join '')$($tableOverrides -join '')</Types>
"@
        Write-TextFileUtf8NoBom -Path (Join-Path $tempRoot '[Content_Types].xml') -Content $contentTypes

        $rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="$script:PackageRelationshipNamespace"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
"@
        Write-TextFileUtf8NoBom -Path (Join-Path $tempRoot '_rels\.rels') -Content $rootRels

        $sheetNodes = for ($index = 0; $index -lt $sheets.Count; $index++) { '<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f $sheets[$index].Name, ($index + 1) }
        $workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="$script:SpreadsheetNamespace" xmlns:r="$script:RelationshipNamespace"><bookViews><workbookView activeTab="0"/></bookViews><sheets>$($sheetNodes -join '')</sheets></workbook>
"@
        Write-TextFileUtf8NoBom -Path (Join-Path $tempRoot 'xl\workbook.xml') -Content $workbookXml

        $workbookRelationships = for ($index = 0; $index -lt $sheets.Count; $index++) { '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f ($index + 1) }
        $stylesRelationshipId = $sheets.Count + 1
        $workbookRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="$script:PackageRelationshipNamespace">$($workbookRelationships -join '')<Relationship Id="rId$stylesRelationshipId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
"@
        Write-TextFileUtf8NoBom -Path (Join-Path $tempRoot 'xl\_rels\workbook.xml.rels') -Content $workbookRelsXml

        $stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="$script:SpreadsheetNamespace"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>
"@
        Write-TextFileUtf8NoBom -Path (Join-Path $tempRoot 'xl\styles.xml') -Content $stylesXml

        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $outputStream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $archive = New-Object IO.Compression.ZipArchive($outputStream, [IO.Compression.ZipArchiveMode]::Create, $false)
            try {
                foreach ($file in Get-ChildItem -LiteralPath $tempRoot -File -Recurse) {
                    $entryName = $file.FullName.Substring($tempRoot.Length + 1).Replace('\', '/')
                    [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file.FullName, $entryName, [IO.Compression.CompressionLevel]::Optimal)
                }
            }
            finally { $archive.Dispose() }
        }
        finally { $outputStream.Dispose() }
    }
    finally {
        $resolvedTempRoot = [IO.Path]::GetFullPath($tempRoot)
        $systemTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        if ($resolvedTempRoot.StartsWith($systemTempRoot, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $resolvedTempRoot)) {
            Remove-Item -LiteralPath $resolvedTempRoot -Recurse -Force
        }
    }
}
