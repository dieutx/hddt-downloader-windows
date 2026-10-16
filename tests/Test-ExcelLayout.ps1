# Metadata mau VBA: kiem tra tren workbook da dong goi, ca sau khi ghi tiep.
Set-StrictMode -Version 2.0
$layoutPath = Join-Path ([IO.Path]::GetTempPath()) ('hddt-layout-' + [guid]::NewGuid().ToString('N') + '.xlsx')
try {
    $layoutInvoice = ConvertFrom-InvoiceXml -Path (Join-Path $PSScriptRoot 'fixtures/sample-invoice.xml') -Direction purchase -Source query
    $null = Export-InvoiceWorkbook -Path $layoutPath -SummaryRows @($layoutInvoice.Summary) -DetailRows @($layoutInvoice.Details) -ErrorRows @()
    for ($pass = 0; $pass -lt 2; $pass++) {
        if ($pass -eq 1) {
            $null = Export-InvoiceWorkbook -Path $layoutPath -SummaryRows @($layoutInvoice.Summary) -DetailRows @($layoutInvoice.Details) -ErrorRows @() -Append
        }
        $layoutZip = [IO.Compression.ZipFile]::OpenRead($layoutPath)
        try {
            [xml]$layoutStyles = Read-TestZipEntryText $layoutZip 'xl/styles.xml'
            $styleNodes = @($layoutStyles.styleSheet.cellXfs.ChildNodes)
            $fontNodes = @($layoutStyles.styleSheet.fonts.ChildNodes)
            $fillNodes = @($layoutStyles.styleSheet.fills.ChildNodes)
            Assert-Equal $styleNodes.Count ([int]$layoutStyles.styleSheet.cellXfs.count) 'Style count matches its entries'
            $customFormats = @($layoutStyles.styleSheet.numFmts.numFmt | ForEach-Object { [int]$_.numFmtId })
            Assert-Equal $customFormats.Count (@($customFormats | Select-Object -Unique).Count) 'Custom number format IDs are unique'
            foreach ($style in $styleNodes) {
                $formatId = [int]$style.numFmtId
                Assert-Equal $true ($formatId -lt 164 -or $customFormats -contains $formatId) 'Every custom number format is defined'
                Assert-Equal $true ([int]$style.fontId -lt $fontNodes.Count) 'Every style font is defined'
                Assert-Equal $true ([int]$style.fillId -lt $fillNodes.Count) 'Every style fill is defined'
                Assert-Equal $true ([int]$style.borderId -lt @($layoutStyles.styleSheet.borders.border).Count) 'Every style border is defined'
            }
            Assert-Equal 1 @($layoutZip.Entries | Where-Object FullName -eq 'xl/theme/theme1.xml').Count 'Original theme is packaged'
            [xml]$theme = Read-TestZipEntryText $layoutZip 'xl/theme/theme1.xml'
            Assert-Equal 'ED7D31' $theme.theme.themeElements.clrScheme.accent2.srgbClr.val 'Original orange theme accent'
            Assert-Equal '70AD47' $theme.theme.themeElements.clrScheme.accent6.srgbClr.val 'Original green theme accent'
            [xml]$rels = Read-TestZipEntryText $layoutZip 'xl/_rels/workbook.xml.rels'
            Assert-Equal 1 @($rels.Relationships.Relationship | Where-Object { $_.Type -like '*/theme' -and $_.Target -eq 'theme/theme1.xml' }).Count 'Theme relationship points at packaged theme'
            $expectedThemes = @(5, 7, -1, 9, 8, -1)
            $expectedRows = @('17.25', '19.5', '19.5', '17.25', '19.5', '19.5', '15', '18')
            $expectedHeaders = @('48', '51', '51', '48', '51', '51', '30', '24.75')
            $expectedColumns = @(64, 34, 32, 64, 34, 32, 17, 15)
            for ($sn = 1; $sn -le 8; $sn++) {
                [xml]$sheet = Read-TestZipEntryText $layoutZip ('xl/worksheets/sheet{0}.xml' -f $sn)
                $sm = New-Object Xml.XmlNamespaceManager($sheet.NameTable)
                $sm.AddNamespace('x', $script:SpreadsheetNamespace)
                Assert-Equal $expectedRows[$sn - 1] $sheet.worksheet.sheetFormatPr.defaultRowHeight ('VBA default row height: ' + $sn)
                Assert-Equal $expectedColumns[$sn - 1] @($sheet.worksheet.cols.col).Count ('VBA column count: ' + $sn)
                $headerNumber = if ($sn -le 6) { 2 } else { 1 }
                $header = $sheet.SelectSingleNode(('/x:worksheet/x:sheetData/x:row[@r="{0}"]' -f $headerNumber), $sm)
                Assert-Equal $expectedHeaders[$sn - 1] $header.GetAttribute('ht') ('VBA header height: ' + $sn)
                if ($sn -le 6) {
                    $title = $sheet.SelectSingleNode('/x:worksheet/x:sheetData/x:row[@r="1"]', $sm)
                    $titleHeight = if ($sn -in @(1, 4)) { '41.25' } else { '34.5' }
                    Assert-Equal $titleHeight $title.GetAttribute('ht') ('VBA title height: ' + $sn)
                    $headerCell = Get-TestCellNode $sheet 'A2'
                    $headerStyle = $styleNodes[[int]$headerCell.GetAttribute('s')]
                    Assert-Equal 'Consolas' $fontNodes[[int]$headerStyle.fontId].name.val ('VBA header font: ' + $sn)
                    $color = $fillNodes[[int]$headerStyle.fillId].patternFill.fgColor
                    if ($expectedThemes[$sn - 1] -ge 0) {
                        Assert-Equal ([string]$expectedThemes[$sn - 1]) $color.GetAttribute('theme') ('Distinct VBA sheet color: ' + $sn)
                    } else { Assert-Equal 'FFE0E1AD' $color.GetAttribute('rgb') ('XML olive header: ' + $sn) }
                    Assert-Equal 'center' $headerStyle.alignment.horizontal ('VBA centered header: ' + $sn)
                    $expectedWrap = if ($sn -in @(1, 4)) { '' } else { '1' }
                    Assert-Equal $expectedWrap $headerStyle.alignment.GetAttribute('wrapText') ('VBA header wrapping: ' + $sn)
                }
                if ($sn -in @(1, 4)) {
                    Assert-Equal '1' $sheet.worksheet.sheetViews.sheetView.pane.xSplit 'Summary freezes first column'
                    Assert-Equal 'B3' $sheet.worksheet.sheetViews.sheetView.pane.topLeftCell 'Summary opens at first visible data column'
                    Assert-Equal ('A2:BA' + $(if ($sn -eq 1) { '3' } else { '2' })) $sheet.worksheet.autoFilter.ref 'Summary filter follows original column boundary'
                }
                Assert-Equal '0.7' $sheet.worksheet.pageMargins.left 'Original print margin'
                foreach ($cell in $sheet.SelectNodes('/x:worksheet/x:sheetData/x:row/x:c', $sm)) {
                    $sid = [int]$cell.GetAttribute('s')
                    Assert-Equal $true ($sid -ge 0 -and $sid -lt $styleNodes.Count) 'Every cell style is defined'
                }
            }
            [xml]$purchase = Read-TestZipEntryText $layoutZip 'xl/worksheets/sheet1.xml'
            [xml]$sale = Read-TestZipEntryText $layoutZip 'xl/worksheets/sheet4.xml'
            Assert-Equal '51.5703125' $purchase.worksheet.cols.col[8].width 'Purchase seller width follows VBA'
            Assert-Equal '46.7109375' $sale.worksheet.cols.col[8].width 'Sales seller width follows VBA'
            Assert-Equal '69.7109375' $purchase.worksheet.cols.col[52].width 'Purchase validation width follows VBA'
            Assert-Equal '24.42578125' $sale.worksheet.cols.col[52].width 'Sales validation width follows VBA'
            Assert-Equal '123' (Get-TestCellValue $purchase 'E3') 'Append retains numeric invoice identity'
            Assert-Equal 3 @($purchase.worksheet.sheetData.row).Count 'Appending the same invoice does not duplicate data'
            $dateNode = Get-TestCellNode $purchase 'F3'
            Assert-Equal '165' (Get-TestCellFormat $layoutZip $dateNode) 'Append retains date number format'
            [xml]$detail = Read-TestZipEntryText $layoutZip 'xl/worksheets/sheet2.xml'
            Assert-Equal '0.1' (Get-TestCellValue $detail 'Z3') 'Append retains tax rate fraction'
            Assert-Equal '9' (Get-TestCellFormat $layoutZip (Get-TestCellNode $detail 'Z3')) 'Append retains percent format'
            Assert-Equal '49' (Get-TestCellFormat $layoutZip (Get-TestCellNode $detail 'H3')) 'Append retains text tax ID'
        }
        finally { $layoutZip.Dispose() }
    }
}
finally { Remove-Item -LiteralPath $layoutPath -Force -ErrorAction SilentlyContinue }
