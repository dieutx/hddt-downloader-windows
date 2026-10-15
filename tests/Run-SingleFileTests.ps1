[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$ExePath)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'src/DesktopUpdates.ps1')
. (Join-Path $root 'src/Logging.ps1')
. (Join-Path $root 'src/Config.ps1')
. (Join-Path $root 'src/Http.ps1')
. (Join-Path $root 'src/InvoiceApi.ps1')
. (Join-Path $root 'src/XmlParser.ps1')
. (Join-Path $root 'src/InvoiceDetail.ps1')
. (Join-Path $root 'src/ExcelExporter.ps1')
. (Join-Path $root 'src/BrowserProfile.ps1')
. (Join-Path $root 'src/XmlScheduler.ps1')
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

function Assert-SingleFile {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAILED: ' + $Message) }
}

function Get-SingleFileSheetCounts {
    param([Parameter(Mandatory = $true)][string]$Path)
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $names = @($archive.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' } | ForEach-Object { $_.FullName } | Sort-Object)
        $counts = foreach ($name in $names) {
            $entry = $archive.GetEntry($name)
            $reader = New-Object IO.StreamReader($entry.Open())
            try { $text = $reader.ReadToEnd() }
            finally { $reader.Dispose() }
            '{0}={1}' -f $name, ([regex]::Matches($text, '<row ')).Count
        }
        return ($counts -join ';')
    }
    finally { $archive.Dispose() }
}

function Get-SingleFileSheetText {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Name)
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $archive.GetEntry($Name)
        if ($null -eq $entry) { return '' }
        $reader = New-Object IO.StreamReader($entry.Open())
        try { return $reader.ReadToEnd() }
        finally { $reader.Dispose() }
    }
    finally { $archive.Dispose() }
}

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('hddt-single-test-' + [guid]::NewGuid().ToString('N'))
$testRoot = [IO.Path]::GetFullPath($testRoot)
$null = New-Item -ItemType Directory -Path $testRoot
try {
    $exeDirectory = Join-Path $testRoot 'only one executable'
    $null = New-Item -ItemType Directory -Path $exeDirectory
    $copiedExe = Join-Path $exeDirectory 'Renamed-HDDT.exe'
    Copy-Item -LiteralPath $ExePath -Destination $copiedExe
    $assembly = [Reflection.Assembly]::Load([IO.File]::ReadAllBytes($copiedExe))
    $payload = $assembly.GetManifestResourceStream('Hddt.Payload.zip')
    Assert-SingleFile ($null -ne $payload) 'EXE embeds its complete runtime'
    $embedded = New-Object IO.Compression.ZipArchive($payload, [IO.Compression.ZipArchiveMode]::Read)
    try {
        foreach ($entry in $embedded.Entries) {
            Assert-SingleFile ($entry.FullName -notmatch '(^|/)(output|xml|logs)(/|$)|(^|/)\.env$|hddt-settings\.json$') 'Payload excludes private files'
            $sourcePath = Join-Path $root $entry.FullName
            $stream = $entry.Open()
            $hash = [Security.Cryptography.SHA256]::Create()
            try { $actual = [BitConverter]::ToString($hash.ComputeHash($stream)).Replace('-', '') }
            finally { $stream.Dispose(); $hash.Dispose() }
            Assert-SingleFile ($actual -eq (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash) ('Embedded source matches reviewed file: ' + $entry.FullName)
        }
        Assert-SingleFile ($null -ne $embedded.GetEntry('ui/hddt.ico')) 'Payload carries the app icon used by the WPF window'
    }
    finally { $embedded.Dispose(); $payload.Dispose() }
    $runDirectory = Join-Path $testRoot 'first run'
    $dataDirectory = Join-Path $runDirectory 'data'
    $null = New-Item -ItemType Directory -Path $dataDirectory -Force
    $settingsPath = Join-Path $dataDirectory 'hddt-settings.json'
    $settingsText = '{"XML_CONCURRENCY":"2","MAX_RETRIES":"1"}'
    [IO.File]::WriteAllText($settingsPath, $settingsText, (New-Object Text.UTF8Encoding($false)))

    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $copiedExe
    $info.Arguments = ('--smoke-test "{0}" "{1}"' -f $runDirectory, (Join-Path $PSScriptRoot 'fixtures'))
    $info.WorkingDirectory = $exeDirectory
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $process = [Diagnostics.Process]::Start($info)
    try {
        if (-not $process.WaitForExit(45000)) { $process.Kill(); throw 'FAILED: Single EXE smoke test timed out.' }
        if ($process.ExitCode -ne 0) {
            $errorPath = Join-Path $runDirectory 'launcher-error.txt'
            if (Test-Path -LiteralPath $errorPath) { throw ([IO.File]::ReadAllText($errorPath)) }
            throw 'FAILED: Single EXE launch failed.'
        }
    }
    finally { $process.Dispose() }

    Assert-SingleFile (@(Get-ChildItem -LiteralPath $exeDirectory -Force).Count -eq 1) 'EXE runs without companion files, even after rename'
    $report = [IO.File]::ReadAllText((Join-Path $runDirectory 'smoke-report.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    Assert-SingleFile ($report.Version -eq (Get-HddtAppVersion $root)) 'Embedded payload version matches source'
    Assert-SingleFile (-not (Test-Path -LiteralPath $report.RuntimeDirectory)) 'Runtime code is cleaned after normal exit'
    Assert-SingleFile ($report.SettingsPath -eq $settingsPath) 'Settings use persistent data directory'
    Assert-SingleFile ([IO.File]::ReadAllText($settingsPath) -ceq $settingsText) 'Existing settings are preserved'
    Assert-SingleFile (Test-Path -LiteralPath $report.OutputPath) 'Single EXE exports offline workbook'
    $archive = [IO.Compression.ZipFile]::OpenRead($report.OutputPath)
    try {
        Assert-SingleFile ($null -ne $archive.GetEntry('xl/workbook.xml')) 'Offline workbook is a valid package'
        Assert-SingleFile (@($archive.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' }).Count -eq 8) 'Offline workbook keeps eight sheets'
    }
    finally { $archive.Dispose() }

    $beforeCounts = Get-SingleFileSheetCounts -Path $report.OutputPath
    $sheetRows = @([regex]::Matches($beforeCounts, '=(\d+)') | ForEach-Object { [int]$_.Groups[1].Value })
    Assert-SingleFile ($beforeCounts.Split(';').Count -eq 8) 'Offline workbook has eight sheet parts'
    Assert-SingleFile (($sheetRows | Measure-Object -Maximum).Maximum -ge 3) 'Offline workbook contains invoice data rows'

    # File da co hoa don cua lan chay truoc (gia lap resume): them mot hoa don
    # khong co trong fixture roi chay lai app tren dung file do.
    $seedInvoice = [pscustomobject]@{
        Direction = 'purchase'; Source = 'query'; InvoiceId = 'seed-1'; TemplateCode = '1'
        InvoiceSeries = 'C26SEED'; InvoiceNumber = '99990001'; InvoiceDate = [datetime]'2026-01-02'
        Currency = 'VND'; ExchangeRate = 1; SellerName = 'Nha cung cap gia lap'; SellerTaxCode = '0999999991'
        BuyerName = 'Nguoi mua gia lap'; BuyerTaxCode = '0999999992'; AmountBeforeTax = 123456; TaxAmount = 12345
    }
    $seeded = Export-InvoiceWorkbook -Path $report.OutputPath -SummaryRows @($seedInvoice) -DetailRows @() -ErrorRows @() -Append
    Assert-SingleFile ($seeded.AppendMode) 'Seed export appends into the existing workbook'
    Assert-SingleFile ((Get-SingleFileSheetText -Path $report.OutputPath -Name 'xl/worksheets/sheet1.xml').Contains('99990001')) 'Seeded invoice is in the purchase summary sheet'
    $seededCounts = Get-SingleFileSheetCounts -Path $report.OutputPath
    Assert-SingleFile ($seededCounts -ne $beforeCounts) 'Seeding adds a row to the existing workbook'
    $stampBeforeRun = (Get-Item -LiteralPath $report.OutputPath).LastWriteTimeUtc

    $appendProcess = [Diagnostics.Process]::Start($info)
    try {
        if (-not $appendProcess.WaitForExit(45000)) { $appendProcess.Kill(); throw 'FAILED: Second single EXE run timed out.' }
        if ($appendProcess.ExitCode -ne 0) {
            $errorPath = Join-Path $runDirectory 'launcher-error.txt'
            if (Test-Path -LiteralPath $errorPath) { throw ([IO.File]::ReadAllText($errorPath)) }
            throw 'FAILED: Second single EXE run failed on an existing output file.'
        }
    }
    finally { $appendProcess.Dispose() }
    $appendArchive = [IO.Compression.ZipFile]::OpenRead($report.OutputPath)
    try {
        Assert-SingleFile (@($appendArchive.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' }).Count -eq 8) 'Second run keeps eight sheets after appending'
    }
    finally { $appendArchive.Dispose() }
    Assert-SingleFile ((Get-Item -LiteralPath $report.OutputPath).LastWriteTimeUtc -gt $stampBeforeRun) 'The app rewrites the workbook on a second run'
    Assert-SingleFile ((Get-SingleFileSheetCounts -Path $report.OutputPath) -ceq $seededCounts) 'A second run neither duplicates the re-parsed invoices nor drops existing ones'
    Assert-SingleFile ((Get-SingleFileSheetText -Path $report.OutputPath -Name 'xl/worksheets/sheet1.xml').Contains('99990001')) 'Invoice from the previous run survives the app run'
    Assert-SingleFile ((Get-SingleFileSheetText -Path $report.OutputPath -Name 'xl/worksheets/sheet1.xml').Contains('Nha cung cap gia lap')) 'Seeded row keeps its seller name'

    $release = [pscustomobject]@{ AssetName = 'HDDT-Downloader-v1.1.1-windows.exe'; Size = (Get-Item -LiteralPath $copiedExe).Length; Digest = ('sha256:' + (Get-FileHash -LiteralPath $copiedExe -Algorithm SHA256).Hash) }
    Assert-HddtReleasePackage -Path $copiedExe -Release $release
    $release.Digest = 'sha256:' + ('0' * 64)
    $rejected = $false
    try { Assert-HddtReleasePackage -Path $copiedExe -Release $release } catch { $rejected = $true }
    Assert-SingleFile $rejected 'Update package rejects wrong checksum'
    $badFile = Join-Path $testRoot 'invalid.exe'
    [IO.File]::WriteAllText($badFile, ('x' * 200))
    $release.Size = 200; $release.Digest = ''
    $rejected = $false
    try { Assert-HddtReleasePackage -Path $badFile -Release $release } catch { $rejected = $true }
    Assert-SingleFile $rejected 'Update package rejects an HTML/text response named EXE'
    Write-Output 'All single-file EXE tests passed.'
}
finally {
    $expectedParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if ([IO.Path]::GetDirectoryName($testRoot) -eq $expectedParent -and [IO.Path]::GetFileName($testRoot).StartsWith('hddt-single-test-')) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
