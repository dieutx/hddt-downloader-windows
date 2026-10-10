[CmdletBinding()]
param([string]$OutputDirectory)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Chi build app WPF tren Windows.' }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path (Join-Path $PSScriptRoot 'work/desktop') ('HDDT-{0}-{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8)))
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Thu muc build da ton tai. Hay chon thu muc moi de khong ghi de du lieu.' }
$windowsDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::Windows)
$compiler = Join-Path $windowsDirectory 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) { throw 'Khong tim thay trinh bien dich .NET Framework cua Windows.' }

$null = New-Item -ItemType Directory -Path $OutputDirectory
$null = New-Item -ItemType Directory -Path (Join-Path $OutputDirectory 'src')
$null = New-Item -ItemType Directory -Path (Join-Path $OutputDirectory 'ui')
# Chi sao chep ma nguon can thiet; khong dong goi .env, output hay hoa don that.
$parts = @('Start-HddtGui.ps1', 'Invoke-Hddt.ps1', 'Parse-LocalXml.ps1', 'README.md', 'CHANGELOG.md',
    'src/Logging.ps1', 'src/Config.ps1', 'src/Http.ps1', 'src/InvoiceApi.ps1', 'src/XmlParser.ps1',
    'src/ExcelExporter.ps1', 'src/Login.ps1', 'src/BrowserProfile.ps1', 'src/XmlScheduler.ps1', 'src/InvoiceDetail.ps1',
    'src/LinkTraCuu.ps1', 'src/Desktop.ps1', 'src/DesktopSettings.ps1', 'ui/MainWindow.xaml', 'ui/AdvancedSettings.xaml')
foreach ($part in $parts) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $part) -Destination (Join-Path $OutputDirectory $part)
}
$exe = Join-Path $OutputDirectory 'HDDT-Downloader.exe'
& $compiler /nologo /target:winexe /platform:anycpu /reference:System.Windows.Forms.dll "/out:$exe" (Join-Path $PSScriptRoot 'ui/Launcher.cs')
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw 'Build launcher .exe that bai.' }

$note = @'
HDDT Downloader - Windows 10 / Windows 11

Giai nen toan bo thu muc, sau do nhan dup HDDT-Downloader.exe.
Giu file .exe cung cac thu muc src/ va ui/; day la app portable.
Khong can cai Excel de tao workbook .xlsx.
Khong can cai PowerShell 7, Node.js hay thu vien ben ngoai.

Tai tu GDT: nhap tai khoan, mat khau, khoang ngay va chon noi luu.
XML co san: chon thu muc XML va noi luu Excel; khong can Internet.
Dung an toan: cho request hien tai, app xuat phan du lieu da xu ly.
Dong cua so khi dang chay cung yeu cau dung an toan truoc khi thoat.
App khong ghi de Excel da co; hay chon ten moi cho moi lan chay.
Mac dinh ket qua nam trong thu muc output/ ben canh app.
Dat app trong thu muc ban co quyen ghi, vi du Documents.

Ban co the nap .env cua minh bang nut Nap cau hinh .env.
Cai dat nang cao: so ket noi XML, gian cach, phuc hoi 429, retry, timeout.
Ghi nho cai dat chi luu tham so tai/nhat ky vao hddt-settings.json canh app.
Goi phat hanh khong kem tai khoan, mat khau hoac du lieu hoa don.
'@
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'HUONG-DAN.txt'), $note, (New-Object Text.UTF8Encoding($false)))
$zipPath = $OutputDirectory + '.zip'
if (Test-Path -LiteralPath $zipPath) { throw 'File ZIP da ton tai; khong ghi de.' }
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
# Duong dan entry ZIP dung / ca tren .NET Framework cu cua Windows.
$zipStream = [IO.File]::Open($zipPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
$archive = $null
try {
    $archive = New-Object IO.Compression.ZipArchive($zipStream, [IO.Compression.ZipArchiveMode]::Create, $true)
    foreach ($part in @($parts + @('HDDT-Downloader.exe', 'HUONG-DAN.txt'))) {
        $entry = $archive.CreateEntry(($part -replace '\\', '/'), [IO.Compression.CompressionLevel]::Optimal)
        $source = [IO.File]::OpenRead((Join-Path $OutputDirectory $part))
        $target = $entry.Open()
        try { $source.CopyTo($target) }
        finally { $target.Dispose(); $source.Dispose() }
    }
}
finally {
    if ($null -ne $archive) { $archive.Dispose() }
    $zipStream.Dispose()
}
Write-Output ('App: ' + $exe)
Write-Output ('ZIP: ' + $zipPath)
