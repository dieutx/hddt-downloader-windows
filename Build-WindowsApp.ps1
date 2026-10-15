[CmdletBinding()]
param([string]$OutputDirectory, [switch]$FolderBundle)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Chi build app WPF tren Windows.' }
. (Join-Path $PSScriptRoot 'src/DesktopUpdates.ps1')
$appVersion = Get-HddtAppVersion $PSScriptRoot
$versionInfo = ConvertTo-HddtReleaseVersion $appVersion
$assemblyVersion = $versionInfo.Number.ToString() + '.0'
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path (Join-Path $PSScriptRoot 'work/desktop') ('HDDT-{0}-{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8)))
}
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Thu muc build da ton tai. Hay chon thu muc moi de khong ghi de du lieu.' }
$compiler = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::Windows)) 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) { throw 'Khong tim thay trinh bien dich .NET Framework cua Windows.' }
# Chi dong goi danh sach ma nguon tuong minh; khong dong goi du lieu ca nhan.
$parts = @('Start-HddtGui.ps1', 'Invoke-Hddt.ps1', 'Parse-LocalXml.ps1', 'README.md', 'CHANGELOG.md',
    'src/Logging.ps1', 'src/Config.ps1', 'src/Http.ps1', 'src/InvoiceApi.ps1', 'src/XmlParser.ps1',
    'src/ExcelExporter.ps1', 'src/Login.ps1', 'src/BrowserProfile.ps1', 'src/XmlScheduler.ps1', 'src/InvoiceDetail.ps1',
    'src/LinkTraCuu.ps1', 'src/Desktop.ps1', 'src/DesktopSettings.ps1', 'src/DesktopUpdates.ps1',
    'app-version.json', 'ui/vi-VN.json', 'ui/MainWindow.xaml', 'ui/AdvancedSettings.xaml')
foreach ($part in $parts) {
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $part) -PathType Leaf)) { throw ('Thieu ma nguon: ' + $part) }
}
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
function Write-HddtBuildArchive {
    param([string]$Path, [string]$Root, [string[]]$Parts)
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $archive = $null
    try {
        $archive = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Create, $true)
        foreach ($part in $Parts) {
            $entry = $archive.CreateEntry(($part -replace '\\', '/'), [IO.Compression.CompressionLevel]::Optimal)
            $source = [IO.File]::OpenRead((Join-Path $Root $part))
            $target = $entry.Open()
            try { $source.CopyTo($target) } finally { $target.Dispose(); $source.Dispose() }
        }
    }
    finally { if ($null -ne $archive) { $archive.Dispose() }; $stream.Dispose() }
}

$null = New-Item -ItemType Directory -Path $OutputDirectory
$workDirectory = Join-Path ([IO.Path]::GetTempPath()) ('hddt-build-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $workDirectory
try {
    $assemblyInfo = Join-Path $workDirectory 'AssemblyInfo.cs'
    [IO.File]::WriteAllText($assemblyInfo, ('using System.Reflection; [assembly: AssemblyVersion("{0}")] [assembly: AssemblyFileVersion("{0}")] [assembly: AssemblyInformationalVersion("{1}")]' -f $assemblyVersion, $appVersion), (New-Object Text.UTF8Encoding($false)))
    $name = 'HDDT-Downloader-v{0}-windows.exe' -f $appVersion
    if ($FolderBundle) { $name = 'HDDT-Downloader.exe' }
    $exe = Join-Path $OutputDirectory $name
    $arguments = @('/nologo', '/target:winexe', '/platform:anycpu', '/reference:System.Windows.Forms.dll',
        '/reference:System.IO.Compression.dll', '/reference:System.IO.Compression.FileSystem.dll', ('/out:' + $exe))
    if (-not $FolderBundle) {
        $payload = Join-Path $workDirectory 'payload.zip'
        Write-HddtBuildArchive -Path $payload -Root $PSScriptRoot -Parts $parts
        $arguments += ('/resource:{0},Hddt.Payload.zip' -f $payload)
    }
    $arguments += @((Join-Path $PSScriptRoot 'ui/Launcher.cs'), $assemblyInfo)
    & $compiler @arguments
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw 'Build EXE that bai.' }
    if ($FolderBundle) {
        foreach ($part in $parts) {
            $destination = Join-Path $OutputDirectory $part
            $null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
            Copy-Item -LiteralPath (Join-Path $PSScriptRoot $part) -Destination $destination
        }
        $zipPath = $OutputDirectory + '.zip'
        Write-HddtBuildArchive -Path $zipPath -Root $OutputDirectory -Parts @($parts + @('HDDT-Downloader.exe'))
        Write-Output ('ZIP: ' + $zipPath)
    }
    Write-Output ('App: ' + $exe)
    Write-Output ('Single file: ' + (-not $FolderBundle))
}
finally {
    # Thu muc nay do build hien tai tao, nam truc tiep trong TEMP.
    $parent = [IO.Path]::GetFullPath([IO.Path]::GetDirectoryName($workDirectory)).TrimEnd([IO.Path]::DirectorySeparatorChar)
    $expected = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if ($parent -eq $expected -and [IO.Path]::GetFileName($workDirectory).StartsWith('hddt-build-')) {
        Remove-Item -LiteralPath $workDirectory -Recurse -Force
    }
}
