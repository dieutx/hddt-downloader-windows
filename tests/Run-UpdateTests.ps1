Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'src/DesktopUpdates.ps1')

function Assert-Update {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAILED: ' + $Message) }
}

function Assert-UpdateThrows {
    param([scriptblock]$Action, [string]$Message)
    $threw = $false
    try { & $Action } catch { $threw = $true }
    Assert-Update $threw $Message
}

function New-TestRelease {
    return [pscustomobject]@{
        tag_name = 'v1.2.0'; html_url = 'https://github.com/dieutx/hddt-downloader-windows/releases/tag/v1.2.0'
        draft = $false; prerelease = $false; body = 'Synthetic release notes'; published_at = '2026-01-01T00:00:00Z'
        assets = @(
            [pscustomobject]@{ name = 'source.zip'; browser_download_url = 'https://github.com/dieutx/hddt-downloader-windows/releases/download/v1.2.0/source.zip'; size = 20 }
            [pscustomobject]@{ name = 'HDDT-Downloader-v1.2.0-windows.zip'; browser_download_url = 'https://github.com/dieutx/hddt-downloader-windows/releases/download/v1.2.0/HDDT-Downloader-v1.2.0-windows.zip'; size = 100 }
        )
    }
}

Assert-Update ((Compare-HddtReleaseVersion 'v1.10.0' '1.9.9') -gt 0) 'Version numbers compare numerically'
Assert-Update ((Compare-HddtReleaseVersion '1.2.0' '1.2.0-rc.9') -gt 0) 'Stable release follows release candidate'
Assert-Update ((Compare-HddtReleaseVersion '1.2.0-rc.10' '1.2.0-rc.9') -gt 0) 'Numeric prerelease suffixes compare numerically'
Assert-Update ((Compare-HddtReleaseVersion '1.2.0+build2' 'v1.2.0+build1') -eq 0) 'Build metadata does not change precedence'
Assert-UpdateThrows { ConvertTo-HddtReleaseVersion 'latest' } 'Invalid version is rejected'
$release = New-TestRelease
$info = ConvertFrom-HddtRelease $release '1.1.0'
Assert-Update $info.IsNewer 'New version is detected'
Assert-Update ($info.AssetName -eq 'HDDT-Downloader-v1.2.0-windows.zip') 'Select Windows portable asset, not source archive'
Assert-Update ($info.Notes -eq $release.body) 'Release notes are available to the UI'
$exeAsset = [pscustomobject]@{ name = 'HDDT-Downloader-v1.2.0-windows.exe'; browser_download_url = 'https://github.com/dieutx/hddt-downloader-windows/releases/download/v1.2.0/HDDT-Downloader-v1.2.0-windows.exe'; size = 100; digest = ('sha256:' + ('a' * 64)) }
$release.assets += $exeAsset
$exeInfo = ConvertFrom-HddtRelease $release '1.1.1'
Assert-Update ($exeInfo.PackageKind -eq 'exe') 'Prefer single EXE over folder ZIP'
Assert-Update ($exeInfo.AssetName -eq $exeAsset.name) 'Select official single EXE asset'
Assert-Update ($exeInfo.Digest -eq $exeAsset.digest) 'Retain API checksum for download validation'
$release.assets = @($exeAsset)
Assert-Update ((ConvertFrom-HddtRelease $release '1.1.1').PackageKind -eq 'exe') 'Release with EXE only is supported'
Assert-Update (-not (ConvertFrom-HddtRelease $release '1.2.0').IsNewer) 'Same release does not prompt an update'
Assert-Update (-not (ConvertFrom-HddtRelease $release '2.0.0').IsNewer) 'Local version newer than public release does not downgrade'
$release.assets = @()
Assert-Update ([string]::IsNullOrEmpty((ConvertFrom-HddtRelease $release '1.1.0').DownloadUrl)) 'Release without Windows asset falls back to release page'
$release = New-TestRelease
$release.assets[1].browser_download_url = 'https://example.com/app.zip'
Assert-Update ([string]::IsNullOrEmpty((ConvertFrom-HddtRelease $release '1.1.0').DownloadUrl)) 'External asset URL is never used'
$release = New-TestRelease
$release.assets[1].name = 'HDDT-Downloader-v1.1.0-windows.zip'
Assert-Update ([string]::IsNullOrEmpty((ConvertFrom-HddtRelease $release '1.1.0').DownloadUrl)) 'Mismatched Windows asset version is ignored'
$release = New-TestRelease
$release.draft = $true
Assert-UpdateThrows { ConvertFrom-HddtRelease $release '1.1.0' } 'Draft release is rejected'
$release.draft = $false; $release.prerelease = $true
Assert-UpdateThrows { ConvertFrom-HddtRelease $release '1.1.0' } 'Prerelease is rejected'
$release = New-TestRelease
$release.html_url = 'https://github.com/other/project/releases/tag/v1.2.0'
Assert-UpdateThrows { ConvertFrom-HddtRelease $release '1.1.0' } 'Unrelated release URL is rejected'
foreach ($url in @('http://github.com/dieutx/hddt-downloader-windows/releases/latest', 'https://github.com.evil.example/dieutx/hddt-downloader-windows/releases/latest', 'https://user@github.com/dieutx/hddt-downloader-windows/releases/latest', 'file:///C:/test.zip')) {
    Assert-Update (-not (Test-HddtReleaseUrl $url)) 'Only official HTTPS release URLs are accepted'
}
$release = New-TestRelease
$release.assets[1].size = 0
Assert-UpdateThrows { ConvertFrom-HddtRelease $release '1.1.0' } 'Empty download is rejected'
Assert-UpdateThrows { Save-HddtReleasePackage -Release ([pscustomobject]@{ DownloadUrl = 'https://example.com/app.zip' }) -Destination 'unused.zip' } 'Downloader validates URL before accessing network'
$cancelledPath = Join-Path ([IO.Path]::GetTempPath()) ('hddt-cancel-' + [guid]::NewGuid().ToString('N') + '.zip')
$context = @{ CancelRequested = $true }
$info = ConvertFrom-HddtRelease (New-TestRelease) '1.1.0'
Assert-UpdateThrows { Save-HddtReleasePackage -Release $info -Destination $cancelledPath -Context $context } 'Closing before download starts cancels without writing ZIP'
Assert-Update (-not (Test-Path -LiteralPath $cancelledPath)) 'Cancelled download does not leave final file'
Assert-Update ($null -eq $context.Cancellation) 'Cancelled download releases cancellation source'
try {
    function Invoke-RestMethod {
        param($Uri, $Headers, $TimeoutSec, $ErrorAction)
        Assert-Update ($Uri -eq 'https://api.github.com/repos/dieutx/hddt-downloader-windows/releases/latest') 'Update check uses official latest stable endpoint'
        Assert-Update ($TimeoutSec -le 20) 'Update check has a bounded timeout'
        Assert-Update (-not $Headers.ContainsKey('Authorization')) 'Update check never uses invoice login token'
        return (New-TestRelease)
    }
    Assert-Update (Get-HddtLatestRelease '1.1.0').IsNewer 'Check parses the API response'
    function Invoke-RestMethod { throw 'Synthetic offline condition' }
    Assert-UpdateThrows { Get-HddtLatestRelease '1.1.0' } 'Offline error propagates to the nonblocking UI handler'
}
finally { Remove-Item Function:\Invoke-RestMethod }
Write-Output 'All update tests passed.'
