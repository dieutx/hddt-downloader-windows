Set-StrictMode -Version 2.0

function Get-HddtAppVersion {
    param([Parameter(Mandatory = $true)][string]$Root)
    $manifest = [IO.File]::ReadAllText((Join-Path $Root 'app-version.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    $null = ConvertTo-HddtReleaseVersion ([string]$manifest.version)
    return [string]$manifest.version
}

function ConvertTo-HddtReleaseVersion {
    param([Parameter(Mandatory = $true)][string]$Text)
    if ($Text -notmatch '^v?(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?(?:\+[0-9A-Za-z.-]+)?$') {
        throw 'Phien ban khong dung dinh dang major.minor.patch.'
    }
    $suffix = ''
    if ($Matches.ContainsKey(4)) { $suffix = $Matches[4] }
    return [pscustomobject]@{ Number = [version]('{0}.{1}.{2}' -f $Matches[1], $Matches[2], $Matches[3]); Suffix = $suffix }
}

function Compare-HddtReleaseVersion {
    param([string]$Latest, [string]$Current)
    $left = ConvertTo-HddtReleaseVersion $Latest
    $right = ConvertTo-HddtReleaseVersion $Current
    $comparison = $left.Number.CompareTo($right.Number)
    if ($comparison -ne 0) { return $comparison }
    if ($left.Suffix -eq $right.Suffix) { return 0 }
    if ($left.Suffix.Length -eq 0) { return 1 }
    if ($right.Suffix.Length -eq 0) { return -1 }
    $leftParts = $left.Suffix.Split('.')
    $rightParts = $right.Suffix.Split('.')
    for ($i = 0; $i -lt [Math]::Min($leftParts.Length, $rightParts.Length); $i++) {
        $a = $leftParts[$i]; $b = $rightParts[$i]
        if ($a -ceq $b) { continue }
        if ($a -match '^\d+$' -and $b -match '^\d+$') {
            $a = $a.TrimStart('0'); $b = $b.TrimStart('0')
            if ($a.Length -ne $b.Length) { return $a.Length.CompareTo($b.Length) }
        }
        elseif ($a -match '^\d+$') { return -1 }
        elseif ($b -match '^\d+$') { return 1 }
        $comparison = [string]::CompareOrdinal($a, $b)
        if ($comparison -ne 0) { return $comparison }
    }
    return $leftParts.Length.CompareTo($rightParts.Length)
}

function Test-HddtReleaseUrl {
    param([string]$Url, [switch]$Asset)
    $uri = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri)) { return $false }
    if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'github.com' -or -not $uri.IsDefaultPort -or $uri.UserInfo.Length -gt 0) { return $false }
    $prefix = '/dieutx/hddt-downloader-windows/releases/'
    if ($Asset) { $prefix += 'download/' }
    return $uri.AbsolutePath.StartsWith($prefix, [StringComparison]::Ordinal)
}

function ConvertFrom-HddtRelease {
    param([Parameter(Mandatory = $true)]$Release, [string]$CurrentVersion)
    # Draft/prerelease khong duoc thong bao cho kenh on dinh.
    foreach ($property in @('tag_name', 'html_url', 'draft', 'prerelease', 'assets', 'body', 'published_at')) {
        if ($null -eq $Release.PSObject.Properties[$property]) { throw 'Thong tin release thieu truong bat buoc.' }
    }
    if ($Release.draft -or $Release.prerelease) { throw 'Release chua thuoc kenh on dinh.' }
    $version = ConvertTo-HddtReleaseVersion ([string]$Release.tag_name)
    if ($version.Suffix.Length -gt 0) { throw 'Release chua thuoc kenh on dinh.' }
    if (-not (Test-HddtReleaseUrl ([string]$Release.html_url))) { throw 'Dia chi release khong thuoc du an.' }
    $asset = $null
    # Uu tien EXE mot file, van ho tro ZIP portable cua cac ban cu.
    foreach ($candidate in @($Release.assets)) {
        if ($null -eq $candidate) { continue }
        if ([string]$candidate.name -notmatch '^HDDT-Downloader-v?(\d+\.\d+\.\d+)-windows\.(exe|zip)$') { continue }
        $assetVersion = $Matches[1]
        $kind = $Matches[2].ToLowerInvariant()
        if ((Compare-HddtReleaseVersion $assetVersion ([string]$Release.tag_name)) -ne 0) { continue }
        if (-not (Test-HddtReleaseUrl ([string]$candidate.browser_download_url) -Asset)) { continue }
        if ($null -eq $asset -or $kind -eq 'exe') { $asset = $candidate }
        if ($kind -eq 'exe') { break }
    }
    $downloadUrl = ''; $assetName = ''; $size = 0L; $packageKind = ''; $digest = ''
    if ($null -ne $asset) {
        $downloadUrl = [string]$asset.browser_download_url; $assetName = [string]$asset.name; $size = [long]$asset.size
        $packageKind = [IO.Path]::GetExtension($assetName).TrimStart('.').ToLowerInvariant()
        if ($null -ne $asset.PSObject.Properties['digest']) { $digest = [string]$asset.digest }
    }
    if ($null -ne $asset -and $size -le 0) { throw 'Goi cap nhat khong hop le.' }
    $notes = [string]$Release.body
    if ($notes.Length -gt 12000) { $notes = $notes.Substring(0, 12000) + '...' }
    return [pscustomobject]@{
        Version = ([string]$Release.tag_name -replace '^v', ''); IsNewer = ((Compare-HddtReleaseVersion ([string]$Release.tag_name) $CurrentVersion) -gt 0)
        ReleaseUrl = [string]$Release.html_url; Notes = $notes; PublishedAt = [string]$Release.published_at
        DownloadUrl = $downloadUrl; AssetName = $assetName; Size = $size; PackageKind = $packageKind; Digest = $digest
    }
}

function Get-HddtLatestRelease {
    param([string]$CurrentVersion)
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $release = Invoke-RestMethod -Uri 'https://api.github.com/repos/dieutx/hddt-downloader-windows/releases/latest' -Headers @{
        'User-Agent' = ('HDDT-Downloader/' + $CurrentVersion); 'Accept' = 'application/vnd.github+json'; 'X-GitHub-Api-Version' = '2022-11-28'
    } -TimeoutSec 15 -ErrorAction Stop
    return (ConvertFrom-HddtRelease -Release $release -CurrentVersion $CurrentVersion)
}

function Assert-HddtReleasePackage {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)]$Release)
    if ((Get-Item -LiteralPath $Path).Length -ne $Release.Size) { throw 'Dung luong file khong khop release.' }
    if ($null -ne $Release.PSObject.Properties['Digest'] -and -not [string]::IsNullOrEmpty($Release.Digest)) {
        if ($Release.Digest -notmatch '^sha256:([a-fA-F0-9]{64})$') { throw 'Checksum release khong hop le.' }
        $expectedHash = $Matches[1]
        if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -ne $expectedHash) { throw 'Checksum file khong khop release.' }
    }
    $extension = [IO.Path]::GetExtension($Release.AssetName).ToLowerInvariant()
    if ($extension -eq '.exe') {
        $stream = [IO.File]::OpenRead($Path)
        $reader = New-Object IO.BinaryReader($stream)
        try {
            if ($stream.Length -lt 128 -or $reader.ReadUInt16() -ne 0x5A4D) { throw 'File tai khong phai EXE Windows.' }
            $stream.Position = 60
            $offset = $reader.ReadInt32()
            if ($offset -lt 64 -or $offset -gt $stream.Length - 4) { throw 'Cau truc EXE khong hop le.' }
            $stream.Position = $offset
            if ($reader.ReadUInt32() -ne 0x00004550) { throw 'Cau truc EXE khong hop le.' }
        }
        finally { $reader.Dispose(); $stream.Dispose() }
    }
    elseif ($extension -eq '.zip') {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $archive = [IO.Compression.ZipFile]::OpenRead($Path)
        try {
            foreach ($part in @('HDDT-Downloader.exe', 'Start-HddtGui.ps1', 'ui/MainWindow.xaml')) {
                if ($null -eq $archive.GetEntry($part)) { throw 'ZIP khong phai goi app Windows portable.' }
            }
        }
        finally { $archive.Dispose() }
    }
    else { throw 'Chi ho tro cap nhat EXE hoac ZIP Windows.' }
}

function Save-HddtReleasePackage {
    param([Parameter(Mandatory = $true)]$Release, [Parameter(Mandatory = $true)][string]$Destination, [hashtable]$Context = @{})
    if (-not (Test-HddtReleaseUrl $Release.DownloadUrl -Asset)) { throw 'Dia chi tai cap nhat khong hop le.' }
    if (Test-Path -LiteralPath $Destination) { throw 'File da ton tai. Hay chon ten moi.' }
    $temporary = $Destination + '.' + [guid]::NewGuid().ToString('N') + '.partial'
    $client = $null; $response = $null; $source = $null; $target = $null
    $cancellation = New-Object Threading.CancellationTokenSource
    try {
        $Context['Cancellation'] = $cancellation
        if ($Context.ContainsKey('CancelRequested') -and $Context.CancelRequested) { $cancellation.Cancel() }
        Add-Type -AssemblyName System.Net.Http
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $client = New-Object Net.Http.HttpClient
        $client.Timeout = [TimeSpan]::FromSeconds(30)
        $client.DefaultRequestHeaders.UserAgent.ParseAdd('HDDT-Downloader')
        $response = $client.GetAsync($Release.DownloadUrl, [Net.Http.HttpCompletionOption]::ResponseHeadersRead, $cancellation.Token).GetAwaiter().GetResult()
        $null = $response.EnsureSuccessStatusCode()
        $source = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $target = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $buffer = New-Object byte[] 65536
        $received = 0L
        while ($true) {
            $cancellation.CancelAfter(30000)
            $count = $source.ReadAsync($buffer, 0, $buffer.Length, $cancellation.Token).GetAwaiter().GetResult()
            if ($count -eq 0) { break }
            $target.Write($buffer, 0, $count)
            $received += $count
            $Context['Received'] = $received
        }
        $target.Dispose(); $target = $null
        if ($received -ne $Release.Size) { throw 'Dung luong goi tai khong khop release. Hay thu lai.' }
        Assert-HddtReleasePackage -Path $temporary -Release $Release
        # Move khong ghi de neu file dich xuat hien trong luc dang tai.
        [IO.File]::Move($temporary, $Destination)
        return $Destination
    }
    finally {
        $Context['Cancellation'] = $null
        if ($null -ne $target) { $target.Dispose() }
        if ($null -ne $source) { $source.Dispose() }
        if ($null -ne $response) { $response.Dispose() }
        if ($null -ne $client) { $client.Dispose() }
        $cancellation.Dispose()
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Start-HddtUpdateTask {
    param([string]$Root, [string]$CurrentVersion, [ValidateSet('check', 'download')][string]$Mode = 'check', $Release = $null, [string]$Destination = '')
    $context = [hashtable]::Synchronized(@{ Received = 0L; Cancellation = $null; CancelRequested = $false })
    $pipeline = [Management.Automation.PowerShell]::Create()
    $worker = @'
param($Module, $Version, $Mode, $Release, $Destination, $Context)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
. $Module
if ($Mode -eq 'check') { Get-HddtLatestRelease -CurrentVersion $Version }
else { Save-HddtReleasePackage -Release $Release -Destination $Destination -Context $Context }
'@
    try {
        $null = $pipeline.AddScript($worker).AddArgument((Join-Path $Root 'src/DesktopUpdates.ps1')).AddArgument($CurrentVersion).AddArgument($Mode).AddArgument($Release).AddArgument($Destination).AddArgument($context)
        return [pscustomobject]@{ Pipeline = $pipeline; Handle = $pipeline.BeginInvoke(); Mode = $Mode; Context = $context; Destination = $Destination }
    }
    catch { $pipeline.Dispose(); throw }
}

function Complete-HddtUpdateTask {
    param([Parameter(Mandatory = $true)]$Task)
    if (-not $Task.Handle.IsCompleted) { throw 'Tac vu cap nhat chua ket thuc.' }
    try {
        $results = @($Task.Pipeline.EndInvoke($Task.Handle))
        if ($Task.Pipeline.HadErrors) { throw 'Khong hoan tat duoc tac vu cap nhat.' }
        if ($results.Count -ne 1) { throw 'Ket qua cap nhat khong hop le.' }
        return $results[0]
    }
    finally { $Task.Pipeline.Dispose() }
}
