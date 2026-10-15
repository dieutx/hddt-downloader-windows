param([string]$SingleFileExePath, [string]$LayoutPreviewDirectory)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Kiem thu WPF can Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Hay chay kiem thu bang powershell.exe -STA.' }
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'Start-HddtGui.ps1') -SmokeTest

function Assert-Desktop {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAILED: ' + $Message) }
}

function Wait-DesktopWindowTask {
    $deadline = [datetime]::UtcNow.AddSeconds(30)
    while ($null -ne $script:GuiRun -and [datetime]::UtcNow -lt $deadline) {
        $frame = New-Object Windows.Threading.DispatcherFrame
        $null = [Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
            [Windows.Threading.DispatcherPriority]::Background, [Action]{ $frame.Continue = $false })
        [Windows.Threading.Dispatcher]::PushFrame($frame)
        Start-Sleep -Milliseconds 20
    }
    Assert-Desktop ($null -eq $script:GuiRun) 'UI task completes and timer releases its runspace'
}

$testDirectory = Join-Path ([IO.Path]::GetTempPath()) ('hddt-wpf-' + [guid]::NewGuid().ToString('N'))
try {
    $script:GuiControls.Username.Text = 'fake-account'
    $script:GuiControls.Password.Password = 'synthetic-ui-password'
    $script:GuiControls.OutputPath.Text = Join-Path $testDirectory 'download.xlsx'
    $guiValues = Get-HddtGuiValues
    Assert-Desktop ($guiValues.GDT_PASSWORD -eq 'synthetic-ui-password') 'WPF passes password unchanged in memory'
    Assert-Desktop ($guiValues.OVERWRITE_OUTPUT -eq 'false') 'WPF prevents overwriting existing output'
    Assert-Desktop ($guiValues.FROM_DATE -match '^\d{2}/\d{2}/\d{4}$') 'DatePicker uses the CLI date contract'
    $guiValues.Clear()

    # Mat khau mac dinh an; nut con mat chi dong bo trong bo nho.
    Assert-Desktop (-not $script:GuiControls.TogglePassword.IsChecked) 'Password is masked by default'
    $script:GuiControls.TogglePassword.IsChecked = $true
    Assert-Desktop ($script:GuiControls.PasswordVisible.Text -eq 'synthetic-ui-password') 'Eye button reveals current password'
    Assert-Desktop ($script:GuiControls.PasswordVisible.Visibility -eq 'Visible') 'Eye button displays editable password text'
    $script:GuiControls.PasswordVisible.Text = 'synthetic-edited-password'
    $guiValues = Get-HddtGuiValues
    Assert-Desktop ($guiValues.GDT_PASSWORD -eq 'synthetic-edited-password') 'Visible password edits reach login values unchanged'
    $guiValues.Clear()
    $script:GuiControls.TogglePassword.IsChecked = $false
    Assert-Desktop ($script:GuiControls.Password.Password -eq 'synthetic-edited-password') 'Hiding password preserves edited value'
    Assert-Desktop ($script:GuiControls.PasswordVisible.Text.Length -eq 0) 'Hiding password clears plain text control'
    $script:GuiControls.Password.Password = 'synthetic-ui-password'
    $script:GuiControls.TogglePassword.IsChecked = $true
    Set-HddtGuiBusy $true
    Assert-Desktop (-not $script:GuiControls.TogglePassword.IsChecked) 'Starting a run restores masked password'
    Set-HddtGuiBusy $false

    # Cai dat nang cao: control that, validation, huy, reset va luu an toan.
    $null = New-Item -ItemType Directory -Path $testDirectory
    # Ban cap nhat phai hien thong tin va duoc xu ly ca khi khong tai hoa don.
    $testRelease = [pscustomobject]@{
        Version = '1.2.0'; IsNewer = $true; ReleaseUrl = 'https://github.com/dieutx/hddt-downloader-windows/releases/tag/v1.2.0'
        Notes = 'Synthetic release notes'; PublishedAt = '2026-01-01T00:00:00Z'
        DownloadUrl = 'https://github.com/dieutx/hddt-downloader-windows/releases/download/v1.2.0/HDDT-Downloader-v1.2.0-windows.zip'
        AssetName = 'HDDT-Downloader-v1.2.0-windows.zip'; Size = 100L
    }
    Update-HddtGuiReleaseState $testRelease
    Assert-Desktop ($script:GuiControls.UpdateTitle.Text.Contains('1.2.0')) 'Update banner shows new version'
    Assert-Desktop ($script:GuiControls.DownloadUpdate.Visibility -eq 'Visible') 'New stable Windows release enables download'
    Assert-Desktop ($script:GuiLatestRelease.Notes -eq 'Synthetic release notes') 'UI shows release notes as plain text'
    $testRelease.DownloadUrl = ''
    Update-HddtGuiReleaseState $testRelease
    Assert-Desktop ($script:GuiControls.DownloadUpdate.Visibility -eq 'Collapsed') 'No Windows asset hides download button'
    $testRelease.IsNewer = $false
    Update-HddtGuiReleaseState $testRelease
    Assert-Desktop ($script:GuiControls.DownloadUpdate.Visibility -eq 'Collapsed') 'Same/older public version never offers downgrade'

    $fakeRoot = Join-Path $testDirectory 'fake-update-root'
    $null = New-Item -ItemType Directory -Path (Join-Path $fakeRoot 'src') -Force
    $fakeModule = Join-Path $fakeRoot 'src/DesktopUpdates.ps1'
    [IO.File]::WriteAllText($fakeModule, 'function Get-HddtLatestRelease { param($CurrentVersion); Start-Sleep -Milliseconds 250; throw "Synthetic offline condition" }')
    $script:GuiControls.CheckUpdates.IsEnabled = $false
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $script:GuiUpdateTask = Start-HddtUpdateTask -Root $fakeRoot -CurrentVersion '1.1.0'
    Assert-Desktop ($watch.ElapsedMilliseconds -lt 1000) 'Starting update check does not block UI'
    $deadline = [datetime]::UtcNow.AddSeconds(5)
    while ($null -ne $script:GuiUpdateTask -and [datetime]::UtcNow -lt $deadline) {
        Receive-HddtGuiUpdateTask
        Start-Sleep -Milliseconds 20
    }
    Assert-Desktop ($null -eq $script:GuiUpdateTask) 'Offline update releases its pipeline'
    Assert-Desktop $script:GuiControls.CheckUpdates.IsEnabled 'Offline update allows retry'
    Assert-Desktop $script:GuiControls.StartRun.IsEnabled 'Offline update does not disable invoice actions'

    $script:GuiStartedAt = (Get-Date).AddSeconds(-20)
    Add-HddtGuiLog '[2/10 | 20%] synthetic progress'
    Assert-Desktop ($script:GuiControls.ProcessedCount.Text -eq '2' -and $script:GuiControls.RemainingCount.Text -eq '8') 'Metrics display progress counts'
    Assert-Desktop ($script:GuiControls.PercentText.Text -eq '20%') 'Metrics display percent'
    Assert-Desktop ($script:GuiControls.EstimatedTime.Text -ne '-') 'Metrics estimate time after progress'
    $script:GuiControls.ActivityLog.SelectedIndex = 0
    Assert-Desktop ($script:GuiControls.LogDetail.Text.Contains('synthetic progress')) 'Selecting a row shows full log text'
    Add-HddtGuiLog (Protect-HddtDesktopText 'synthetic-ui-secret' @('synthetic-ui-secret'))
    Assert-Desktop (-not $script:GuiActivity[$script:GuiActivity.Count - 1].Message.Contains('synthetic-ui-secret')) 'Activity grid uses redacted log text'
    $script:GuiActivity.Clear()

    $script:GuiSettingsPath = Join-Path $testDirectory 'hddt-settings.json'
    $script:GuiDefaults['XML_CONCURRENCY'] = '2'
    $script:GuiDefaults['XML_MAX_CONCURRENCY'] = '5'
    $settingsDialog = New-HddtGuiSettingsDialog
    Assert-Desktop ($settingsDialog.Controls.XML_CONCURRENCY.Text -eq '2') 'Advanced dialog shows the loaded config value'
    $settingsDialog.Controls.XML_CONCURRENCY.Text = '6'
    $settingsDialog.Controls.ApplySettings.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    Assert-Desktop (-not $settingsDialog.Applied) 'Advanced dialog rejects initial connections above max'
    Assert-Desktop ($settingsDialog.Controls.SettingsError.Text.Length -gt 0) 'Advanced validation displays an inline error'
    Assert-Desktop ($script:GuiDefaults.XML_CONCURRENCY -eq '2') 'Invalid setting does not mutate the current config'
    Assert-Desktop (-not (Test-Path -LiteralPath $script:GuiSettingsPath)) 'Invalid setting is never persisted'
    $settingsDialog.Controls.XML_CONCURRENCY.Text = '3'
    $settingsDialog.Controls.XML_REQUEST_INTERVAL_MS.Text = '1200'
    $settingsDialog.Controls.HTTP_TIMEOUT_SECONDS.Text = '150'
    $settingsDialog.Controls.ApplySettings.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    Assert-Desktop $settingsDialog.Applied 'Applying valid settings succeeds'
    $savedSettings = Read-HddtDesktopSettings $script:GuiSettingsPath
    Assert-Desktop ($savedSettings.XML_CONCURRENCY -eq '3') 'Saved connections reload correctly'
    Assert-Desktop ($savedSettings.HTTP_TIMEOUT_SECONDS -eq '150') 'Saved timeout reloads correctly'
    $guiValues = Get-HddtGuiValues
    $settingsRunConfig = Get-HddtConfig -EnvFile (Join-Path $testDirectory 'missing.env') -RepositoryRoot $root -Values $guiValues
    Assert-Desktop ($settingsRunConfig.XmlConcurrency -eq 3 -and $settingsRunConfig.XmlMaxConcurrency -eq 5) 'Advanced settings reach the actual run config'
    Assert-Desktop ($settingsRunConfig.XmlRequestIntervalMs -eq 1200 -and $settingsRunConfig.HttpTimeoutSeconds -eq 150) 'Advanced interval and timeout reach the actual run config'
    $guiValues.Clear()
    Assert-Desktop (-not ([IO.File]::ReadAllText($script:GuiSettingsPath).Contains('synthetic-ui-password'))) 'Settings file contains no login password'
    Set-HddtGuiBusy $true
    Assert-Desktop (-not $script:GuiControls.AdvancedSettings.IsEnabled) 'Advanced settings are locked while a task runs'
    Set-HddtGuiBusy $false
    $settingsDialog = New-HddtGuiSettingsDialog
    $settingsDialog.Controls.ResetSettings.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    Assert-Desktop ($settingsDialog.Controls.XML_CONCURRENCY.Text -eq '4' -and $settingsDialog.Controls.XML_REQUEST_INTERVAL_MS.Text -eq '800') 'Reset fills default connections and interval'
    $settingsDialog.Controls.CancelSettings.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    Assert-Desktop ($script:GuiDefaults.XML_CONCURRENCY -eq '3') 'Cancel discards reset and field edits'
    $settingsDialog = New-HddtGuiSettingsDialog
    $settingsDialog.Controls.RememberSettings.IsChecked = $false
    $settingsDialog.Controls.XML_CONCURRENCY.Text = '1'
    $settingsDialog.Controls.ApplySettings.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    Assert-Desktop ($script:GuiDefaults.XML_CONCURRENCY -eq '1') 'Session-only settings apply in memory'
    Assert-Desktop ((Read-HddtDesktopSettings $script:GuiSettingsPath).XML_CONCURRENCY -eq '3') 'Session-only apply preserves persisted settings'

    $script:GuiControls.ModeTabs.SelectedIndex = 1
    $script:GuiControls.XmlDirectory.Text = Join-Path $PSScriptRoot 'fixtures'
    $script:GuiControls.LocalDirection.SelectedIndex = 1
    $script:GuiControls.OutputPath.Text = Join-Path $testDirectory 'local.xlsx'
    $guiValues = Get-HddtGuiValues
    Assert-Desktop (-not $guiValues.ContainsKey('GDT_PASSWORD')) 'Offline UI does not pass GDT credentials'
    $guiValues.Clear()
    $script:GuiTimer.Start()
    $script:GuiControls.StartRun.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    Assert-Desktop (-not $script:GuiControls.StartRun.IsEnabled) 'UI disables Start while running'
    Wait-DesktopWindowTask
    Assert-Desktop (Test-Path -LiteralPath (Join-Path $testDirectory 'local.xlsx')) 'Clicking Start creates an offline workbook'
    Assert-Desktop $script:GuiControls.OpenExcel.IsEnabled 'UI enables Open Excel after completion'
    Assert-Desktop ($script:GuiControls.Log.Text.Length -gt 0) 'UI shows background logs'
    Assert-Desktop ($script:GuiControls.Progress.Value -eq 100) 'UI marks completed progress'

    # Dung ngay truoc khi parse: khong tao workbook rong hoac ghi de ket qua.
    $script:GuiControls.OutputPath.Text = Join-Path $testDirectory 'stopped.xlsx'
    $script:GuiControls.StartRun.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    $script:GuiControls.StopRun.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    Wait-DesktopWindowTask
    Assert-Desktop (-not (Test-Path -LiteralPath (Join-Path $testDirectory 'stopped.xlsx'))) 'Stopping before data does not create an empty workbook'
    Assert-Desktop (-not $script:GuiControls.StopRun.IsEnabled) 'UI resets Stop after completion'
    Assert-Desktop ($script:GuiControls.Status.Text -eq $script:GuiText.ProgressStopped) 'UI reports safe stop'

    # Co du lieu roi moi dung: workbook van chua phan XML da xu ly.
    $manyXmlDirectory = Join-Path $testDirectory 'many-xml'
    $null = New-Item -ItemType Directory -Path $manyXmlDirectory
    for ($i = 1; $i -le 500; $i++) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'fixtures/sample-invoice.xml') -Destination (Join-Path $manyXmlDirectory ('fake-{0}.xml' -f $i))
    }
    $script:GuiControls.XmlDirectory.Text = $manyXmlDirectory
    $script:GuiControls.OutputPath.Text = Join-Path $testDirectory 'partial.xlsx'
    $script:GuiControls.StartRun.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    $deadline = [datetime]::UtcNow.AddSeconds(15)
    while ($script:GuiControls.Log.Text -notmatch '\[\d+/500\s*\|\s*\d+%\]' -and [datetime]::UtcNow -lt $deadline) {
        $frame = New-Object Windows.Threading.DispatcherFrame
        $null = [Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
            [Windows.Threading.DispatcherPriority]::Background, [Action]{ $frame.Continue = $false })
        [Windows.Threading.Dispatcher]::PushFrame($frame)
        Start-Sleep -Milliseconds 20
    }
    Assert-Desktop ($null -ne $script:GuiRun) 'Partial-stop test observes a task while still running'
    Request-HddtGuiStop
    Wait-DesktopWindowTask
    $partialWorkbook = Join-Path $testDirectory 'partial.xlsx'
    Assert-Desktop (Test-Path -LiteralPath $partialWorkbook) 'Stopping after data keeps a workbook'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($partialWorkbook)
    try {
        $stream = New-Object IO.StreamReader($archive.GetEntry('xl/worksheets/sheet1.xml').Open())
        try { [xml]$sheet = $stream.ReadToEnd() } finally { $stream.Dispose() }
        $summaryCount = $sheet.SelectNodes("//*[local-name()='t' and text()='C26TABC']").Count
        Assert-Desktop ($summaryCount -gt 0 -and $summaryCount -lt 500) 'Partial workbook contains processed invoices only'
    }
    finally { $archive.Dispose() }

    # Dong cua so cung doi tac vu dung an toan.
    $script:GuiControls.OutputPath.Text = Join-Path $testDirectory 'closed.xlsx'
    $script:GuiControls.StartRun.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
    $script:GuiWindow.Close()
    Assert-Desktop $script:GuiCloseWhenDone 'Closing requests a safe stop before exiting'
    Wait-DesktopWindowTask
    & (Join-Path $PSScriptRoot 'Run-UpdateTests.ps1')
    if (-not [string]::IsNullOrWhiteSpace($SingleFileExePath)) { & (Join-Path $PSScriptRoot 'Run-SingleFileTests.ps1') -ExePath $SingleFileExePath }
    if (-not [string]::IsNullOrWhiteSpace($LayoutPreviewDirectory)) { & (Join-Path $PSScriptRoot 'Run-LayoutTests.ps1') -PreviewDirectory $LayoutPreviewDirectory }
    Write-Host 'All desktop tests passed.'
}
finally {
    $script:GuiTimer.Stop()
    if ($null -ne $script:GuiRun) {
        Request-HddtDesktopStop $script:GuiRun
        $script:GuiRun.Pipeline.Stop()
        $script:GuiRun.Pipeline.Dispose()
        $script:GuiRun.Runspace.Dispose()
    }
    $script:GuiControls.Password.Clear()
    $script:GuiControls.ProxyPassword.Clear()
    $resolvedTestDirectory = [IO.Path]::GetFullPath($testDirectory)
    $allowedPrefix = Join-Path ([IO.Path]::GetFullPath([IO.Path]::GetTempPath())) 'hddt-wpf-'
    if (-not $resolvedTestDirectory.StartsWith($allowedPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Thu muc kiem thu nam ngoai thu muc tam.' }
    if (Test-Path -LiteralPath $resolvedTestDirectory) { Remove-Item -LiteralPath $resolvedTestDirectory -Recurse -Force }
}
