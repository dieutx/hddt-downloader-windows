[CmdletBinding()]
param([switch]$SmokeTest, [string]$PreviewPath, [int]$PreviewWidth = 1160, [int]$PreviewHeight = 736,
    [string]$DataDirectory, [string]$DefaultOutputDirectory, [string]$SmokeReportPath, [string]$SmokeXmlDirectory)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Giao dien WPF chi chay tren Windows 10/11.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Hay chay bang powershell.exe -STA -File Start-HddtGui.ps1.' }

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
. (Join-Path $PSScriptRoot 'src/Config.ps1')
. (Join-Path $PSScriptRoot 'src/ExcelExporter.ps1')
. (Join-Path $PSScriptRoot 'src/Desktop.ps1')
. (Join-Path $PSScriptRoot 'src/DesktopSettings.ps1')
. (Join-Path $PSScriptRoot 'src/DesktopUpdates.ps1')

$script:GuiRoot = $PSScriptRoot
$script:GuiDataRoot = $PSScriptRoot
if (-not [string]::IsNullOrWhiteSpace($DataDirectory)) {
    $script:GuiDataRoot = [IO.Path]::GetFullPath($DataDirectory)
    $null = [IO.Directory]::CreateDirectory($script:GuiDataRoot)
}
$script:GuiOutputRoot = Join-Path $script:GuiDataRoot 'output'
if (-not [string]::IsNullOrWhiteSpace($DefaultOutputDirectory)) { $script:GuiOutputRoot = [IO.Path]::GetFullPath($DefaultOutputDirectory) }
$script:GuiDefaults = @{}
$script:GuiRun = $null
$script:GuiLastWorkbook = ''
$script:GuiCloseWhenDone = $false
$script:GuiControls = @{}
$script:GuiSettingsPath = Join-Path $script:GuiDataRoot 'hddt-settings.json'
$script:GuiSettingsDialog = $null
$script:GuiUpdateTask = $null
$script:GuiLatestRelease = $null
$script:GuiUpdateFile = ''
$script:GuiCloseAfterUpdate = $false
$script:GuiVersion = Get-HddtAppVersion $PSScriptRoot
$script:GuiText = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ui/vi-VN.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
$script:GuiStartedAt = $null
$script:GuiProgress = $null
$script:GuiActivity = New-Object 'Collections.ObjectModel.ObservableCollection[object]'
$script:GuiRecentActivity = New-Object 'Collections.ObjectModel.ObservableCollection[object]'
$script:GuiRecentLimit = 6
$script:GuiLogHistory = New-Object Text.StringBuilder
$script:GuiLogDialog = $null
$script:GuiUpdateLastError = ''
$script:GuiPasswordSync = $false
$reader = New-Object IO.StringReader([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ui/MainWindow.xaml'), [Text.Encoding]::UTF8))
$xmlReader = [Xml.XmlReader]::Create($reader)
try { $script:GuiWindow = [Windows.Markup.XamlReader]::Load($xmlReader) }
finally { $xmlReader.Dispose(); $reader.Dispose() }
$script:GuiWindow.Height = [Math]::Min(768, [Math]::Max(560, [Windows.SystemParameters]::WorkArea.Height - 16))
$script:GuiWindow.Width = [Math]::Min(1160, [Math]::Max(880, [Windows.SystemParameters]::WorkArea.Width - 16))
foreach ($name in @('LoadConfig', 'AdvancedSettings', 'ModeTabs', 'Username', 'Password', 'FromDate', 'ToDate', 'Direction',
    'IncludeRegular', 'IncludeSco', 'ProxyUrl', 'ProxyUsername', 'ProxyPassword', 'FetchRelated', 'RedownloadXml',
    'BrowseXml', 'XmlDirectory', 'LocalDirection', 'OutputPath', 'BrowseOutput', 'StartRun', 'StopRun',
    'OpenExcel', 'OpenFolder', 'Status', 'Progress', 'Log', 'ActivityLog', 'LogDetail',
    'PercentText', 'ProcessedCount', 'RemainingCount', 'ElapsedTime', 'EstimatedTime',
    'CurrentVersion', 'CheckUpdates', 'ReleasePage', 'UpdateTitle', 'UpdateStatus', 'DownloadUpdate', 'ShowUpdateFile',
    'UpdateBanner', 'ViewReleaseNotes', 'DismissUpdate', 'ProxyOptions', 'ProxyPopup', 'ShowFullLog', 'AppState', 'StateDot', 'WorkspaceRoot', 'ActivityTab', 'TechnicalTab', 'LogDetailRow', 'UpdateIndicator',
    'PasswordVisible', 'TogglePassword', 'PasswordEyeSlash')) {
    $control = $script:GuiWindow.FindName($name)
    if ($null -eq $control) { throw ('Thieu control: {0}' -f $name) }
    $script:GuiControls[$name] = $control
}
$script:GuiControls.CurrentVersion.Text = 'v' + $script:GuiVersion
$script:GuiControls.ActivityLog.ItemsSource = $script:GuiRecentActivity
$script:GuiControls.FromDate.SelectedDate = (Get-Date).Date.AddDays(1 - (Get-Date).Day)
$script:GuiControls.ToDate.SelectedDate = (Get-Date).Date
$script:GuiControls.OutputPath.Text = Join-Path $script:GuiOutputRoot ('HoaDonDienTu_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Set-HddtGuiPasswordVisibility {
    param([bool]$Visible)
    $c = $script:GuiControls
    $script:GuiPasswordSync = $true
    try {
        if ($Visible) {
            $c.PasswordVisible.Text = $c.Password.Password
            $c.Password.Visibility = 'Collapsed'
            $c.PasswordVisible.Visibility = 'Visible'
            $c.PasswordEyeSlash.Visibility = 'Visible'
            $c.TogglePassword.ToolTip = $script:GuiText.HidePassword
        }
        else {
            $c.Password.Visibility = 'Visible'
            $c.PasswordVisible.Visibility = 'Collapsed'
            $c.PasswordVisible.Clear()
            $c.PasswordEyeSlash.Visibility = 'Collapsed'
            $c.TogglePassword.ToolTip = $script:GuiText.ShowPassword
        }
    }
    finally { $script:GuiPasswordSync = $false }
}

function Set-HddtGuiState {
    param([ValidateSet('idle', 'connecting', 'downloading', 'exporting', 'completed', 'error', 'stopping')][string]$State)
    $colors = @{ idle = '#64748B'; connecting = '#2563EB'; downloading = '#2563EB'; exporting = '#7C3AED'; completed = '#17834B'; error = '#C43535'; stopping = '#B26A13' }
    $script:GuiControls.AppState.Text = $script:GuiText.PSObject.Properties['State_' + $State].Value
    $script:GuiControls.StateDot.Fill = (New-Object Windows.Media.BrushConverter).ConvertFromString($colors[$State])
}

function New-HddtGuiTextDialog {
    param([string]$Title, [string]$Text, [switch]$Technical, [switch]$OfferDownload)
    $window = New-Object Windows.Window
    $window.Title = $Title
    if ($script:GuiWindow.IsVisible) { $window.Owner = $script:GuiWindow }
    $window.WindowStartupLocation = 'CenterOwner'
    $window.Width = [Math]::Min(960, [Windows.SystemParameters]::WorkArea.Width - 32)
    $window.Height = [Math]::Min(620, [Windows.SystemParameters]::WorkArea.Height - 32)
    $window.MinWidth = 560; $window.MinHeight = 320
    $window.FontFamily = New-Object Windows.Media.FontFamily('Segoe UI')
    $window.FontSize = 13
    $window.Resources = $script:GuiWindow.Resources
    $panel = New-Object Windows.Controls.DockPanel
    $panel.Margin = New-Object Windows.Thickness(12)
    $close = New-Object Windows.Controls.Button
    $close.Content = $script:GuiText.CloseDialog
    $close.HorizontalAlignment = 'Right'
    $close.Margin = New-Object Windows.Thickness(0, 8, 0, 0)
    [Windows.Controls.DockPanel]::SetDock($close, 'Bottom')
    $null = $panel.Children.Add($close)
    $textBox = New-Object Windows.Controls.TextBox
    $textBox.IsReadOnly = $true; $textBox.Text = $Text
    $textBox.VerticalScrollBarVisibility = 'Auto'; $textBox.HorizontalScrollBarVisibility = 'Auto'
    $textBox.TextWrapping = 'Wrap'
    if ($Technical) {
        $textBox.FontFamily = New-Object Windows.Media.FontFamily('Consolas')
        $textBox.TextWrapping = 'NoWrap'
        $textBox.Background = [Windows.Media.Brushes]::AliceBlue
    }
    $null = $panel.Children.Add($textBox)
    $window.Content = $panel
    $close.Add_Click({ param($sender, $eventArgs); [Windows.Window]::GetWindow($sender).Close() })
    if ($OfferDownload) {
        $download = New-Object Windows.Controls.Button
        $download.Content = $script:GuiText.DownloadExeButton
        $download.HorizontalAlignment = 'Left'
        $download.Margin = New-Object Windows.Thickness(0, 8, 0, 0)
        [Windows.Controls.DockPanel]::SetDock($download, 'Bottom')
        $panel.Children.Insert(1, $download)
        $download.IsEnabled = ($null -eq $script:GuiUpdateTask)
        $download.Add_Click({
            param($sender, $eventArgs)
            [Windows.Window]::GetWindow($sender).Close()
            $script:GuiControls.DownloadUpdate.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
        })
    }
    return [pscustomobject]@{ Window = $window; Text = $textBox }
}

function Update-HddtGuiRecentLog {
    if ($script:GuiControls.ActivityLog.ActualWidth -gt 420) {
        $width = [Math]::Max(240, $script:GuiControls.ActivityLog.ActualWidth - 182)
        $script:GuiControls.ActivityLog.Columns[2].Width = New-Object Windows.Controls.DataGridLength($width)
    }
    $limit = [Math]::Max(1, [Math]::Min(6, [int][Math]::Floor(($script:GuiControls.ActivityLog.ActualHeight - 25) / 23)))
    if ($limit -gt $script:GuiRecentLimit) {
        $selected = $script:GuiControls.ActivityLog.SelectedItem
        $script:GuiRecentActivity.Clear()
        for ($index = [Math]::Max(0, $script:GuiActivity.Count - $limit); $index -lt $script:GuiActivity.Count; $index++) { $script:GuiRecentActivity.Add($script:GuiActivity[$index]) }
        if ($null -ne $selected -and $script:GuiRecentActivity.Contains($selected)) { $script:GuiControls.ActivityLog.SelectedItem = $selected }
    }
    $script:GuiRecentLimit = $limit
    while ($script:GuiRecentActivity.Count -gt $limit) { $script:GuiRecentActivity.RemoveAt(0) }
}

function Update-HddtGuiLayout {
    $height = $script:GuiControls.WorkspaceRoot.ActualHeight
    if ($height -gt 0 -and $height -lt 640) {
        $script:GuiControls.UpdateBanner.MaxHeight = 0
        $script:GuiControls.UpdateBanner.Margin = New-Object Windows.Thickness(0)
        $script:GuiControls.LogDetailRow.Height = New-Object Windows.GridLength(0)
    }
    else {
        $script:GuiControls.UpdateBanner.MaxHeight = [double]::PositiveInfinity
        $script:GuiControls.UpdateBanner.Margin = New-Object Windows.Thickness(0, 0, 0, 8)
        $script:GuiControls.LogDetailRow.Height = New-Object Windows.GridLength(22)
    }
}

function Save-HddtGuiPreview {
    param([string]$Path, [int]$Width = 1160, [int]$Height = 736, [double]$Dpi = 96)
    $content = $script:GuiWindow.Content
    $script:GuiWindow.Content = $null
    $visual = New-Object Windows.Controls.Border
    $visual.Resources = $script:GuiWindow.Resources
    $visual.Background = $script:GuiWindow.Background
    [Windows.Documents.TextElement]::SetFontFamily($visual, $script:GuiWindow.FontFamily)
    [Windows.Documents.TextElement]::SetFontSize($visual, $script:GuiWindow.FontSize)
    [Windows.Documents.TextElement]::SetForeground($visual, $script:GuiWindow.Foreground)
    $visual.Child = $content
    $visual.Measure((New-Object Windows.Size($Width, $Height)))
    $visual.Arrange((New-Object Windows.Rect(0, 0, $Width, $Height)))
    $visual.UpdateLayout()
    $bitmap = New-Object Windows.Media.Imaging.RenderTargetBitmap([int][Math]::Round($Width * $Dpi / 96), [int][Math]::Round($Height * $Dpi / 96), $Dpi, $Dpi, [Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($visual)
    $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.File]::Create($Path)
    try { $encoder.Save($stream) } finally { $stream.Dispose() }
    $visual.Child = $null
    $script:GuiWindow.Content = $content
}

function Update-HddtGuiMetrics {
    if ($null -eq $script:GuiStartedAt) { return }
    $elapsed = (Get-Date) - $script:GuiStartedAt
    $script:GuiControls.ElapsedTime.Text = ('{0:00}:{1:00}:{2:00}' -f [Math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds)
    if ($null -ne $script:GuiProgress) {
        $p = $script:GuiProgress
        $script:GuiControls.ProcessedCount.Text = [string]$p.Current
        $script:GuiControls.RemainingCount.Text = [string][Math]::Max(0, $p.Total - $p.Current)
        $script:GuiControls.PercentText.Text = ('{0}%' -f $p.Percent)
        if ($p.Current -gt 0) {
            $remaining = [TimeSpan]::FromSeconds([Math]::Max(0, $elapsed.TotalSeconds * ($p.Total - $p.Current) / $p.Current))
            $script:GuiControls.EstimatedTime.Text = ('{0:00}:{1:00}:{2:00}' -f [Math]::Floor($remaining.TotalHours), $remaining.Minutes, $remaining.Seconds)
        }
    }
}

function Open-HddtGuiReleaseUrl {
    param([string]$Url)
    if (-not (Test-HddtReleaseUrl $Url)) { throw 'Dia chi release khong hop le.' }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $Url
    $info.UseShellExecute = $true
    $null = [Diagnostics.Process]::Start($info)
}

function Start-HddtGuiUpdateCheck {
    if ($null -ne $script:GuiUpdateTask) { return }
    try {
        $script:GuiControls.CheckUpdates.IsEnabled = $false
        $script:GuiControls.DownloadUpdate.IsEnabled = $false
        $script:GuiControls.CheckUpdates.Content = $script:GuiText.UpdateCheckingShort
        $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateChecking
        $script:GuiUpdateTask = Start-HddtUpdateTask -Root $script:GuiRoot -CurrentVersion $script:GuiVersion
    }
    catch {
        $script:GuiControls.CheckUpdates.IsEnabled = $true
        $script:GuiControls.CheckUpdates.Content = $script:GuiText.UpdateButton
        $script:GuiControls.DownloadUpdate.IsEnabled = ($null -ne $script:GuiLatestRelease)
        $script:GuiControls.UpdateBanner.Visibility = 'Visible'
        $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateCheckFailed
        $script:GuiUpdateLastError = $script:GuiText.UpdateCheckFailed
        $script:GuiControls.UpdateIndicator.Visibility = 'Visible'
        $script:GuiControls.UpdateIndicator.Content = $script:GuiText.UpdateErrorIndicator
        $script:GuiControls.UpdateIndicator.ToolTip = $script:GuiUpdateLastError
    }
}

function Update-HddtGuiReleaseState {
    param($Release)
    $script:GuiLatestRelease = $Release
    $script:GuiUpdateLastError = ''
    $script:GuiControls.UpdateIndicator.Content = $script:GuiText.UpdateNewIndicator
    $script:GuiControls.UpdateIndicator.ToolTip = $script:GuiText.UpdateDetailsHint
    if (-not $Release.IsNewer) { $script:GuiControls.UpdateIndicator.Visibility = 'Collapsed' }
    $script:GuiControls.UpdateBanner.Visibility = 'Visible'
    if ($Release.IsNewer) {
        $script:GuiControls.UpdateIndicator.Visibility = 'Visible'
        $script:GuiControls.UpdateTitle.Text = ($script:GuiText.UpdateNewTitle -f $Release.Version)
        $script:GuiControls.UpdateStatus.Text = ($script:GuiText.UpdateNewStatus -f $script:GuiVersion, $Release.PublishedAt.Substring(0, [Math]::Min(10, $Release.PublishedAt.Length)))
        if ([string]::IsNullOrWhiteSpace($Release.DownloadUrl)) {
            $script:GuiControls.DownloadUpdate.Visibility = 'Collapsed'
            $script:GuiControls.UpdateStatus.Text += $script:GuiText.UpdateMissingAsset
        }
        else { $script:GuiControls.DownloadUpdate.Visibility = 'Visible' }
    }
    else {
        $script:GuiControls.UpdateTitle.Text = $script:GuiText.UpdateLatest
        if ((Compare-HddtReleaseVersion $script:GuiVersion $Release.Version) -gt 0) { $script:GuiControls.UpdateTitle.Text = $script:GuiText.UpdateAhead }
        $script:GuiControls.UpdateStatus.Text = ($script:GuiText.UpdateVersions -f $Release.Version, $script:GuiVersion)
        $script:GuiControls.DownloadUpdate.Visibility = 'Collapsed'
    }
}

function Receive-HddtGuiUpdateTask {
    if ($null -eq $script:GuiUpdateTask) { return }
    $task = $script:GuiUpdateTask
    if ($task.Mode -eq 'download') {
        $script:GuiControls.UpdateBanner.Visibility = 'Visible'
        $script:GuiControls.UpdateStatus.Text = ($script:GuiText.UpdateDownloading -f ($task.Context.Received / 1MB), ($script:GuiLatestRelease.Size / 1MB))
    }
    if (-not $task.Handle.IsCompleted) { return }
    try {
        $result = Complete-HddtUpdateTask $task
        if ($task.Mode -eq 'check') { Update-HddtGuiReleaseState $result }
        else {
            $script:GuiUpdateFile = [string]$result
            $script:GuiControls.UpdateIndicator.Visibility = 'Visible'
            $script:GuiControls.UpdateIndicator.Content = $script:GuiText.UpdateSavedIndicator
            $script:GuiControls.UpdateIndicator.ToolTip = $script:GuiText.UpdateSavedHint
            $script:GuiControls.ShowUpdateFile.Visibility = 'Visible'
            $script:GuiControls.UpdateTitle.Text = $script:GuiText.UpdateDownloaded
            $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateDownloadedHelp
            if ([IO.Path]::GetExtension($script:GuiUpdateFile) -eq '.exe') { $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateDownloadedExeHelp }
        }
    }
    catch {
        $script:GuiControls.UpdateBanner.Visibility = 'Visible'
        $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateFailed
        $script:GuiUpdateLastError = $script:GuiText.UpdateFailed
        $script:GuiControls.UpdateIndicator.Visibility = 'Visible'
        $script:GuiControls.UpdateIndicator.Content = $script:GuiText.UpdateErrorIndicator
        $script:GuiControls.UpdateIndicator.ToolTip = $script:GuiUpdateLastError
    }
    finally {
        $script:GuiUpdateTask = $null
        $script:GuiControls.CheckUpdates.IsEnabled = $true
        $script:GuiControls.CheckUpdates.Content = $script:GuiText.UpdateButton
        $script:GuiControls.DownloadUpdate.IsEnabled = $true
    }
    if ($script:GuiCloseAfterUpdate -and $null -eq $script:GuiRun) { $script:GuiWindow.Close() }
}

function Add-HddtGuiLog {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    $log = $script:GuiControls.Log
    if ($log.Text.Length -gt 60000) {
        $text = $log.Text.Substring($log.Text.Length - 40000)
        $newline = $text.IndexOf([char]10)
        if ($newline -ge 0) { $text = $text.Substring($newline + 1) }
        $log.Text = $text
    }
    $null = $script:GuiLogHistory.AppendLine($Line)
    $atBottom = ($log.VerticalOffset + $log.ViewportHeight -ge $log.ExtentHeight - 1)
    $log.AppendText($Line + [Environment]::NewLine)
    if ($atBottom) { $log.ScrollToEnd() }
    if ($null -ne $script:GuiLogDialog -and $script:GuiLogDialog.Window.IsVisible) {
        $fullLog = $script:GuiLogDialog.Text
        $follow = ($fullLog.VerticalOffset + $fullLog.ViewportHeight -ge $fullLog.ExtentHeight - 1)
        $fullLog.AppendText($Line + [Environment]::NewLine)
        if ($follow) { $fullLog.ScrollToEnd() }
    }
    $level = 'INFO'; $tone = '#64748B'
    if ($Line -match '(?i)\[(ERROR|LOI)\]|that bai|HTTP (500|504)') { $level = 'ERROR'; $tone = '#C43535' }
    elseif ($Line -match '(?i)\[WARN\]|429|thu lai') { $level = 'WARNING'; $tone = '#A46713' }
    elseif ($Line -match '(?i)\[OK\]|hoan tat') { $level = 'SUCCESS'; $tone = '#17834B' }
    $row = [pscustomobject]@{ Time = (Get-Date -Format 'HH:mm:ss'); Level = $level; Tone = $tone; Message = $Line }
    $script:GuiActivity.Add($row)
    $script:GuiRecentActivity.Add($row)
    while ($script:GuiRecentActivity.Count -gt $script:GuiRecentLimit) { $script:GuiRecentActivity.RemoveAt(0) }
    if ($script:GuiActivity.Count -gt 500) { $script:GuiActivity.RemoveAt(0) }
    if ($null -ne $script:GuiRun -and $script:GuiRun.Context.StopRequested) { return }
    $progress = Get-HddtDesktopProgress $Line
    if ($null -ne $progress) {
        $script:GuiControls.Progress.IsIndeterminate = $false
        $script:GuiControls.Progress.Value = $progress.Percent
        Set-HddtGuiState downloading
        $script:GuiProgress = $progress
        Update-HddtGuiMetrics
        $script:GuiControls.Status.Text = ($script:GuiText.ProgressStatus -f $progress.Current, $progress.Total, $progress.Percent)
    }
    elseif ($Line -match '\[XU.T FILE\]') {
        $script:GuiControls.Progress.IsIndeterminate = $true
        Set-HddtGuiState exporting
        $script:GuiControls.Status.Text = $script:GuiText.ProgressExport
    }
}

function Set-HddtGuiBusy {
    param([bool]$Busy)
    foreach ($name in @('ModeTabs', 'LoadConfig', 'AdvancedSettings', 'OutputPath', 'BrowseOutput', 'StartRun', 'ProxyOptions')) {
        $script:GuiControls[$name].IsEnabled = -not $Busy
    }
    if ($Busy) {
        $script:GuiControls.TogglePassword.IsChecked = $false
        $script:GuiControls.ProxyPopup.IsOpen = $false
    }
    $script:GuiControls.StopRun.IsEnabled = $Busy
}

function Get-HddtGuiValues {
    $c = $script:GuiControls
    $values = Copy-HddtConfigValues $script:GuiDefaults
    $output = $c.OutputPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($output) -or [IO.Path]::GetExtension($output).ToLowerInvariant() -ne '.xlsx') {
        throw 'Hay chon file dau ra co phan mo rong .xlsx.'
    }
    $output = Resolve-RepositoryPath $script:GuiDataRoot $output
    if (Test-Path -LiteralPath $output) { throw 'File dau ra da ton tai. Hay chon ten moi de giu nguyen du lieu cu.' }
    $values['OUTPUT_DIR'] = [IO.Path]::GetDirectoryName($output)
    $values['OUTPUT_XLSX'] = [IO.Path]::GetFileName($output)
    $values['LOCAL_OUTPUT_XLSX'] = [IO.Path]::GetFileName($output)
    $values['OVERWRITE_OUTPUT'] = 'false'
    $values['PROGRESS_EVERY'] = '1'
    if ($c.ModeTabs.SelectedIndex -eq 1) {
        $source = Resolve-RepositoryPath $script:GuiDataRoot $c.XmlDirectory.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($c.XmlDirectory.Text) -or -not (Test-Path -LiteralPath $source -PathType Container)) {
            throw 'Hay chon thu muc XML hop le.'
        }
        $values['LOCAL_XML_DIR'] = $source
        $values['LOCAL_DIRECTION'] = [string]$c.LocalDirection.SelectedItem.Tag
        foreach ($key in @('GDT_USERNAME', 'GDT_PASSWORD', 'PROXY_USERNAME', 'PROXY_PASSWORD', 'PROXY_URL')) { $values.Remove($key) }
    }
    else {
        $fromDate = [datetime]::MinValue
        $toDate = [datetime]::MinValue
        $culture = [Globalization.CultureInfo]::GetCultureInfo('vi-VN')
        if (-not [datetime]::TryParse($c.FromDate.Text, $culture, [Globalization.DateTimeStyles]::None, [ref]$fromDate) -or
            -not [datetime]::TryParse($c.ToDate.Text, $culture, [Globalization.DateTimeStyles]::None, [ref]$toDate)) {
            throw 'Hay chon khoang ngay hop le (dd/MM/yyyy).'
        }
        $values['GDT_USERNAME'] = $c.Username.Text.Trim()
        $values['GDT_PASSWORD'] = $c.Password.Password
        $values['FROM_DATE'] = $fromDate.ToString('dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
        $values['TO_DATE'] = $toDate.ToString('dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
        $values['INVOICE_DIRECTION'] = [string]$c.Direction.SelectedItem.Tag
        $values['INCLUDE_REGULAR'] = ([bool]$c.IncludeRegular.IsChecked).ToString().ToLowerInvariant()
        $values['INCLUDE_SCO'] = ([bool]$c.IncludeSco.IsChecked).ToString().ToLowerInvariant()
        $values['FETCH_RELATED'] = ([bool]$c.FetchRelated.IsChecked).ToString().ToLowerInvariant()
        $values['REDOWNLOAD_XML'] = ([bool]$c.RedownloadXml.IsChecked).ToString().ToLowerInvariant()
        $values['PROXY_URL'] = $c.ProxyUrl.Text.Trim()
        $values['PROXY_USERNAME'] = $c.ProxyUsername.Text.Trim()
        $values['PROXY_PASSWORD'] = $c.ProxyPassword.Password
        if ([string]::IsNullOrWhiteSpace($values['PROXY_URL'])) {
            $values['PROXY_USERNAME'] = ''
            $values['PROXY_PASSWORD'] = ''
        }
        # Dung cung validator voi CLI, khong ghi .env tam.
        $null = Get-HddtConfig -EnvFile (Join-Path $script:GuiRoot '.env') -RepositoryRoot $script:GuiRoot -Values $values
    }
    return $values
}

function Request-HddtGuiStop {
    if ($null -eq $script:GuiRun) { return }
    Set-HddtGuiState stopping
    Request-HddtDesktopStop $script:GuiRun
    $script:GuiControls.StopRun.IsEnabled = $false
    $script:GuiControls.Status.Text = $script:GuiText.ProgressStopping
    $script:GuiControls.Progress.IsIndeterminate = $true
}

function Set-HddtGuiSettingsFields {
    param([hashtable]$Values)
    foreach ($field in @(Get-HddtDesktopSettingFields)) {
        $control = $script:GuiSettingsDialog.Controls[$field.Key]
        $text = Get-EnvValue $Values $field.Key $field.Default
        if ($field.Kind -eq 'Boolean') { $control.IsChecked = ConvertTo-EnvBoolean $field.Key $text }
        else { $control.Text = $text }
    }
}

function New-HddtGuiSettingsDialog {
    $reader = New-Object IO.StringReader([IO.File]::ReadAllText((Join-Path $script:GuiRoot 'ui/AdvancedSettings.xaml'), [Text.Encoding]::UTF8))
    $xmlReader = [Xml.XmlReader]::Create($reader)
    try { $window = [Windows.Markup.XamlReader]::Load($xmlReader) }
    finally { $xmlReader.Dispose(); $reader.Dispose() }
    $window.Height = [Math]::Min(750, [Windows.SystemParameters]::WorkArea.Height - 32)
    $window.Width = [Math]::Min(720, [Windows.SystemParameters]::WorkArea.Width - 32)
    $controls = @{}
    $names = @('RememberSettings', 'SettingsError', 'ResetSettings', 'CancelSettings', 'ApplySettings')
    $names += @(Get-HddtDesktopSettingFields | ForEach-Object { $_.Key })
    foreach ($name in $names) {
        $control = $window.FindName($name)
        if ($null -eq $control) { throw ('Thieu control cai dat: ' + $name) }
        $controls[$name] = $control
    }
    $script:GuiSettingsDialog = [pscustomobject]@{ Window = $window; Controls = $controls; Applied = $false }
    Set-HddtGuiSettingsFields $script:GuiDefaults
    $controls.ResetSettings.Add_Click({
        Set-HddtGuiSettingsFields @{}
        $script:GuiSettingsDialog.Controls.SettingsError.Text = ''
    })
    $controls.CancelSettings.Add_Click({ $script:GuiSettingsDialog.Window.Close() })
    $controls.ApplySettings.Add_Click({
        try {
            $values = @{}
            foreach ($field in @(Get-HddtDesktopSettingFields)) {
                $control = $script:GuiSettingsDialog.Controls[$field.Key]
                if ($field.Kind -eq 'Boolean') { $values[$field.Key] = ([bool]$control.IsChecked).ToString().ToLowerInvariant() }
                else { $values[$field.Key] = $control.Text.Trim() }
            }
            $settings = ConvertTo-HddtDesktopSettings $values
            if ($script:GuiSettingsDialog.Controls.RememberSettings.IsChecked) {
                Save-HddtDesktopSettings -Path $script:GuiSettingsPath -Values $settings
            }
            foreach ($key in $settings.Keys) { $script:GuiDefaults[$key] = $settings[$key] }
            $script:GuiSettingsDialog.Applied = $true
            $script:GuiControls.Status.Text = 'Da ap dung cai dat nang cao cho lan chay tiep theo.'
            $script:GuiSettingsDialog.Window.Close()
        }
        catch {
            $script:GuiSettingsDialog.Controls.SettingsError.Text = Protect-HddtDesktopText $_.Exception.Message @($script:GuiControls.Password.Password, $script:GuiControls.ProxyPassword.Password)
        }
    })
    return $script:GuiSettingsDialog
}

try { $script:GuiDefaults = Read-HddtDesktopSettings $script:GuiSettingsPath }
catch {
    $script:GuiControls.Status.Text = 'Khong nap duoc cai dat da luu; dang dung mac dinh. Hay mo Cai dat nang cao.'
    Add-HddtGuiLog '[WARN] File hddt-settings.json khong hop le hoac khong doc duoc; dang dung mac dinh.'
}
$script:GuiControls.AdvancedSettings.Add_Click({
    if ($null -ne $script:GuiRun) { return }
    $dialog = New-HddtGuiSettingsDialog
    $dialog.Window.Owner = $script:GuiWindow
    $null = $dialog.Window.ShowDialog()
})

$script:GuiControls.BrowseXml.Add_Click({
    $dialog = New-Object Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Chon thu muc XML'
    try {
        if ($dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) { $script:GuiControls.XmlDirectory.Text = $dialog.SelectedPath }
    }
    finally { $dialog.Dispose() }
})
$script:GuiControls.BrowseOutput.Add_Click({
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = 'Excel workbook (*.xlsx)|*.xlsx'
    $dialog.DefaultExt = '.xlsx'
    $dialog.FileName = [IO.Path]::GetFileName($script:GuiControls.OutputPath.Text)
    $dialog.OverwritePrompt = $true
    if ($dialog.ShowDialog($script:GuiWindow)) { $script:GuiControls.OutputPath.Text = $dialog.FileName }
})
$script:GuiControls.LoadConfig.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = 'Cau hinh .env (.env;*.env)|.env;*.env|Tat ca file (*.*)|*.*'
    $dialog.FileName = '.env'
    $dialog.InitialDirectory = $script:GuiDataRoot
    if (-not $dialog.ShowDialog($script:GuiWindow)) { return }
    try {
        $values = Read-DotEnvFile $dialog.FileName
        $c = $script:GuiControls
        foreach ($field in @(@('Username', 'GDT_USERNAME'), @('ProxyUrl', 'PROXY_URL'), @('ProxyUsername', 'PROXY_USERNAME'), @('XmlDirectory', 'LOCAL_XML_DIR'))) {
            $c[$field[0]].Text = Get-EnvValue $values $field[1] ''
        }
        $c.Password.Password = Get-EnvValue $values 'GDT_PASSWORD' ''
        $c.ProxyPassword.Password = Get-EnvValue $values 'PROXY_PASSWORD' ''
        foreach ($field in @(@('FromDate', 'FROM_DATE'), @('ToDate', 'TO_DATE'))) {
            $dateText = Get-EnvValue $values $field[1] ''
            if (-not [string]::IsNullOrWhiteSpace($dateText)) {
                $c[$field[0]].SelectedDate = [datetime]::ParseExact($dateText, 'dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
            }
        }
        foreach ($field in @(@('Direction', 'INVOICE_DIRECTION', 'both'), @('LocalDirection', 'LOCAL_DIRECTION', 'auto'))) {
            $tag = Get-EnvValue $values $field[1] $field[2]
            foreach ($item in $c[$field[0]].Items) { if ($item.Tag -eq $tag) { $c[$field[0]].SelectedItem = $item } }
        }
        foreach ($field in @(@('IncludeRegular', 'INCLUDE_REGULAR', 'true'), @('IncludeSco', 'INCLUDE_SCO', 'true'),
            @('FetchRelated', 'FETCH_RELATED', 'true'), @('RedownloadXml', 'REDOWNLOAD_XML', 'false'))) {
            $c[$field[0]].IsChecked = ConvertTo-EnvBoolean $field[1] (Get-EnvValue $values $field[1] $field[2])
        }
        $outputDirectory = Resolve-RepositoryPath $script:GuiDataRoot (Get-EnvValue $values 'OUTPUT_DIR' 'output')
        $c.OutputPath.Text = Join-Path $outputDirectory ('HoaDonDienTu_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $values.Remove('GDT_PASSWORD')
        $values.Remove('PROXY_PASSWORD')
        $script:GuiDefaults = $values
        $c.Status.Text = 'Da nap cau hinh. File .env duoc giu nguyen.'
    }
    catch {
        $safe = Protect-HddtDesktopText $_.Exception.Message @($script:GuiControls.Password.Password, $script:GuiControls.ProxyPassword.Password)
        $null = [Windows.MessageBox]::Show($script:GuiWindow, $safe, 'Khong nap duoc cau hinh', 'OK', 'Warning')
    }
})
$script:GuiControls.StartRun.Add_Click({
    if ($null -ne $script:GuiRun) { return }
    try {
        $values = Get-HddtGuiValues
        $mode = 'download'
        if ($script:GuiControls.ModeTabs.SelectedIndex -eq 1) { $mode = 'local' }
        $script:GuiLastWorkbook = Join-Path $values['OUTPUT_DIR'] $values['OUTPUT_XLSX']
        $script:GuiControls.Log.Clear()
        $script:GuiActivity.Clear()
        $script:GuiRecentActivity.Clear()
        $null = $script:GuiLogHistory.Clear()
        Set-HddtGuiState connecting
        $script:GuiControls.LogDetail.Text = $script:GuiText.LogHint
        $script:GuiProgress = $null
        $script:GuiStartedAt = Get-Date
        $script:GuiControls.ProcessedCount.Text = '-'
        $script:GuiControls.RemainingCount.Text = '-'
        $script:GuiControls.PercentText.Text = '0%'
        $script:GuiControls.EstimatedTime.Text = '-'
        Update-HddtGuiMetrics
        $script:GuiControls.OpenExcel.IsEnabled = $false
        $script:GuiControls.Status.Text = $script:GuiText.ProgressStarting
        $script:GuiControls.Progress.Value = 0
        $script:GuiControls.Progress.IsIndeterminate = $true
        $script:GuiRun = Start-HddtDesktopRun -Root $script:GuiRoot -Values $values -Mode $mode
        $values.Clear()
        Set-HddtGuiBusy $true
    }
    catch {
        $script:GuiControls.Progress.IsIndeterminate = $false
        Set-HddtGuiState error
        $script:GuiControls.Status.Text = 'Khong bat dau duoc. Hay kiem tra thong tin.'
        $safe = Protect-HddtDesktopText $_.Exception.Message @($script:GuiControls.Password.Password, $script:GuiControls.ProxyPassword.Password)
        if ($SmokeTest) { throw $safe }
        $null = [Windows.MessageBox]::Show($script:GuiWindow, $safe, 'Kiem tra thong tin', 'OK', 'Warning')
    }
})
$script:GuiControls.StopRun.Add_Click({ Request-HddtGuiStop })
$script:GuiControls.OpenExcel.Add_Click({
    if (Test-Path -LiteralPath $script:GuiLastWorkbook -PathType Leaf) {
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = $script:GuiLastWorkbook
        $info.UseShellExecute = $true
        try { $null = [Diagnostics.Process]::Start($info) }
        catch { $script:GuiControls.Status.Text = 'Khong mo duoc Excel. Hay mo file tu thu muc ket qua.' }
    }
})
$script:GuiControls.OpenFolder.Add_Click({
    try {
        $directory = [IO.Path]::GetDirectoryName((Resolve-RepositoryPath $script:GuiDataRoot $script:GuiControls.OutputPath.Text))
        if (Test-Path -LiteralPath $directory -PathType Container) {
            $info = New-Object Diagnostics.ProcessStartInfo
            $info.FileName = $directory
            $info.UseShellExecute = $true
            $null = [Diagnostics.Process]::Start($info)
        }
        else { $script:GuiControls.Status.Text = 'Thu muc ket qua chua duoc tao.' }
    }
    catch { $script:GuiControls.Status.Text = 'Khong mo duoc thu muc ket qua.' }
})
$script:GuiControls.TogglePassword.Add_Checked({ Set-HddtGuiPasswordVisibility $true })
$script:GuiControls.TogglePassword.Add_Unchecked({ Set-HddtGuiPasswordVisibility $false })
$script:GuiControls.Password.Add_PasswordChanged({
    if ($script:GuiPasswordSync -or -not $script:GuiControls.TogglePassword.IsChecked) { return }
    $script:GuiPasswordSync = $true
    try { $script:GuiControls.PasswordVisible.Text = $script:GuiControls.Password.Password }
    finally { $script:GuiPasswordSync = $false }
})
$script:GuiControls.PasswordVisible.Add_TextChanged({
    if ($script:GuiPasswordSync -or -not $script:GuiControls.TogglePassword.IsChecked) { return }
    $script:GuiPasswordSync = $true
    try { $script:GuiControls.Password.Password = $script:GuiControls.PasswordVisible.Text }
    finally { $script:GuiPasswordSync = $false }
})

$script:GuiControls.ProxyOptions.Add_Click({ $script:GuiControls.ProxyPopup.IsOpen = -not $script:GuiControls.ProxyPopup.IsOpen })
$script:GuiControls.DismissUpdate.Add_Click({ $script:GuiControls.UpdateBanner.Visibility = 'Collapsed' })
$script:GuiControls.ActivityTab.Add_Checked({ $script:GuiControls.ActivityLog.Visibility = 'Visible'; $script:GuiControls.Log.Visibility = 'Collapsed' })
$script:GuiControls.TechnicalTab.Add_Checked({ $script:GuiControls.ActivityLog.Visibility = 'Collapsed'; $script:GuiControls.Log.Visibility = 'Visible' })
$script:GuiControls.WorkspaceRoot.Add_SizeChanged({ Update-HddtGuiLayout })
$script:GuiControls.UpdateIndicator.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($script:GuiUpdateFile) -and [string]::IsNullOrWhiteSpace($script:GuiUpdateLastError)) { $script:GuiControls.ShowUpdateFile.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent))) }
    else { $script:GuiControls.ViewReleaseNotes.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent))) }
})
$script:GuiControls.ViewReleaseNotes.Add_Click({
    $text = $script:GuiControls.UpdateStatus.Text
    if ($null -ne $script:GuiLatestRelease) { $text = $script:GuiLatestRelease.Notes }
    if (-not [string]::IsNullOrWhiteSpace($script:GuiUpdateLastError)) { $text = $script:GuiUpdateLastError }
    $offer = ($null -ne $script:GuiLatestRelease -and $script:GuiLatestRelease.IsNewer -and -not [string]::IsNullOrWhiteSpace($script:GuiLatestRelease.DownloadUrl))
    $dialog = New-HddtGuiTextDialog -Title $script:GuiControls.UpdateTitle.Text -Text $text -OfferDownload:$offer
    $null = $dialog.Window.ShowDialog()
})
$script:GuiControls.ShowFullLog.Add_Click({
    $dialog = New-HddtGuiTextDialog -Title $script:GuiText.FullLogTitle -Text $script:GuiLogHistory.ToString() -Technical
    $script:GuiLogDialog = $dialog
    $dialog.Text.ScrollToEnd()
    try { $null = $dialog.Window.ShowDialog() } finally { $script:GuiLogDialog = $null }
})
$script:GuiControls.ActivityLog.Add_SizeChanged({ Update-HddtGuiRecentLog })

$script:GuiControls.ActivityLog.Add_SelectionChanged({
    if ($null -ne $script:GuiControls.ActivityLog.SelectedItem) { $script:GuiControls.LogDetail.Text = $script:GuiControls.ActivityLog.SelectedItem.Message }
})
$script:GuiControls.CheckUpdates.Add_Click({ Start-HddtGuiUpdateCheck })
$script:GuiControls.ReleasePage.Add_Click({
    try {
        $url = 'https://github.com/dieutx/hddt-downloader-windows/releases/latest'
        if ($null -ne $script:GuiLatestRelease) { $url = $script:GuiLatestRelease.ReleaseUrl }
        Open-HddtGuiReleaseUrl $url
    }
    catch { $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateBrowserFailed }
})
$script:GuiControls.DownloadUpdate.Add_Click({
    if ($null -ne $script:GuiUpdateTask -or $null -eq $script:GuiLatestRelease) { return }
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $extension = [IO.Path]::GetExtension($script:GuiLatestRelease.AssetName).ToLowerInvariant()
    $dialog.Filter = 'Ung dung Windows (*.exe)|*.exe'
    if ($extension -eq '.zip') { $dialog.Filter = 'Goi Windows portable (*.zip)|*.zip' }
    $dialog.DefaultExt = $extension
    $dialog.FileName = $script:GuiLatestRelease.AssetName
    $dialog.OverwritePrompt = $true
    if (-not $dialog.ShowDialog($script:GuiWindow)) { return }
    if (Test-Path -LiteralPath $dialog.FileName) {
        $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateExists
        return
    }
    try {
        $script:GuiUpdateTask = Start-HddtUpdateTask -Root $script:GuiRoot -CurrentVersion $script:GuiVersion -Mode download -Release $script:GuiLatestRelease -Destination $dialog.FileName
        $script:GuiControls.CheckUpdates.IsEnabled = $false
        $script:GuiControls.DownloadUpdate.IsEnabled = $false
        $script:GuiControls.ShowUpdateFile.Visibility = 'Collapsed'
        $script:GuiUpdateFile = ''
    }
    catch { $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateStartFailed }
})
$script:GuiControls.ShowUpdateFile.Add_Click({
    try {
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = [IO.Path]::GetDirectoryName($script:GuiUpdateFile)
        $info.UseShellExecute = $true
        $null = [Diagnostics.Process]::Start($info)
    }
    catch { $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateFolderFailed + $script:GuiUpdateFile }
})

$script:GuiTimer = New-Object Windows.Threading.DispatcherTimer
$script:GuiTimer.Interval = [TimeSpan]::FromMilliseconds(200)
$script:GuiTimer.Add_Tick({
    Receive-HddtGuiUpdateTask
    if ($null -eq $script:GuiRun) { return }
    Update-HddtGuiMetrics
    foreach ($line in @(Receive-HddtDesktopLog $script:GuiRun)) { Add-HddtGuiLog $line }
    if (-not $script:GuiRun.Handle.IsCompleted) { return }
    # Lay not dong log cuoi cung truoc khi giai phong pipeline.
    foreach ($line in @(Receive-HddtDesktopLog $script:GuiRun)) { Add-HddtGuiLog $line }
    Complete-HddtDesktopRun $script:GuiRun
    $context = $script:GuiRun.Context
    if (-not [string]::IsNullOrWhiteSpace([string]$context.Error)) { Add-HddtGuiLog ('[ERROR] ' + $context.Error) }
    $script:GuiControls.Progress.IsIndeterminate = $false
    $hasWorkbook = Test-Path -LiteralPath $script:GuiLastWorkbook -PathType Leaf
    $script:GuiControls.OpenExcel.IsEnabled = $hasWorkbook
    if ($context.ExitCode -ne 0) { Set-HddtGuiState error }
    elseif ($context.StopRequested) { Set-HddtGuiState idle }
    else { Set-HddtGuiState completed }
    if ($context.ExitCode -eq 2) { $script:GuiControls.Status.Text = $script:GuiText.ProgressNoData }
    elseif ($context.ExitCode -ne 0) { $script:GuiControls.Status.Text = $script:GuiText.ProgressFailed }
    elseif ($context.StopRequested) { $script:GuiControls.Status.Text = $script:GuiText.ProgressStopped }
    else {
        $script:GuiControls.Status.Text = $script:GuiText.ProgressDone
        $script:GuiControls.Progress.Value = 100
        $script:GuiControls.PercentText.Text = '100%'
        $script:GuiControls.EstimatedTime.Text = '00:00:00'
    }
    $script:GuiRun = $null
    Set-HddtGuiBusy $false
    if ($script:GuiCloseWhenDone -and $null -eq $script:GuiUpdateTask) { $script:GuiWindow.Close() }
})
$script:GuiWindow.Add_Closing({
    param($sender, $eventArgs)
    if ($null -ne $script:GuiUpdateTask) {
        $eventArgs.Cancel = $true
        $script:GuiCloseAfterUpdate = $true
        $script:GuiUpdateTask.Context.CancelRequested = $true
        try {
            $cancellation = $script:GuiUpdateTask.Context.Cancellation
            if ($null -ne $cancellation) { $cancellation.Cancel() }
        }
        catch [ObjectDisposedException] { }
        $script:GuiControls.UpdateStatus.Text = $script:GuiText.UpdateClosing
    }
    if ($null -ne $script:GuiRun) {
        $eventArgs.Cancel = $true
        $script:GuiCloseWhenDone = $true
        Request-HddtGuiStop
    }
})
$script:GuiWindow.Add_Closed({
    $script:GuiTimer.Stop()
    $script:GuiControls.TogglePassword.IsChecked = $false
    $script:GuiControls.PasswordVisible.Clear()
    $script:GuiControls.Password.Clear()
    $script:GuiControls.ProxyPassword.Clear()
    $script:GuiDefaults.Clear()
})

if (-not [string]::IsNullOrWhiteSpace($PreviewPath)) { Save-HddtGuiPreview -Path $PreviewPath -Width $PreviewWidth -Height $PreviewHeight }
if ($SmokeTest) {
    if (-not [string]::IsNullOrWhiteSpace($SmokeXmlDirectory)) {
        $script:GuiControls.ModeTabs.SelectedIndex = 1
        $script:GuiControls.XmlDirectory.Text = $SmokeXmlDirectory
        $script:GuiControls.LocalDirection.SelectedIndex = 1
        $script:GuiControls.OutputPath.Text = Join-Path $script:GuiOutputRoot 'offline-test.xlsx'
        $script:GuiTimer.Start()
        $script:GuiControls.StartRun.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
        $deadline = [datetime]::UtcNow.AddSeconds(30)
        while ($null -ne $script:GuiRun -and [datetime]::UtcNow -lt $deadline) {
            $frame = New-Object Windows.Threading.DispatcherFrame
            $null = [Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background, [Action]{ $frame.Continue = $false })
            [Windows.Threading.Dispatcher]::PushFrame($frame)
            Start-Sleep -Milliseconds 20
        }
        $script:GuiTimer.Stop()
        if ($null -ne $script:GuiRun -or -not (Test-Path -LiteralPath $script:GuiControls.OutputPath.Text)) { throw 'Offline EXE smoke test that bai.' }
    }
    if (-not [string]::IsNullOrWhiteSpace($SmokeReportPath)) {
        $report = @{ Version = $script:GuiVersion; RuntimeDirectory = $script:GuiRoot; DataDirectory = $script:GuiDataRoot;
            SettingsPath = $script:GuiSettingsPath; OutputPath = $script:GuiControls.OutputPath.Text }
        [IO.File]::WriteAllText($SmokeReportPath, ($report | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
    }
    Write-Output 'WPF smoke test passed.'
    return
}
try {
    $script:GuiTimer.Start()
    $script:GuiWindow.Add_ContentRendered({ Start-HddtGuiUpdateCheck })
    $null = $script:GuiWindow.ShowDialog()
}
finally { $script:GuiTimer.Stop() }
