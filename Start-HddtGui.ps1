[CmdletBinding()]
param([switch]$SmokeTest, [string]$PreviewPath)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Giao dien WPF chi chay tren Windows 10/11.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Hay chay bang powershell.exe -STA -File Start-HddtGui.ps1.' }

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
. (Join-Path $PSScriptRoot 'src/Config.ps1')
. (Join-Path $PSScriptRoot 'src/ExcelExporter.ps1')
. (Join-Path $PSScriptRoot 'src/Desktop.ps1')
. (Join-Path $PSScriptRoot 'src/DesktopSettings.ps1')

$script:GuiRoot = $PSScriptRoot
$script:GuiDefaults = @{}
$script:GuiRun = $null
$script:GuiLastWorkbook = ''
$script:GuiCloseWhenDone = $false
$script:GuiControls = @{}
$script:GuiSettingsPath = Join-Path $PSScriptRoot 'hddt-settings.json'
$script:GuiSettingsDialog = $null
$reader = New-Object IO.StringReader([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ui/MainWindow.xaml'), [Text.Encoding]::UTF8))
$xmlReader = [Xml.XmlReader]::Create($reader)
try { $script:GuiWindow = [Windows.Markup.XamlReader]::Load($xmlReader) }
finally { $xmlReader.Dispose(); $reader.Dispose() }
$script:GuiWindow.Height = [Math]::Min(860, [Math]::Max(680, [Windows.SystemParameters]::WorkArea.Height - 32))
$script:GuiWindow.Width = [Math]::Min(1000, [Math]::Max(820, [Windows.SystemParameters]::WorkArea.Width - 32))
foreach ($name in @('LoadConfig', 'AdvancedSettings', 'ModeTabs', 'Username', 'Password', 'FromDate', 'ToDate', 'Direction',
    'IncludeRegular', 'IncludeSco', 'ProxyUrl', 'ProxyUsername', 'ProxyPassword', 'FetchRelated', 'RedownloadXml',
    'BrowseXml', 'XmlDirectory', 'LocalDirection', 'OutputPath', 'BrowseOutput', 'StartRun', 'StopRun',
    'OpenExcel', 'OpenFolder', 'Status', 'Progress', 'Log')) {
    $control = $script:GuiWindow.FindName($name)
    if ($null -eq $control) { throw ('Thieu control: {0}' -f $name) }
    $script:GuiControls[$name] = $control
}
$script:GuiControls.FromDate.SelectedDate = (Get-Date).Date.AddDays(1 - (Get-Date).Day)
$script:GuiControls.ToDate.SelectedDate = (Get-Date).Date
$script:GuiControls.OutputPath.Text = Join-Path (Join-Path $PSScriptRoot 'output') ('HoaDonDienTu_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

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
    $log.AppendText($Line + [Environment]::NewLine)
    $log.ScrollToEnd()
    if ($null -ne $script:GuiRun -and $script:GuiRun.Context.StopRequested) { return }
    $progress = Get-HddtDesktopProgress $Line
    if ($null -ne $progress) {
        $script:GuiControls.Progress.IsIndeterminate = $false
        $script:GuiControls.Progress.Value = $progress.Percent
        $script:GuiControls.Status.Text = ('Da xu ly {0}/{1} ({2}%)' -f $progress.Current, $progress.Total, $progress.Percent)
    }
    elseif ($Line -match '\[XU.T FILE\]') {
        $script:GuiControls.Progress.IsIndeterminate = $true
        $script:GuiControls.Status.Text = 'Dang xuat Excel...'
    }
}

function Set-HddtGuiBusy {
    param([bool]$Busy)
    foreach ($name in @('ModeTabs', 'LoadConfig', 'AdvancedSettings', 'OutputPath', 'BrowseOutput', 'StartRun')) {
        $script:GuiControls[$name].IsEnabled = -not $Busy
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
    $output = Resolve-RepositoryPath $script:GuiRoot $output
    if (Test-Path -LiteralPath $output) { throw 'File dau ra da ton tai. Hay chon ten moi de giu nguyen du lieu cu.' }
    $values['OUTPUT_DIR'] = [IO.Path]::GetDirectoryName($output)
    $values['OUTPUT_XLSX'] = [IO.Path]::GetFileName($output)
    $values['LOCAL_OUTPUT_XLSX'] = [IO.Path]::GetFileName($output)
    $values['OVERWRITE_OUTPUT'] = 'false'
    $values['PROGRESS_EVERY'] = '1'
    if ($c.ModeTabs.SelectedIndex -eq 1) {
        $source = Resolve-RepositoryPath $script:GuiRoot $c.XmlDirectory.Text.Trim()
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
    Request-HddtDesktopStop $script:GuiRun
    $script:GuiControls.StopRun.IsEnabled = $false
    $script:GuiControls.Status.Text = 'Dang dung an toan; cho request hien tai va xuat phan du lieu da xu ly...'
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
    $dialog.InitialDirectory = $script:GuiRoot
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
        $outputDirectory = Resolve-RepositoryPath $script:GuiRoot (Get-EnvValue $values 'OUTPUT_DIR' 'output')
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
        $script:GuiControls.OpenExcel.IsEnabled = $false
        $script:GuiControls.Status.Text = 'Dang xu ly...'
        $script:GuiControls.Progress.Value = 0
        $script:GuiControls.Progress.IsIndeterminate = $true
        $script:GuiRun = Start-HddtDesktopRun -Root $script:GuiRoot -Values $values -Mode $mode
        $values.Clear()
        Set-HddtGuiBusy $true
    }
    catch {
        $script:GuiControls.Progress.IsIndeterminate = $false
        $script:GuiControls.Status.Text = 'Khong bat dau duoc. Hay kiem tra thong tin.'
        $safe = Protect-HddtDesktopText $_.Exception.Message @($script:GuiControls.Password.Password, $script:GuiControls.ProxyPassword.Password)
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
        $directory = [IO.Path]::GetDirectoryName((Resolve-RepositoryPath $script:GuiRoot $script:GuiControls.OutputPath.Text))
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
$script:GuiTimer = New-Object Windows.Threading.DispatcherTimer
$script:GuiTimer.Interval = [TimeSpan]::FromMilliseconds(200)
$script:GuiTimer.Add_Tick({
    if ($null -eq $script:GuiRun) { return }
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
    if ($context.ExitCode -eq 2) { $script:GuiControls.Status.Text = 'Khong lay duoc hoa don; xem workbook bao cao loi va nhat ky.' }
    elseif ($context.ExitCode -ne 0) { $script:GuiControls.Status.Text = 'Tac vu co loi. Xem nhat ky de biet chi tiet.' }
    elseif ($context.StopRequested) { $script:GuiControls.Status.Text = 'Da dung an toan. Du lieu da xu ly duoc giu lai.' }
    else {
        $script:GuiControls.Status.Text = 'Hoan tat. Ban co the mo file Excel.'
        $script:GuiControls.Progress.Value = 100
    }
    $script:GuiRun = $null
    Set-HddtGuiBusy $false
    if ($script:GuiCloseWhenDone) { $script:GuiWindow.Close() }
})
$script:GuiWindow.Add_Closing({
    param($sender, $eventArgs)
    if ($null -ne $script:GuiRun) {
        $eventArgs.Cancel = $true
        $script:GuiCloseWhenDone = $true
        Request-HddtGuiStop
    }
})
$script:GuiWindow.Add_Closed({
    $script:GuiTimer.Stop()
    $script:GuiControls.Password.Clear()
    $script:GuiControls.ProxyPassword.Clear()
    $script:GuiDefaults.Clear()
})

if (-not [string]::IsNullOrWhiteSpace($PreviewPath)) {
    $content = $script:GuiWindow.Content
    $script:GuiWindow.Content = $null
    $visual = New-Object Windows.Controls.Border
    $visual.Resources = $script:GuiWindow.Resources
    $visual.Background = $script:GuiWindow.Background
    [Windows.Documents.TextElement]::SetFontFamily($visual, $script:GuiWindow.FontFamily)
    [Windows.Documents.TextElement]::SetFontSize($visual, $script:GuiWindow.FontSize)
    $visual.Child = $content
    $visual.Measure((New-Object Windows.Size(1000, 860)))
    $visual.Arrange((New-Object Windows.Rect(0, 0, 1000, 860)))
    $visual.UpdateLayout()
    $bitmap = New-Object Windows.Media.Imaging.RenderTargetBitmap(1000, 860, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($visual)
    $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.File]::Create($PreviewPath)
    try { $encoder.Save($stream) } finally { $stream.Dispose() }
    $visual.Child = $null
    $script:GuiWindow.Content = $content
}
if ($SmokeTest) { Write-Output 'WPF smoke test passed.'; return }
try {
    $script:GuiTimer.Start()
    $null = $script:GuiWindow.ShowDialog()
}
finally { $script:GuiTimer.Stop() }
