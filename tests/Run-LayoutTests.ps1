[CmdletBinding()]
param([string]$PreviewDirectory)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'Start-HddtGui.ps1') -SmokeTest
if ([string]::IsNullOrWhiteSpace($PreviewDirectory)) { $PreviewDirectory = Join-Path ([IO.Path]::GetTempPath()) ('hddt-layout-' + [guid]::NewGuid().ToString('N')) }
$null = [IO.Directory]::CreateDirectory($PreviewDirectory)

function Assert-Layout {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAILED: ' + $Message) }
}

$cases = @(
    @{ Name = '1160x768-100'; Width = 1160; Height = 736; Dpi = 96 }
    @{ Name = '1366x768-100'; Width = 1366; Height = 736; Dpi = 96 }
    @{ Name = '1440x900-100'; Width = 1440; Height = 868; Dpi = 96 }
    @{ Name = '1920x1080-100'; Width = 1920; Height = 1048; Dpi = 96 }
    @{ Name = '1160x768-125'; Width = 928; Height = 584; Dpi = 120 }
    @{ Name = '1366x768-125'; Width = 1093; Height = 584; Dpi = 120 }
    @{ Name = '1440x900-125'; Width = 1152; Height = 690; Dpi = 120 }
    @{ Name = '1920x1080-125'; Width = 1536; Height = 834; Dpi = 120 }
)
$checks = @('Username', 'Password', 'Direction', 'FromDate', 'ToDate', 'IncludeRegular', 'IncludeSco', 'BrowseOutput',
    'StartRun', 'StopRun', 'AdvancedSettings', 'LoadConfig', 'Status', 'Progress', 'ActivityLog', 'ShowFullLog', 'OpenExcel', 'OpenFolder')
$results = New-Object Collections.Generic.List[object]
foreach ($case in $cases) {
    foreach ($state in @('idle', 'downloading', 'completed', 'error')) {
        Set-HddtGuiState $state
        Set-HddtGuiBusy ($state -eq 'downloading')
        $script:GuiControls.OpenExcel.IsEnabled = ($state -eq 'completed')
        $script:GuiControls.UpdateBanner.Visibility = 'Collapsed'
        $script:GuiControls.UpdateIndicator.Visibility = 'Collapsed'
        $script:GuiControls.ModeTabs.SelectedIndex = 0
        $script:GuiActivity.Clear(); $script:GuiRecentActivity.Clear()
        $script:GuiProgress = $null
        $script:GuiControls.ProcessedCount.Text = '-'; $script:GuiControls.RemainingCount.Text = '-'
        $script:GuiControls.EstimatedTime.Text = '-'
        $script:GuiControls.Log.Clear(); $null = $script:GuiLogHistory.Clear()
        $script:GuiStartedAt = (Get-Date).AddSeconds(-30)
        for ($i = 1; $i -le 8; $i++) { Add-HddtGuiLog ('[INFO] Synthetic activity {0}' -f $i) }
        if ($state -eq 'downloading') { Add-HddtGuiLog '[3/10 | 30%] Synthetic XML progress' }
        elseif ($state -eq 'error') {
            Add-HddtGuiLog '[ERROR] Synthetic network error; open full log for details.'
            $script:GuiControls.Status.Text = $script:GuiText.ProgressFailed
        }
        elseif ($state -eq 'completed') {
            Add-HddtGuiLog '[OK] Synthetic export completed.'
            $script:GuiControls.Status.Text = $script:GuiText.ProgressDone
            $script:GuiControls.Progress.Value = 100; $script:GuiControls.PercentText.Text = '100%'
            $script:GuiProgress = [pscustomobject]@{ Current = 10; Total = 10; Percent = 100 }
            Update-HddtGuiMetrics
        }
        else {
            $script:GuiControls.Status.Text = $script:GuiText.State_idle
            $script:GuiControls.Progress.Value = 0; $script:GuiControls.PercentText.Text = '0%'
        }
        Set-HddtGuiState $state
        $image = Join-Path $PreviewDirectory ($case.Name + '-' + $state + '.png')
        Save-HddtGuiPreview -Path $image -Width $case.Width -Height $case.Height -Dpi $case.Dpi
        foreach ($name in $checks) {
            $control = $script:GuiControls[$name]
            Assert-Layout ($control.ActualWidth -gt 0 -and $control.ActualHeight -gt 0) ($case.Name + ' visible ' + $name)
            $bounds = $control.TransformToAncestor($script:GuiControls.WorkspaceRoot).TransformBounds((New-Object Windows.Rect(0, 0, $control.ActualWidth, $control.ActualHeight)))
            Assert-Layout ($bounds.Left -ge -1 -and $bounds.Top -ge -1 -and $bounds.Right -le $case.Width - 23 -and $bounds.Bottom -le $case.Height - 15) ($case.Name + ' no overflow ' + $name)
        }
        $ancestor = [Windows.Media.VisualTreeHelper]::GetParent($script:GuiControls.Status)
        while ($null -ne $ancestor -and $ancestor -ne $script:GuiControls.WorkspaceRoot) {
            Assert-Layout ($ancestor -isnot [Windows.Controls.ScrollViewer]) 'Progress is not inside a page ScrollViewer'
            $ancestor = [Windows.Media.VisualTreeHelper]::GetParent($ancestor)
        }
        Assert-Layout ($script:GuiControls.ActivityLog.Columns[2].ActualWidth -ge $script:GuiControls.ActivityLog.ActualWidth - 184) ($case.Name + ' message column uses available width')
        Assert-Layout ($script:GuiRecentLimit -ge 4 -and $script:GuiRecentLimit -le 6) ($case.Name + ' four to six recent log rows fit')
        $results.Add([pscustomobject]@{ Size = $case.Name; State = $state; RecentRows = $script:GuiRecentLimit; FullPageScroll = $false })
    }
}
# Mo thong bao cap nhat khong che cac nut chinh; man hinh ngan dung indicator.
$script:GuiControls.UpdateBanner.Visibility = 'Visible'
$script:GuiControls.UpdateIndicator.Visibility = 'Visible'
Save-HddtGuiPreview -Path (Join-Path $PreviewDirectory 'compact-notification.png') -Width 928 -Height 584 -Dpi 120
Assert-Layout ($script:GuiControls.UpdateBanner.MaxHeight -eq 0) 'Compact viewport uses header update indicator'
Assert-Layout ($script:GuiControls.UpdateIndicator.ActualWidth -gt 0) 'Compact update details remain accessible'

$dialog = New-HddtGuiTextDialog -Title 'Synthetic log' -Text $script:GuiLogHistory.ToString() -Technical
Assert-Layout ($dialog.Text.Text.Contains('Synthetic activity 1')) 'Expanded log retains rows outside compact history'
$dialog.Window.Close()
[IO.File]::WriteAllText((Join-Path $PreviewDirectory 'layout-results.json'), ($results | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
Write-Output ('All layout tests passed. Previews: ' + $PreviewDirectory)
