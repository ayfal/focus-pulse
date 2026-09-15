[CmdletBinding()]
param(
    [string]$TaskFolder,
    [int]$DefaultIntervalMinutes = 10,
    [switch]$SelfTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($TaskFolder)) {
    $scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $TaskFolder = Join-Path $scriptRoot 'Tasks'
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$script:State = [ordered]@{
    TaskFolder       = $TaskFolder
    Tasks            = @()
    CurrentTask      = $null
    CurrentTaskPath  = $null
    IsWorking        = $false
    StartedAt        = $null
    UsedCounter      = [TimeSpan]::Zero
    WastedCounter    = [TimeSpan]::Zero
    CurrentInterval  = [TimeSpan]::FromMinutes($DefaultIntervalMinutes)
    ReminderTimer    = $null
    Window           = $null
    TaskList         = $null
    StatusText       = $null
    FolderText       = $null
    IntervalText     = $null
}

function Format-FocusPulseDuration {
    param([TimeSpan]$Duration)

    $hours = [math]::Floor($Duration.TotalHours)
    '{0:00}:{1:00}:{2:00}' -f $hours, $Duration.Minutes, $Duration.Seconds
}

function Format-FocusPulseDate {
    param([datetime]$DateTime)

    $DateTime.ToString('yyyy-MM-dd HH:mm')
}

function Get-FocusPulseTaskFileName {
    param([datetime]$DateTime)

    $DateTime.ToString('yyyy-MM-dd-HH-mm') + '.txt'
}

function ConvertFrom-FocusPulseTaskFileName {
    param([string]$FileName)

    if ($FileName -notmatch '^(?<stamp>\d{4}-\d{2}-\d{2}-\d{2}-\d{2})\.txt$') {
        return $null
    }

    try {
        [datetime]::ParseExact($Matches.stamp, 'yyyy-MM-dd-HH-mm', [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        $null
    }
}

function Get-FocusPulseTaskBody {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return ''
    }

    [System.IO.File]::ReadAllText($Path)
}

function Set-FocusPulseTaskBody {
    param(
        [string]$Path,
        [string]$Body
    )

    if ($null -eq $Body) {
        $Body = ''
    }

    [System.IO.File]::WriteAllText($Path, $Body)
}

function Get-FocusPulseUniqueTaskPath {
    param(
        [string]$Folder,
        [datetime]$DueAt
    )

    $candidateDueAt = $DueAt
    while ($true) {
        $candidatePath = Join-Path $Folder (Get-FocusPulseTaskFileName -DateTime $candidateDueAt)
        if (-not (Test-Path -LiteralPath $candidatePath)) {
            return [pscustomobject]@{
                Path  = $candidatePath
                DueAt = $candidateDueAt
            }
        }

        $candidateDueAt = $candidateDueAt.AddMinutes(1)
    }
}

function Get-FocusPulseTaskItem {
    param([string]$Path)

    $dueAt = ConvertFrom-FocusPulseTaskFileName -FileName ([System.IO.Path]::GetFileName($Path))
    if (-not $dueAt) {
        return $null
    }

    $body = Get-FocusPulseTaskBody -Path $Path
    $firstLine = ($body -split "`r?`n", 2)[0].Trim()
    if ([string]::IsNullOrWhiteSpace($firstLine)) {
        $firstLine = '(empty)'
    }

    [pscustomobject]@{
        DueAt   = $dueAt
        Path    = $Path
        Body    = $body
        Display = ('{0}  {1}' -f (Format-FocusPulseDate $dueAt), $firstLine)
    }
}

# function Get-FocusPulseTaskPreview {
#     param(
#         [pscustomobject]$Item,
#         [int]$MaxLength = 800
#     )

#     if (-not $Item) {
#         return 'No task selected.'
#     }

#     $preview = $Item.Body.Trim()
#     if ([string]::IsNullOrWhiteSpace($preview)) {
#         return '(empty task)'
#     }

#     if ($preview.Length -gt $MaxLength) {
#         return $preview.Substring(0, $MaxLength) + '...'
#     }

#     return $preview
# }

function Get-FocusPulseTaskItems {
    param([string]$Folder)

    if (-not (Test-Path -LiteralPath $Folder)) {
        New-Item -ItemType Directory -Force -Path $Folder | Out-Null
    }

    Get-ChildItem -LiteralPath $Folder -Filter '*.txt' -File |
        ForEach-Object { Get-FocusPulseTaskItem -Path $_.FullName } |
        Where-Object { $null -ne $_ } |
        Sort-Object DueAt, Path
}

function Save-FocusPulseTaskItem {
    param(
        [string]$Folder,
        [pscustomobject]$Item,
        [string]$Body,
        [datetime]$DueAt
    )

    $resolved = Get-FocusPulseUniqueTaskPath -Folder $Folder -DueAt $DueAt

    if ($Item -and $Item.Path -and (Test-Path -LiteralPath $Item.Path)) {
        Set-FocusPulseTaskBody -Path $Item.Path -Body $Body

        if ($Item.Path -ne $resolved.Path) {
            Move-Item -LiteralPath $Item.Path -Destination $resolved.Path
        }
    }
    else {
        Set-FocusPulseTaskBody -Path $resolved.Path -Body $Body
    }

    Get-FocusPulseTaskItem -Path $resolved.Path
}

function New-FocusPulseTaskItem {
    param(
        [string]$Folder,
        [datetime]$DueAt,
        [string]$Body = ''
    )

    $resolved = Get-FocusPulseUniqueTaskPath -Folder $Folder -DueAt $DueAt
    Set-FocusPulseTaskBody -Path $resolved.Path -Body $Body
    Get-FocusPulseTaskItem -Path $resolved.Path
}

function Update-FocusPulseCounters {
    $script:State.StatusText.Text = 'Used: {0} | Wasted: {1}' -f (
        Format-FocusPulseDuration -Duration $script:State.UsedCounter),
        (Format-FocusPulseDuration -Duration $script:State.WastedCounter)
}

function Update-FocusPulseSelectionPreview {
    if ($script:State.CurrentTask) {
        $script:State.StatusText.Text = 'Selected: ' + (Format-FocusPulseDate $script:State.CurrentTask.DueAt)
    }
    else {
        $script:State.StatusText.Text = 'Ready'
    }
}

function Refresh-FocusPulseTaskList {
    param([string]$SelectPath)

    $script:State.Tasks = @(Get-FocusPulseTaskItems -Folder $script:State.TaskFolder)
    $script:State.TaskList.ItemsSource = $script:State.Tasks

    $selection = $null
    if ($SelectPath) {
        $selection = $script:State.Tasks | Where-Object Path -eq $SelectPath | Select-Object -First 1
    }
    elseif ($script:State.CurrentTaskPath) {
        $selection = $script:State.Tasks | Where-Object Path -eq $script:State.CurrentTaskPath | Select-Object -First 1
    }
    elseif ($script:State.Tasks.Count -gt 0) {
        $selection = $script:State.Tasks[0]
    }

    if ($selection) {
        $script:State.TaskList.SelectedItem = $selection
        Load-FocusPulseTaskSelection -Item $selection
    }
    else {
        $script:State.CurrentTask = $null
        $script:State.CurrentTaskPath = $null
        Update-FocusPulseSelectionPreview
    }
}

function Load-FocusPulseTaskSelection {
    param($Item)

    if (-not $Item) {
        return
    }

    $script:State.CurrentTask = $Item
    $script:State.CurrentTaskPath = $Item.Path
    Update-FocusPulseSelectionPreview
}

function Sync-FocusPulseCurrentTask {
    if ($script:State.CurrentTask -and (Test-Path -LiteralPath $script:State.CurrentTaskPath)) {
        $script:State.CurrentTask = Get-FocusPulseTaskItem -Path $script:State.CurrentTaskPath
    }
}

function Invoke-FocusPulseReminder {
    if ($script:State.IsWorking) {
        $elapsed = (Get-Date) - $script:State.StartedAt
        $script:State.UsedCounter = $script:State.UsedCounter.Add($elapsed)
        $message = 'Well done, time''s up! Go back to your task list to reschedule this task.'
    }
    else {
        $script:State.WastedCounter = $script:State.WastedCounter.Add($script:State.CurrentInterval)
        $message = 'Too bad, you''re slacking off! Go to your task list and do the next task.'
    }

    [System.Windows.MessageBox]::Show(
        $message + "`r`n`r`nSo far:`r`nTime Used: " + (Format-FocusPulseDuration -Duration $script:State.UsedCounter) + "`r`nTime Wasted: " + (Format-FocusPulseDuration -Duration $script:State.WastedCounter),
        'To-Do Reminder',
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    ) | Out-Null

    $script:State.IsWorking = $false
    $script:State.StartedAt = $null
}

function Start-FocusPulseTimer {
    param([int]$Minutes)

    if ($script:State.ReminderTimer) {
        $script:State.ReminderTimer.Stop()
    }

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMinutes($Minutes)
    $timer.add_Tick({ Invoke-FocusPulseReminder })
    $timer.Start()

    $script:State.CurrentInterval = [TimeSpan]::FromMinutes($Minutes)
    $script:State.ReminderTimer = $timer
}

function Stop-FocusPulseTimer {
    if ($script:State.ReminderTimer) {
        $script:State.ReminderTimer.Stop()
        $script:State.ReminderTimer = $null
    }
}

function Invoke-FocusPulseStartWork {
    param([int]$Minutes)

    if ($script:State.IsWorking -and $script:State.StartedAt) {
        $script:State.UsedCounter = $script:State.UsedCounter.Add((Get-Date) - $script:State.StartedAt)
        [System.Windows.MessageBox]::Show(
            'Well done!' + "`r`n`r`nSo far:`r`nTime Used: " + (Format-FocusPulseDuration -Duration $script:State.UsedCounter) + "`r`nTime Wasted: " + (Format-FocusPulseDuration -Duration $script:State.WastedCounter),
            'To-Do Reminder',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
    }

    $script:State.IsWorking = $true
    $script:State.StartedAt = Get-Date
    Start-FocusPulseTimer -Minutes $Minutes

    $script:State.StatusText.Text = 'Work until ' + $script:State.StartedAt.AddMinutes($Minutes).ToString('HH:mm')
}

function Save-FocusPulseSelectedTask {
    if (-not $script:State.CurrentTask) {
        return $null
    }

    $saved = Save-FocusPulseTaskItem -Folder $script:State.TaskFolder -Item $script:State.CurrentTask -Body $script:State.CurrentTask.Body -DueAt $script:State.CurrentTask.DueAt
    $script:State.CurrentTask = $saved
    $script:State.CurrentTaskPath = $saved.Path
    $saved
}

function Invoke-FocusPulseRescheduleSelectedTask {
    param([datetime]$NewDueAt)

    if (-not $script:State.CurrentTask) {
        [System.Windows.MessageBox]::Show('Select a task first.', 'FocusPulse', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    $saved = Save-FocusPulseTaskItem -Folder $script:State.TaskFolder -Item $script:State.CurrentTask -Body $script:State.CurrentTask.Body -DueAt $NewDueAt
    $script:State.CurrentTask = $saved
    $script:State.CurrentTaskPath = $saved.Path
    Refresh-FocusPulseTaskList -SelectPath $saved.Path
}

function Invoke-FocusPulseNewTask {
    $dueAt = Show-FocusPulseDueTimeDialog -InitialDueAt (Get-Date).AddMinutes([double]$script:State.CurrentInterval.TotalMinutes) -TaskTitle 'New task due time'
    if (-not $dueAt) {
        return
    }

    $content = Show-FocusPulseContentDialog -InitialBody '' -DialogTitle 'New task content' -AllowRightToLeft
    if (-not $content) {
        return
    }

    $newItem = New-FocusPulseTaskItem -Folder $script:State.TaskFolder -DueAt $dueAt -Body $content.Body
    Refresh-FocusPulseTaskList -SelectPath $newItem.Path
}

function Show-FocusPulseFolderPicker {
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = 'Choose the folder that stores FocusPulse task files.'
    $dialog.SelectedPath = $script:State.TaskFolder
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }

    $null
}

function Show-FocusPulseDueTimeDialog {
    param(
        [datetime]$InitialDueAt,
        [string]$TaskTitle = 'Edit due time'
    )

    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Edit due time"
        Width="420"
        Height="220"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        Background="#FF111827"
        Foreground="#FFF3F4F6"
        FontFamily="Segoe UI"
        FontSize="14">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <TextBlock x:Name="TaskLabel" FontSize="16" FontWeight="SemiBold" TextWrapping="Wrap" />
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,16,0,16">
            <DatePicker x:Name="DatePicker" Width="170" Margin="0,0,10,0" />
            <TextBox x:Name="HourBox" Width="48" Text="00" Margin="0,0,6,0" />
            <TextBlock Text=":" VerticalAlignment="Center" Margin="0,0,6,0" />
            <TextBox x:Name="MinuteBox" Width="48" Text="00" />
        </StackPanel>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="OkButton" Content="Save" Width="84" Margin="0,0,8,0" IsDefault="True" />
            <Button x:Name="CancelButton" Content="Cancel" Width="84" IsCancel="True" />
        </StackPanel>
    </Grid>
</Window>
'@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    $window.Owner = $script:State.Window
    $window.FindName('TaskLabel').Text = $TaskTitle
    $window.FindName('DatePicker').SelectedDate = $InitialDueAt.Date
    $window.FindName('HourBox').Text = $InitialDueAt.ToString('HH')
    $window.FindName('MinuteBox').Text = $InitialDueAt.ToString('mm')

    $result = $null
    $window.FindName('OkButton').Add_Click({
        $dateValue = $window.FindName('DatePicker').SelectedDate
        $hour = 0
        $minute = 0
        if (-not $dateValue -or -not [int]::TryParse($window.FindName('HourBox').Text, [ref]$hour) -or -not [int]::TryParse($window.FindName('MinuteBox').Text, [ref]$minute)) {
            [System.Windows.MessageBox]::Show('Enter a valid date and time.', 'FocusPulse', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        if ($hour -lt 0 -or $hour -gt 23 -or $minute -lt 0 -or $minute -gt 59) {
            [System.Windows.MessageBox]::Show('Enter a valid hour (00-23) and minute (00-59).', 'FocusPulse', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        $result = [datetime]::new($dateValue.Value.Year, $dateValue.Value.Month, $dateValue.Value.Day, $hour, $minute, 0)
        $window.DialogResult = $true
        $window.Close()
    })

    if ($window.ShowDialog()) {
        return $result
    }

    $null
}

function Show-FocusPulseContentDialog {
    param(
        [string]$InitialBody,
        [string]$DialogTitle = 'Edit task content',
        [switch]$AllowRightToLeft
    )

    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Edit task content"
    Width="920"
    Height="640"
        WindowStartupLocation="CenterOwner"
        Background="#FF111827"
        Foreground="#FFF3F4F6"
        FontFamily="Segoe UI"
        FontSize="14">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" VerticalAlignment="Center">
            <TextBlock x:Name="TitleText" FontSize="16" FontWeight="SemiBold" Margin="0,0,16,0" />
            <CheckBox x:Name="RtlBox" Content="Right-to-left" VerticalAlignment="Center" />
        </StackPanel>
        <Border Grid.Row="1" Margin="0,12,0,12" Background="#FFF9FAFB" CornerRadius="10" Padding="8">
            <RichTextBox x:Name="Editor" BorderThickness="0" AcceptsReturn="True" AcceptsTab="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" />
        </Border>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="OkButton" Content="Save" Width="84" Margin="0,0,8,0" IsDefault="True" />
            <Button x:Name="CancelButton" Content="Cancel" Width="84" IsCancel="True" />
        </StackPanel>
    </Grid>
</Window>
'@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    $window.Owner = $script:State.Window
    $window.Title = $DialogTitle
    $window.FindName('TitleText').Text = $DialogTitle
    $window.FindName('RtlBox').IsChecked = $true

    $editor = $window.FindName('Editor')
    $range = [System.Windows.Documents.TextRange]::new($editor.Document.ContentStart, $editor.Document.ContentEnd)
    $range.Text = $InitialBody

    $applyDirection = {
        $flow = if ($window.FindName('RtlBox').IsChecked) {
            [System.Windows.FlowDirection]::RightToLeft
        }
        else {
            [System.Windows.FlowDirection]::LeftToRight
        }

        $editor.FlowDirection = $flow
        $editor.Document.FlowDirection = $flow
    }
    & $applyDirection
    $window.FindName('RtlBox').Add_Click($applyDirection)

    $window.FlowDirection = [System.Windows.FlowDirection]::RightToLeft
    $editor.Document.PageWidth = 820

    $result = $null
    $window.FindName('OkButton').Add_Click({
        $range = [System.Windows.Documents.TextRange]::new($editor.Document.ContentStart, $editor.Document.ContentEnd)
        $result = [pscustomobject]@{
            Body = $range.Text.TrimEnd("`r", "`n")
            FlowDirection = $editor.FlowDirection
        }
        $window.DialogResult = $true
        $window.Close()
    })

    if ($window.ShowDialog()) {
        return $result
    }

    $null
}

function Initialize-FocusPulseWindow {
    [xml]$xaml = @'
<!--<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FocusPulse"
        Width="1380"
        Height="820"
        WindowStartupLocation="CenterScreen"
        Background="#FF111827"
        Foreground="#FFF3F4F6"
        FontFamily="Segoe UI"
        FontSize="14">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#FF1F2937" CornerRadius="12" Padding="14" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Task folder" VerticalAlignment="Center" Margin="0,0,10,0" />
                    <TextBox x:Name="FolderText" Width="620" Margin="0,0,10,0" Background="#FFF9FAFB" Foreground="#FF111827" />
                    <Button x:Name="BrowseButton" Content="Browse" Width="90" Margin="0,0,16,0" />
                    <TextBlock Text="Interval (min)" VerticalAlignment="Center" Margin="0,0,10,0" />
                    <TextBox x:Name="IntervalText" Width="72" Text="10" Background="#FFF9FAFB" Foreground="#FF111827" />
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="NewButton" Content="New Task" Width="92" Margin="0,0,8,0" />
                    <Button x:Name="RefreshButton" Content="Refresh" Width="82" Margin="0,0,8,0" />
                    <Button x:Name="StartButton" Content="Start" Width="80" Margin="0,0,8,0" />
                    <Button x:Name="StopButton" Content="Stop" Width="80" />
                </StackPanel>
            </Grid>
        </Border>

        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="820" />
                <ColumnDefinition Width="16" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Background="#FF1F2937" CornerRadius="12" Padding="12">
                <DockPanel>
                    <TextBlock DockPanel.Dock="Top" Text="Tasks" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,10" />
                    <ListBox x:Name="TaskList" DisplayMemberPath="Display" Background="#FFF9FAFB" Foreground="#FF111827" BorderThickness="0" />
                </DockPanel>
            </Border>

            <Border Grid.Column="2" Background="#FF1F2937" CornerRadius="12" Padding="12">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0" Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="Selected task" FontSize="18" FontWeight="SemiBold" Margin="0,0,14,0" />
                        <TextBlock x:Name="DueText" Text="Due: -" VerticalAlignment="Center" Margin="0,0,16,0" />
                    </StackPanel>

                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,12,0,12">
                        <Button x:Name="EditBodyButton" Content="Edit Contents" Width="112" Margin="0,0,8,0" />
                        <Button x:Name="EditDueButton" Content="Edit Due Time" Width="116" Margin="0,0,8,0" />
                        <Button x:Name="RescheduleMinutesButton" Content="+ Interval" Width="92" Margin="0,0,8,0" />
                        <Button x:Name="RescheduleDayButton" Content="+ Day" Width="72" Margin="0,0,8,0" />
                        <Button x:Name="RescheduleWeekButton" Content="+ Week" Width="74" Margin="0,0,8,0" />
                    </StackPanel>

                    <Border Grid.Row="2" Background="#FFF9FAFB" CornerRadius="10" Padding="12">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <TextBlock x:Name="PreviewText" Foreground="#FF111827" TextWrapping="Wrap" FontSize="15" />
                        </ScrollViewer>
                    </Border>
                </Grid>
            </Border>
        </Grid>

        <Border Grid.Row="2" Background="#FF1F2937" CornerRadius="12" Padding="12" Margin="0,12,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="StatusText" Grid.Column="0" Text="Ready" VerticalAlignment="Center" />
                <TextBlock Grid.Column="1" Text="FocusPulse" Opacity="0.75" VerticalAlignment="Center" />
            </Grid>
        </Border>
    </Grid>
</Window>-->
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FocusPulse"
        Width="950"
        Height="820"
        WindowStartupLocation="CenterScreen"
        Background="#FF111827"
        Foreground="#FFF3F4F6"
        FontFamily="Segoe UI"
        FontSize="14">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#FF1F2937" CornerRadius="12" Padding="14" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Task folder" VerticalAlignment="Center" Margin="0,0,10,0" />
                    <TextBox x:Name="FolderText" Width="320" Margin="0,0,10,0" Background="#FFF9FAFB" Foreground="#FF111827" />
                    <Button x:Name="BrowseButton" Content="Browse" Width="90" Margin="0,0,16,0" />
                    <TextBlock Text="Interval (min)" VerticalAlignment="Center" Margin="0,0,10,0" />
                    <TextBox x:Name="IntervalText" Width="72" Text="10" Background="#FFF9FAFB" Foreground="#FF111827" />
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="NewButton" Content="New Task" Width="92" Margin="0,0,8,0" />
                    <Button x:Name="RefreshButton" Content="Refresh" Width="82" Margin="0,0,8,0" />
                    <Button x:Name="StartButton" Content="Start" Width="80" Margin="0,0,8,0" />
                    <Button x:Name="StopButton" Content="Stop" Width="80" />
                </StackPanel>
            </Grid>
        </Border>

        <Border Grid.Row="1" Background="#FF1F2937" CornerRadius="12" Padding="12">
            <DockPanel>
                <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto" />
                        <ColumnDefinition Width="*" />
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Grid.Column="0" Text="Tasks" FontSize="18" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,20,0"/>
                    
                    <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center">
                        <Button x:Name="EditBodyButton" Content="Edit Contents" Width="112" Margin="0,0,8,0" />
                        <Button x:Name="EditDueButton" Content="Edit Due Time" Width="116" Margin="0,0,8,0" />
                        <Button x:Name="RescheduleMinutesButton" Content="+ Interval" Width="92" Margin="0,0,8,0" />
                        <Button x:Name="RescheduleDayButton" Content="+ Day" Width="72" Margin="0,0,8,0" />
                        <Button x:Name="RescheduleWeekButton" Content="+ Week" Width="74" Margin="0,0,8,0" />
                    </StackPanel>
                </Grid>

                <ListBox x:Name="TaskList" 
                         Background="#FFF9FAFB" 
                         Foreground="#FF111827" 
                         BorderThickness="0" 
                         HorizontalContentAlignment="Stretch"
                         ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                    <ListBox.ItemTemplate>
                        <DataTemplate>
                            <TextBlock Text="{Binding Display}" 
                                       TextWrapping="Wrap" 
                                       FlowDirection="RightToLeft" 
                                       Padding="4"/>
                        </DataTemplate>
                    </ListBox.ItemTemplate>
                </ListBox>
            </DockPanel>
        </Border>

        <Border Grid.Row="2" Background="#FF1F2937" CornerRadius="12" Padding="12" Margin="0,12,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="StatusText" Grid.Column="0" Text="Ready" VerticalAlignment="Center" />
                <TextBlock Grid.Column="1" Text="FocusPulse" Opacity="0.75" VerticalAlignment="Center" />
            </Grid>
        </Border>
    </Grid>
</Window>
'@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    $script:State.Window = $window
    $script:State.TaskList = $window.FindName('TaskList')
    $script:State.PreviewText = $window.FindName('PreviewText')
    $script:State.DueText = $window.FindName('DueText')
    $script:State.StatusText = $window.FindName('StatusText')
    $script:State.FolderText = $window.FindName('FolderText')
    $script:State.IntervalText = $window.FindName('IntervalText')

    $window.FindName('BrowseButton').Add_Click({
        $picked = Show-FocusPulseFolderPicker
        if ($picked) {
            $script:State.TaskFolder = $picked
            $script:State.FolderText.Text = $picked
            Refresh-FocusPulseTaskList -SelectPath $script:State.CurrentTaskPath
        }
    })

    $window.FindName('RefreshButton').Add_Click({
        Refresh-FocusPulseTaskList -SelectPath $script:State.CurrentTaskPath
    })

    $window.FindName('NewButton').Add_Click({
        Invoke-FocusPulseNewTask
    })

    $window.FindName('EditBodyButton').Add_Click({
        if (-not $script:State.CurrentTask) {
            [System.Windows.MessageBox]::Show('Select a task first.', 'FocusPulse', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        $content = Show-FocusPulseContentDialog -InitialBody $script:State.CurrentTask.Body -DialogTitle 'Edit task content' -AllowRightToLeft
        if ($content) {
            $script:State.CurrentTask = Save-FocusPulseTaskItem -Folder $script:State.TaskFolder -Item $script:State.CurrentTask -Body $content.Body -DueAt $script:State.CurrentTask.DueAt
            $script:State.CurrentTaskPath = $script:State.CurrentTask.Path
            Refresh-FocusPulseTaskList -SelectPath $script:State.CurrentTaskPath
        }
    })

    $window.FindName('EditDueButton').Add_Click({
        if (-not $script:State.CurrentTask) {
            [System.Windows.MessageBox]::Show('Select a task first.', 'FocusPulse', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        $dueAt = Show-FocusPulseDueTimeDialog -InitialDueAt $script:State.CurrentTask.DueAt -TaskTitle 'Edit due time'
        if ($dueAt) {
            Invoke-FocusPulseRescheduleSelectedTask -NewDueAt $dueAt
        }
    })

    $window.FindName('StartButton').Add_Click({
        $minutes = 0
        if (-not [int]::TryParse($script:State.IntervalText.Text, [ref]$minutes) -or $minutes -lt 1) {
            [System.Windows.MessageBox]::Show('Enter a positive whole number of minutes.', 'FocusPulse', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        Invoke-FocusPulseStartWork -Minutes $minutes
    })

    $window.FindName('StopButton').Add_Click({
        Stop-FocusPulseTimer
        if ($script:State.IsWorking -and $script:State.StartedAt) {
            $script:State.UsedCounter = $script:State.UsedCounter.Add((Get-Date) - $script:State.StartedAt)
        }

        $script:State.IsWorking = $false
        $script:State.StartedAt = $null
        Update-FocusPulseCounters
    })

    $window.FindName('RescheduleMinutesButton').Add_Click({
        $minutes = 0
        if (-not [int]::TryParse($script:State.IntervalText.Text, [ref]$minutes) -or $minutes -lt 1) {
            [System.Windows.MessageBox]::Show('Enter a positive whole number of minutes.', 'FocusPulse', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        Invoke-FocusPulseRescheduleSelectedTask -NewDueAt (Get-Date).AddMinutes($minutes)
    })

    $window.FindName('RescheduleDayButton').Add_Click({
        Invoke-FocusPulseRescheduleSelectedTask -NewDueAt (Get-Date).AddDays(1)
    })

    $window.FindName('RescheduleWeekButton').Add_Click({
        Invoke-FocusPulseRescheduleSelectedTask -NewDueAt (Get-Date).AddDays(7)
    })

    $script:State.TaskList.Add_SelectionChanged({
        $selected = $script:State.TaskList.SelectedItem
        if ($selected) {
            Load-FocusPulseTaskSelection -Item $selected
        }
    })

    $window.Add_Closing({
        Stop-FocusPulseTimer
    })

    if (-not (Test-Path -LiteralPath $script:State.TaskFolder)) {
        New-Item -ItemType Directory -Force -Path $script:State.TaskFolder | Out-Null
    }

    $script:State.FolderText.Text = $script:State.TaskFolder
    $script:State.IntervalText.Text = [string]$DefaultIntervalMinutes
    # $script:State.DueText.Text = 'Due: -'

    Refresh-FocusPulseTaskList
    Start-FocusPulseTimer -Minutes $DefaultIntervalMinutes
    Update-FocusPulseCounters

    $window
}

function Invoke-FocusPulseSelfTest {
    $tempFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('focuspulse-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $tempFolder | Out-Null

    try {
        $early = New-FocusPulseTaskItem -Folder $tempFolder -DueAt (Get-Date '2026-06-01T08:15:00') -Body 'beta'
        $late = New-FocusPulseTaskItem -Folder $tempFolder -DueAt (Get-Date '2026-06-01T09:30:00') -Body "alpha`nline2"

        $tasks = @(Get-FocusPulseTaskItems -Folder $tempFolder)
        if ($tasks.Count -ne 2) {
            throw 'Expected two task files.'
        }

        if ($tasks[0].Path -ne $early.Path) {
            throw 'Task sorting failed.'
        }

        $parsed = ConvertFrom-FocusPulseTaskFileName -FileName ([System.IO.Path]::GetFileName($late.Path))
        if (-not $parsed -or $parsed.ToString('yyyy-MM-dd-HH-mm') -ne '2026-06-01-09-30') {
            throw 'Filename parsing failed.'
        }

        $renamed = Save-FocusPulseTaskItem -Folder $tempFolder -Item $late -Body 'alpha edited' -DueAt (Get-Date '2026-06-01T12:45:00')
        if (-not (Test-Path -LiteralPath $renamed.Path)) {
            throw 'Reschedule rename failed.'
        }

        if ((Get-FocusPulseTaskBody -Path $renamed.Path) -ne 'alpha edited') {
            throw 'Task text was not preserved.'
        }

        'Self-test passed.'
    }
    finally {
        Remove-Item -LiteralPath $tempFolder -Recurse -Force
    }
}

if ($SelfTest) {
    Invoke-FocusPulseSelfTest
    return
}

$window = Initialize-FocusPulseWindow
$null = $window.ShowDialog()