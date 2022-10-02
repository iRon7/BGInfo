<#
.Synopsis
    Installs SSO_BGInfo
.Description
    Installs background information window
.Parameter Path
    Defines where the script will be copied to and launched from
.Notes
    Author:    Ronald Bode
    Version:   0.1.11
    Created:   2022-05-10
#>

[CmdletBinding()]param(
    [string]$Path            = '$[par_m_scriptpath]',
    [switch]$Uninstall       = '$[par_m_action]'       -eq 'Uninstall',
    [string]$RefreshInterval = '$[par_m_refreshinterval]',
    [switch]$OnTop           = '$[par_m_ontop]'        -eq 'True',
    [switch]$EnableClose     = '$[par_m_enableclose]'  -eq 'True',
    [switch]$ShowPSWindow    = '$[par_m_showpswindow]' -eq 'True'
)

if ($Path -like '$[[]*[]]') { $Path = 'ProgramFiles:IT4IT\SSO_BGInfo\SSO_BGInfo.ps1' }

function Install([ScriptBlock]$ScriptBlock, $Parameters) {
    $MyName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $TaskName = "Start-$MyName"
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$False -ErrorAction SilentlyContinue
    $Drive, $FilePath = $Path.Split(':', 2)
    if ($Drive -in [Environment+SpecialFolder]::GetNames([Environment+SpecialFolder])) {
        $Path = Join-Path ([environment]::getfolderpath($Drive)) $FilePath
    }
    if ($Uninstall) {
        SSO_WriteLog INFO "Uninstalling $MyName"
        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path }
    }
    else {
        SSO_WriteLog INFO "Installing $MyName"

        $Null = New-Item -ItemType Directory -Force -Path (Split-Path $Path)
        Set-Content -LiteralPath $Path -Value $ScriptBlock

        $ShortPath = (New-Object -ComObject Scripting.FileSystemObject).getfile($Path).ShortPath
        $Argument = @(
            '-ExecutionPolicy ByPass'
            if (!$ShowPSWindow) { '-WindowStyle hidden' }
            "-File $ShortPath"
            $ScriptBlock.Ast.ParamBlock.Parameters.ForEach{
                $Name = $_.Name.VariablePath.UserPath
                $Value = Get-Variable -Name $Name -Scope Script -ValueOnly
                if ($_.StaticType.Name -eq 'SwitchParameter') { if ($Value) { "-$Name" } }
                else { if ($Value -notlike '$[[]*[]]') { "-$Name $Value" } }
            }
        ).Where{ $_ } -Join ' '

        SSO_WriteLog INFO "Command line: $Argument"
        $Users = (Get-LocalGroup).Where{ $_.SID -eq 'S-1-5-32-545' } # Buildin\Users (could be translated)
        $Task = @{
            Description = "Automatically starts $Name at logon"
            TaskPath    = 'IT4IT'
            TaskName    = $TaskName
            Principal   = New-ScheduledTaskPrincipal -GroupId $Users
            Action      = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument $Argument
            Settings    = New-ScheduledTaskSettingsSet -ExecutionTimeLimit '999:00:00' -DontStopOnIdleEnd
            Trigger     = New-ScheduledTaskTrigger -AtLogOn
        }
        $Task.Settings.CimInstanceProperties.Item('MultipleInstances').Value = 3   # 3 corresponds to 'Stop the existing instance'
        Register-ScheduledTask @Task |Start-ScheduledTask
    }
}

Install -ScriptBlock {
<#
.Synopsis
    SSO_BGInfo
.Description
    Dynamically display relevant information about a Windows computer on the desktop's background,
    such as the computer name, IP address, service pack version, and more.
.Parameter RefreshInterval
    Defines the interval between eash dynamic (ScriptBlock) value update
.Parameter OnTop
    Places the information on top (rather than at the background)
.Parameter EnableClose
    For testing purposes: allows to close (Alt-F4) the information Window
.Parameter ShowPSWindow
    For testing purposes: Keeps the PowerShell Window available
.Parameter Install
    Copies the script to the program files folder and runs it at logon
.Notes
    Author:    Ronald Bode
    Version:   0.1.0
    Created:   2022-05-10
#>

[CmdletBinding()]param(
    [TimeSpan]$RefreshInterval = '00:05:00',
    [switch]$OnTop,
    [switch]$EnableClose,
    [switch]$ShowPSWindow
)

$List = [Ordered]@{ # ScriptBlocks will be dynamicaly updated
    'Inrichting:'     = $Env:PI
    1                 = '-'
    'Applicatie:'     = $Env:Applicatie
    'Applicatie Rol:' = $Env:ApplicatieRol
    'Klantenteam:'    = $Env:Klantenteam
    'Omgeving:'       = $Env:Omgeving
    2                 = '-'
    'Logon Server:'   = $Env:LogonServer
    'Pagefile:'       = { @(Get-CimInstance Win32_PageFileUsage).ForEach{ "$($_.Name) ($($_.CurrentUsage)Mb)" } }
    'Host Name:'      = $Env:ComputerName
    'IP Addres:'      = { (Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled="true"').IPAddress.Where{ $_ -Match '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}' } }
    'OS:'             = (Get-CimInstance Win32_OperatingSystem).Caption
    'Boot Time:'      = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    'Free Space:'     = { @(Get-CimInstance Win32_LogicalDisk).ForEach{ "$($_.DeviceID) $(Format-ByteSize $_.FreeSpace)" } -Join ', ' }
}

$Xaml = [xml](@'
<?xml version="1.0" encoding="UTF-8"?>
<Window FontFamily="Calibri, Arial" FontSize="14" TextElement.Foreground="White"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="SSO_BGInfo" WindowStyle="None" ResizeMode="NoResize" WindowState="Maximized" ShowInTaskbar="False" AllowsTransparency="True">
    <Window.Background>
        <SolidColorBrush Opacity="0"/>
    </Window.Background>
    <WrapPanel Margin="40" Orientation="Vertical" HorizontalAlignment="Right" VerticalAlignment="Bottom">
        <Grid VerticalAlignment="Top" HorizontalAlignment="Left">
            <Grid.ColumnDefinitions>
                <ColumnDefinition />
                <ColumnDefinition Width="10" />
                <ColumnDefinition />
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition />
                <RowDefinition />
                <RowDefinition Height="10"/>
                <RowDefinition />
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Grid.Column="0" Grid.RowSpan="99" Background="#164273">
                <Image Grid.Row="0" Grid.Column="0" Grid.RowSpan="99" Margin="6,0,6,20" VerticalAlignment="bottom" Stretch="None" Name="Logo" />
            </Border>
            <TextBlock Grid.Row="0" Grid.Column="2" FontFamily="Cambria, Garamond, Times New Roman" FontSize="16">
                Justiti&#x00eb;le ICT Organisatie
            </TextBlock>
            <TextBlock Grid.Row="1" Grid.Column="2" FontFamily="Cambria, Garamond, Times New Roman" FontSize="16" FontStyle="Italic">
                <TextBlock.LayoutTransform>
                   <ScaleTransform ScaleX=".85"/>
                </TextBlock.LayoutTransform>
                Ministerie van Justitie en Veiligheid
            </TextBlock>
            <Grid Grid.Row="3" Grid.Column="2" Name="List">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition />
                    <ColumnDefinition Width="10" />
                    <ColumnDefinition />
                </Grid.ColumnDefinitions>
            </Grid>
        </Grid>
        <Grid Height="40" />
        <!-- Control bar spacer in case aligned to bottom -->
    </WrapPanel>
</Window>
'@)

SSO_WriteLog INFO "My process id: $PID"
$Script:User32 = Add-Type -Debug:$False -MemberDefinition '
    [DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X,int Y, int cx, int cy, uint uFlags);
' -Name "User32Functions" -namespace User32Functions -PassThru

if (!$ShowPSWindow) { [Void]($User32::ShowWindow((Get-Process -id $PID).MainWindowHandle, 0)) }

$Processes = @(Get-CimInstance win32_process -Filter "name='powershell.exe'").Where{
    (Invoke-CimMethod -InputObject $_ -MethodName GetOwner).User -eq $Env:Username
}
$MyProcess = $Processes.Where{ $_.ProcessId -eq $PID }
if ($MyProcess.CommandLine) {
    $Processes.Where{
        $_.ProcessName -eq $MyProcess.ProcessName -and
        $_.CommandLine -eq $MyProcess.CommandLine -and
        $_.ProcessId -ne $MyProcess.ProcessId
    }.ForEach{
        SSO_WriteLog WARNING "Killing already running process $($_.ProcessId)"
        Stop-Process $_.ProcessId
    }
}

$Logo = '/9j/4AAQSkZJRgABAQEAkACQAAD/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAEAAAAAAAD/7AARRHVja3kAAQAEAAAAZAAA/+EDkmh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8APD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4NCjx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuNS1jMDIxIDc5LjE1NTc3MiwgMjAxNC8wMS8xMy0xOTo0NDowMCAgICAgICAgIj4NCgk8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPg0KCQk8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIgeG1wTU06T3JpZ2luYWxEb2N1bWVudElEPSJ4bXAuZGlkOmNjNGRhNWQ2LTIxZTktNGE1OC04YTlhLWM0Y2VjNTgxMWI3ZiIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDo0QzE3MUE3RTgzNTkxMUU3QjJDOENGMTZFMTkxNzJGNyIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDo0QzE3MUE3RDgzNTkxMUU3QjJDOENGMTZFMTkxNzJGNyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ0MgMjAxNCAoTWFjaW50b3NoKSI+DQoJCQk8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDoxYjMyYzQwOS01OTBmLTQ2ZjEtOWI3Ny0xN2VhMjlmMzZhZjgiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6Y2M0ZGE1ZDYtMjFlOS00YTU4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'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName presentationframework

$Bitmap = [System.Windows.Media.Imaging.BitmapImage]::new()
$Bitmap.BeginInit()
$Bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($Logo)
$Bitmap.EndInit()
$Bitmap.Freeze()

$Script:Shlwapi = Add-Type -MemberDefinition '
    [DllImport("Shlwapi.dll", CharSet=CharSet.Auto)]public static extern int StrFormatByteSize(long fileSize, System.Text.StringBuilder pwszBuff, int cchBuff);
' -Name "ShlwapiFunctions" -namespace ShlwapiFunctions -PassThru

function Script:Format-ByteSize([Long]$Size) {
    $Bytes = New-Object Text.StringBuilder 20
    $Return = $Shlwapi::StrFormatByteSize($Size, $Bytes, $Bytes.Capacity)
    if ($Return) {$Bytes.ToString()}
}

function AddTextBlock([int]$Row, [int]$Column) {}

function Script:Update {
    SSO_WriteLog INFO 'Updating info'
    $Handle = [System.Windows.Interop.WindowInteropHelper]::new($Script:Window).Handle
    $InsertAfter = if ($OnTop) { -1 } else { 1 }
    [Void]$Script:User32::SetWindowPos($Handle, $InsertAfter, 0, 0, 0, 0, 0x53)

    $Script:Dynamic.GetEnumerator().ForEach{
        $Value = &$_.Value.Tag
        if ($_.Value.Text -cne $Value) { # Only update UI when required
            SSO_WriteLog INFO "$($_.Name) $Value"
            $_.Value.Text = $Value
        }
    }
}

SSO_WriteLog INFO "Loading UI"
$Script:Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader($Xaml)))
$Window.FindName('Logo').Source = $Bitmap

SSO_WriteLog INFO "Assigning info"
$Script:Dynamic = @{}
$Grid = $Window.FindName('List')
$Row = 0
$List.GetEnumerator().ForEach{
    SSO_WriteLog INFO "$($_.Name) $($_.Value)"
    $Grid.RowDefinitions.add([System.Windows.Controls.RowDefinition]::new())
    if ($_.Value -Match '^\s*-+\s*') {
        $Separator = [System.Windows.Controls.Separator]::new()
        $Grid.AddChild($Separator)
        [System.Windows.Controls.Grid]::SetRow($Separator, $Row)
        [System.Windows.Controls.Grid]::SetColumn($Separator, 0)
        [System.Windows.Controls.Grid]::SetColumnSpan($Separator, 99)
    }
    else {
        $TextBlock = [System.Windows.Controls.TextBlock]@{ Text = $_.Name; TextAlignment = 'Right' }
        $Grid.AddChild($TextBlock)
        [System.Windows.Controls.Grid]::SetRow($TextBlock, $Row)
        [System.Windows.Controls.Grid]::SetColumn($TextBlock, 0)
        if ($_.Value -is [ScriptBlock]) {
            $TextBlock = [System.Windows.Controls.TextBlock]@{ Tag = $_.Value }
            $Script:Dynamic[$_.Name] = $TextBlock
        }
        else {
            $TextBlock = [System.Windows.Controls.TextBlock]@{ Text = $_.Value }
        }
        $Grid.AddChild($TextBlock)
        [System.Windows.Controls.Grid]::SetRow($TextBlock, $Row)
        [System.Windows.Controls.Grid]::SetColumn($TextBlock, 2)
    }
    $Row++
}

$Script:Window.Add_Loaded({ Update })
$Script:Window.Add_Activated({ Update })
if (!$EnableClose) { $Script:Window.Add_Closing({ $_.Cancel = $True }) }
$Timer = [System.Windows.Forms.Timer]::new()
$Timer.Interval = ([TimeSpan]$RefreshInterval).TotalMilliseconds
$Timer.Add_Tick({ Update })
$Timer.Start()
SSO_WriteLog INFO "Displaying UI"
[Void]$Script:Window.ShowDialog()
$Timer.Stop()
$Timer.Dispose()
}