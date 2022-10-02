[CmdletBinding()]param(
    [TimeSpan]$RefreshInterval = '00:05:00',    # Refresh every 5 minutes
    [switch]$OnTop,                             # Put the information on top
    [switch]$EnableClose,                       # Click on a visible item and click <alt>-<F4> to close
    [switch]$ShowPSWindow                       # Don't hide PowerShell window
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName presentationframework

$List = [Ordered]@{ # ScriptBlocks will be dynamicaly updated
    'Description'     = 'PowerShell BGInfo Example'
    1                 = '-'
    'Host Name:'      = $Env:ComputerName
    'Logon Server:'   = $Env:LogonServer
    'Pagefile:'       = { @(Get-CimInstance Win32_PageFileUsage).ForEach{ "$($_.Name) ($($_.CurrentUsage)Mb)" } }
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
    <WrapPanel Margin="40" Orientation="Vertical" HorizontalAlignment="Right" VerticalAlignment="Top">
        <Grid VerticalAlignment="Top" HorizontalAlignment="Left" Name="List">
            <Grid.ColumnDefinitions>
                <ColumnDefinition />
                <ColumnDefinition Width="10" />
                <ColumnDefinition />
            </Grid.ColumnDefinitions>
        </Grid>
        <Grid Height="40" />
        <!-- Control bar spacer in case aligned to bottom -->
    </WrapPanel>
</Window>
'@)

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
        Stop-Process $_.ProcessId
    }
}

$Script:Shlwapi = Add-Type -MemberDefinition '
    [DllImport("Shlwapi.dll", CharSet=CharSet.Auto)]public static extern int StrFormatByteSize(long fileSize, System.Text.StringBuilder pwszBuff, int cchBuff);
' -Name "ShlwapiFunctions" -namespace ShlwapiFunctions -PassThru

function Script:Format-ByteSize([Long]$Size) {
    $Bytes = New-Object Text.StringBuilder 20
    $Return = $Shlwapi::StrFormatByteSize($Size, $Bytes, $Bytes.Capacity)
    if ($Return) {$Bytes.ToString()}
}

function Script:Update {
    $Handle = [System.Windows.Interop.WindowInteropHelper]::new($Script:Window).Handle
    $InsertAfter = if ($OnTop) { -1 } else { 1 }
    [Void]$Script:User32::SetWindowPos($Handle, $InsertAfter, 0, 0, 0, 0, 0x53)

    $Script:Dynamic.GetEnumerator().ForEach{
        $Value = &$_.Value.Tag
        if ($_.Value.Text -cne $Value) { # Only update UI when required
            $_.Value.Text = $Value
        }
    }
}

$Script:Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader($Xaml)))

$Script:Dynamic = @{}
$Grid = $Window.FindName('List')
$Row = 0
$List.GetEnumerator().ForEach{
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
[Void]$Script:Window.ShowDialog()
$Timer.Stop()
$Timer.Dispose()
