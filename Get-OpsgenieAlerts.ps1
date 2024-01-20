$Error.Clear()
[Console]::CursorVisible = $false
$Title = "Opsgenie Alerts Dashboard"
Write-Host "`r`n`t $Title - Loading, Please wait . . ." -ForegroundColor Yellow
$Host.UI.RawUI.WindowTitle = "$Title - Temp ID $((Get-Date).ToFileTimeUtc())"

$path = Split-Path $MyInvocation.MyCommand.Path

[string[]]$data = [System.IO.File]::ReadAllLines("$path\Config.ini", [System.Text.Encoding]::ASCII)
$data = $data -notmatch "^(\#)|^(\;)" -match "\S"

$config = @{}

foreach ($line in $data) {
    $obj = $line.Split("=", 2)
    $config.Add($obj[0].Trim(), $obj[1].Trim())
}

$Api = $config.Api
$Address = $config.Address
$Alarm = [string[]]($config.Alarm.Split(",; ") -match "\S")
$Seconds = $config.Refresh
$Limit = $config.Limit
$Timeout = $config.Timeout
$GenieKey = $config.GenieKey
$Query = $config.Query -replace "\s", "+"
$OldReport = @()

$Player = New-Object System.Media.SoundPlayer
$Player.SoundLocation = Join-Path -Path $path -ChildPath "Alarm.wav"

### Required to popup a message box with help menu
Add-Type -AssemblyName PresentationFramework

### Requirements to use .NET clipboard class
Add-Type -AssemblyName PresentationCore

### Detect key stroke
Add-Type -AssemblyName WindowsBase

### Detect mouse click
Add-Type -AssemblyName System.Windows.Forms

### Required to get current active window
Add-Type -MemberDefinition @"
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
"@ -Name Utils -Namespace Win32

### Maximize Window
# Ref1: https://communary.net/2015/10/11/change-the-powershell-console-size-and-state-programmatically/
# Ref2: https://gist.github.com/Nora-Ballard/11240204

$Win32ShowWindowAsync = Add-Type –memberDefinition @"
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name Win32ShowWindowAsync -Namespace Win32Functions –passThru

$MainWindowHandle = (Get-Process | Where-Object { $_.MainWindowTitle -eq "$($Host.UI.RawUI.WindowTitle)" }).MainWindowHandle
$Win32ShowWindowAsync::ShowWindowAsync($MainWindowHandle, 3) | Out-Null

if ($Error) {
    Write-Host "Something went wrong. Send above errors to the dev." -ForegroundColor Red -BackgroundColor Black
    return $null
}

### Some usefull links for future implementation ###
#$url = "https://api.eu.opsgenie.com/v2/alerts/<Opsgenie ID>/recipients"
#$url = "https://api.eu.opsgenie.com/v2/alerts/<Opsgenie ID>/logs"
#$url = "https://api.eu.opsgenie.com/v2/alerts/<Opsgenie ID>/notes"

Function GetHelp {

$message = @"
 - Press F5 to refresh manually.

 - Press F8 to freeze/unfreeze auto refresh.

 - Alt + <First letter of column's name> - Sort by specific comlumn in ascending/descending order.

 - Copy alert URL in the clipboard: From column 'Link' mark a value, e.g. 'id=8', then press right mouse button.

 - Open alert URL in default browser: From column 'Link' mark a value, e.g. 'id=12' and press Enter.

 - Copy full alert info in the clipboard: From column 'Alert' mark a value, e.g. 'og=16', then press right mouse button.

 - Get full alert info in the dashboard: From column 'Alert' mark a value, e.g. 'og=24' and press Enter. Press escape to go back in the dashboard.
"@

    [System.Windows.MessageBox]::Show($message, "Help Menu") | Out-Null
}

Function GetOutput([string]$text, [string]$color) {
    $lines = ($text | Measure-Object -Line).Lines + 2
    $height = $Host.UI.RawUI.WindowSize.Height

    if ($lines -gt $height) { $height = $lines }

    $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size($Host.UI.RawUI.BufferSize.Width, $height)

    Clear-Host
    if ($color) { Write-Host $text -ForegroundColor $color -NoNewline }
    else { Write-Host $text -NoNewline }
}

Function SortReport($key, $order) {
    $SortKey = switch ($key) {
        'S' {"source"}
        'L' {"link"}
        'A' {"alert"}
        'P' {"priority"}
        'M' {"message"}
    }

    $Script:report = $Script:report | Sort-Object -Property $SortKey -Descending:$order
    $Script:TextReport = ($Script:report | Format-Table -Property source, link, alert, priority, message -AutoSize | Out-String).Trim()

    GetOutput $Script:TextReport
    Write-Host $Script:ReportStats -ForegroundColor Yellow -NoNewline
    $Host.UI.RawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, 0)
    [Console]::CursorVisible = $false
}

Function GetAlert($key) {
    $Clipboard = ([System.Windows.Clipboard]::GetText()).Trim()
    $id = ($Script:report | Where-Object {$_.Link -eq $Clipboard}).id
    $og = ($Script:report | Where-Object {$_.Alert -eq $Clipboard}).id

    if ($id) {
        $link = "https://$($Script:Address)/alert/detail/$id/details"
        if ($key -eq "Enter") {
            [System.Windows.Clipboard]::Clear()
            Start-Process -FilePath $link
        }
        else { [System.Windows.Clipboard]::SetText($link) }
    }
    elseif ($og) {
        [System.Windows.Clipboard]::Clear()
        Clear-Host
        Write-Host "`r`n`t Retrieving alert data . . . " -NoNewline -ForegroundColor Yellow

        $Url = "https://$($Script:Api)/v2/alerts/$($og)?identifierType=id"
        $alert = Invoke-RestMethod -Uri $Url -Method Get -Headers @{"Authorization" = "GenieKey $Script:GenieKey"}

        if ($Error) {
            Write-Host "`r`n Press Esc to exit. " -ForegroundColor Yellow -NoNewline
            $Host.UI.RawUI.FlushInputBuffer()
            while ($Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode -ne 27) {}

            GetOutput $Script:TextReport
            Write-Host $Script:ReportStats -ForegroundColor Yellow -NoNewline
            $Host.UI.RawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, 0)
            [Console]::CursorVisible = $false

            $Error.Clear()
            return $null
        }

        $LastOccurred = [System.DateTime]::UtcNow - ([DateTime]$alert.data.lastOccurredAt).ToUniversalTime()
        $Created = [System.DateTime]::UtcNow - ([DateTime]$alert.data.createdAt).ToUniversalTime()
        $Updated = [System.DateTime]::UtcNow - ([DateTime]$alert.data.updatedAt).ToUniversalTime()

        $LastOccurred = "$("{0:d1}d {1:d2}h {2:d2}m" -f $LastOccurred.Days, $LastOccurred.Hours, $LastOccurred.Minutes)"
        $Created = "$("{0:d1}d {1:d2}h {2:d2}m" -f $Created.Days, $Created.Hours, $Created.Minutes)"
        $Updated = "$("{0:d1}d {1:d2}h {2:d2}m" -f $Updated.Days, $Updated.Hours, $Updated.Minutes)"

        $stat = $alert | Select-Object -Property `
            @{Label="Link"; Expression={"https://$($Script:Address)/alert/detail/$($_.data.id)/details"}}, `
            @{Label="Occurred"; Expression={$LastOccurred}}, `
            @{Label="Created"; Expression={$Created}}, `
            @{Label="Updated"; Expression={$Updated}}

        $Text = ($stat | Format-List | Out-String).Trim() + "`r`n" + "_" * 64 + "`r`n" * 2
        $Text += ($alert.data | Select-Object * -ExcludeProperty id, tinyId, alias, details | Format-List | Out-String).Trim()
        $Text += "`r`n" + "_" * 64 + "`r`n" * 2 + ($alert.data.details | Format-List | Out-String).Trim()

        if ($key -eq "Enter") {
            GetOutput $Text "Cyan"
            Write-Host "`r`n`r`n`t Press Esc to exit. " -ForegroundColor Yellow -NoNewline
            $Host.UI.RawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, 0)
            [Console]::CursorVisible = $false

            $Host.UI.RawUI.FlushInputBuffer()
            while ($Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode -ne 27) {}
        }
        else { [System.Windows.Clipboard]::SetText($Text) }

        GetOutput $Script:TextReport
        Write-Host $Script:ReportStats -ForegroundColor Yellow -NoNewline
        $Host.UI.RawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, 0)
        [Console]::CursorVisible = $false
    }
    else { return $null }
}

Function GetApiData {
    $Script:report = @()
    $Script:GenieId = 0
    $queries = 0
    $offset = 0

    while ($true) {
        $Url = "https://$($Script:Api)/v2/alerts?query=$Script:Query&offset=$offset&limit=100"
        $data = Invoke-RestMethod -Uri $Url -Method Get -Headers @{"Authorization" = "GenieKey $Script:GenieKey"}
        $Script:report += $data.data
        $queries++

        if ($Error) { $Error.Clear() ; return $null }
        elseif (($data.data.Count -eq 100) -and ($Script:report.Count -lt $Script:Limit)) { $offset += 100 }
        else { break }

        Remove-Variable data
        Start-Sleep -Milliseconds $Script:Timeout
    }

    $Script:report = $Script:report | Sort-Object priority, source, message | Select-Object -Property id, priority, source, message, `
        @{Label="link"; Expression={"id=" + $Script:GenieId}}, `
        @{Label="alert"; Expression={"og=" + $Script:GenieId; $Script:GenieId++}}

    $Script:TextReport = ($Script:report | Format-Table -Property source, link, alert, priority, message -AutoSize | Out-String).Trim()
    $Script:ReportStats = "`r`n`r`n`t Queries: $queries | Records: $($Script:report.Count)"

    GetOutput $Script:TextReport
    Write-Host $Script:ReportStats -ForegroundColor Yellow -NoNewline

    $Host.UI.RawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, 0)
    [Console]::CursorVisible = $false

    $ring = $Script:report | Where-Object {($Script:Alarm -contains $_.Prio) -and ($Script:OldReport.id -notcontains $_.id)}
    if ($ring) { $Script:Player.Play() }

    $Script:OldReport = $Script:report | Where-Object {$Script:Alarm -contains $_.Prio}
}

GetApiData
$Host.UI.RawUI.FlushInputBuffer()

$timer = [System.Diagnostics.Stopwatch]::StartNew()
$timespan = [timespan]::new(0, 0, $Seconds)
$refresh = $true
$SortKey = @("S", "L", "A", "P", "M")
$OldKey = $order = $false

while ($true) {
    $remain = $timespan - $timer.Elapsed

    if ([Win32.Utils]::GetForegroundWindow() -eq $MainWindowHandle) {
        ### Check F1 key stroke
        if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::F1)) {
            while ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::F1)) { Start-Sleep -Milliseconds 12 }
            GetHelp
        }

        ### Check F5 key stroke
        if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::F5)) {
            while ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::F5)) { Start-Sleep -Milliseconds 12 }
            $Host.UI.RawUI.WindowTitle = "$Title - refreshing . . ."
            GetApiData
            $OldKey = $order = $false
            $timer.Restart()
        }

        ### Check F8 key stroke
        if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::F8)) {
            while ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::F8)) { Start-Sleep -Milliseconds 12 }
            $refresh = -not $refresh
        }

        ### Check Enter key stroke
        elseif ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Enter)) {
            while ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Enter)) { Start-Sleep -Milliseconds 12 }
            GetAlert "Enter"
        }

        ### Check right click
        elseif ([System.Windows.Forms.UserControl]::MouseButtons -eq "Right") {
            while ([System.Windows.Forms.UserControl]::MouseButtons -eq "Right") { Start-Sleep -Milliseconds 12 }
            GetAlert "Right"
        }

        ### Check Sort key combination
        elseif ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftAlt) -or [Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::RightAlt)) {
            $stroke = $key = $false
            foreach ($key in $SortKey) {
                $stroke = [Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::$key)
                if ($stroke) { break }
                else { $key = $false }
            }

            if ($key) {
                while ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::$key)) { Start-Sleep -Milliseconds 12 }

                if ($OldKey -eq $key) { $order = -not $order }
                $OldKey = $key
                SortReport $key $order
            }
        }
    }

    if ($refresh) {
        if ($remain.TotalMilliseconds -le 0) {
            $Host.UI.RawUI.WindowTitle = "$Title - refreshing . . ."
            GetApiData
            $OldKey = $order = $false
            $timer.Restart()
        }

        $Host.UI.RawUI.WindowTitle = "$Title - Auto refresh in: $("{0:d1}:{1:d2}" -f $remain.Minutes, $remain.Seconds) | Press F1 for help"
    }
    else {
        $Host.UI.RawUI.WindowTitle = "$Title - Auto refresh freezed | Press F1 for help"
    }

    Start-Sleep -Milliseconds 12
}
