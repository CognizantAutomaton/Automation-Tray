$ErrorActionPreference = "Stop"

#region dependencies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsFormsIntegration
#endregion

#region debug
[Array]$debuggers = @("Visual Studio Code Host","Windows PowerShell ISE Host")
if ($debuggers.Contains($host.Name)) {
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog

    if ($fbd.ShowDialog() -eq "OK") {
        Set-Location -LiteralPath $fbd.SelectedPath
    } else {
        Exit
    }
}
#endregion

#region tray icon
function Invoke-ConvertBase64ToIcon {
    param([string]$base64_icon)
    [byte[]]$byteImage = [convert]::FromBase64String($base64_icon)
    [System.IO.MemoryStream]$ms = New-Object System.IO.MemoryStream
    $ms.Write($byteImage, 0, $byteImage.Count)
    $bmp = New-Object System.Drawing.Bitmap($ms)
    $Hicon = $bmp.GetHicon()
    return ([System.Drawing.Icon]::FromHandle($Hicon))
}

# white folder base64 encoded
$base64_icon = @"
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAA
JcEhZcwAADsMAAA7DAcdvqGQAAABWSURBVDhPY6A14AZiBSA2RsMgMZAcXsAHxMpPnjzZ9B8NgMSAciBDwEALiN
FtAONXr17th+rBAFBDQHoZjKFiJAOoRaMGjBpAHQNwJmUiMDgpUwAYGADfzW1xfNnb7gAAAABJRU5ErkJggg==
"@

$icon = Invoke-ConvertBase64ToIcon $base64_icon
#endregion

#region create the tray object
$app = New-Object System.Windows.Forms.ApplicationContext

$MainTrayItem = New-Object System.Windows.Forms.NotifyIcon -Property @{
    Text = "Automation Tray"
    Icon = $icon
    Visible = $true
}

$CtxMnu = New-Object System.Windows.Forms.ContextMenu
$MainTrayItem.ContextMenu = $CtxMnu
#endregion

#region build context menu from folder structure
# create menu items based on current working folder
Get-ChildItem -Recurse | Select-Object FullName, PsIsContainer, @{
    N="Label"
    E={ [IO.Path]::GetFilenameWithoutExtension($_.Name) }
}, @{
    N="RelativePath"
    E={
        if ($_.PsIsContainer) {
            $_.FullName.Substring((Get-Location).Path.Length + 1)
        } else {
            $_.FullName.Substring(0, $_.FullName.LastIndexOf("\")).Substring((Get-Location).Path.Length + 1)
        }
    }
} | Sort-Object -Property FullName | ForEach-Object {
    $mnuItem = New-Object System.Windows.Forms.MenuItem -Property @{
        Text = if ($_.PsIsContainer) { $_.Label } else { $_.Label.Replace("--","\") }
        Tag = @{ Path = $_.RelativePath; Command = $_.FullName }
    }

    if (-not $_.PsIsContainer) {
        $mnuItem.add_Click({
            [string]$CommandString = ([Hashtable]$this.Tag).Command
            if (Test-Path -LiteralPath $CommandString) {
                Invoke-Item -LiteralPath $CommandString
            } else {
                [System.Windows.Forms.MessageBox]::Show("File not found:`n$CommandString", "That item does not exist", "OK", "Error")
            }
        })
    }

    [Array]$Branch = @($_.RelativePath.Split("\"))
    $node = $CtxMnu
    for ([int]$n = 0; $n -lt $(if ($_.PsIsContainer) { $Branch.Count - 1 } else { $Branch.Count }); $n++) {
        $node = $node.MenuItems | Where-Object { $_.Text -eq $Branch[$n] }
    }

    [void]$node.MenuItems.Add($mnuItem)
}
#endregion

#region Refresh menu item
$mnuRefresh = New-Object System.Windows.Forms.MenuItem -Property @{
    Text = "Refresh"
}
$mnuRefresh.add_Click({
    $MainTrayItem.Visible = $false
    Invoke-Item -LiteralPath $PSCommandPath.Replace(".ps1",".vbs")
    Stop-Process $PID
})
[void]$MainTrayItem.ContextMenu.MenuItems.Add($mnuRefresh)
#endregion

#region Exit menu item
$mnuExit = New-Object System.Windows.Forms.MenuItem -Property @{
    Text = "Exit"
}
$mnuExit.add_Click({
    $MainTrayItem.Visible = $false
    Stop-Process $PID
})
[void]$MainTrayItem.ContextMenu.MenuItems.Add($mnuExit)
#endregion

# keep tray application alive
[void][System.Windows.Forms.Application]::Run($app)
