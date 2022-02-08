MenuStructure = "C:\Users\USERNAME\Automation\Tray"

Set objShell = CreateObject("Wscript.Shell")

' Path to default version of PowerShell
PowershellPath = """%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"""

' Path to PowerShell script
AppPath = """" & Replace(WScript.ScriptFullName,".vbs",".ps1") & """"

' Current working folder
objShell.CurrentDirectory = MenuStructure

' Run the script
cmd = PowershellPath & " -windowstyle hidden -executionpolicy bypass -noninteractive -file " & AppPath
objShell.Run(cmd),0
