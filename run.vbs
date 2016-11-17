command = "powershell.exe -nologo -command <PATH_TO_SCRIPT>\PoshEarthView.ps1"

set shell = CreateObject("WScript.Shell")

shell.Run command,0
