' Run script from scheduled task with command: wscript.exe and with variables to be the path of the vbs script
Dim shell,command
command = "powershell.exe -nologo -ExecutionPolicy Bypass -File C:\PathToScript\MyScript.ps1"
Set shell = CreateObject("WScript.Shell")
shell.Run command,0