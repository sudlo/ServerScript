Set objShell = WScript.CreateObject("WScript.Shell")
objShell.RegWrite "HKCR\.VBS\ShellNew\FileName","template.vbs"
