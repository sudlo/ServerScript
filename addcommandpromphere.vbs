' Add "Command Prompt Here" to Windows Explorer

' Windows Server 2003 : Yes
' Windows XP : Yes
' Windows 2000 : Yes
' Windows NT 4.0 : Yes
' Windows 98 : Yes

Set objShell = CreateObject("WScript.Shell")
 
objShell.RegWrite "HKCR\Folder\Shell\MenuText\Command\", _
    "cmd.exe /k cd " & chr(34) & "%1" & chr(34)
objShell.RegWrite "HKCR\Folder\Shell\MenuText\", "Command Prompt Here"