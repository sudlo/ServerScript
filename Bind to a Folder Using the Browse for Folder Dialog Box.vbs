Const WINDOW_HANDLE = 0
Const NO_OPTIONS = 0

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder _
    (WINDOW_HANDLE, "Select a folder:", NO_OPTIONS, "C:\Scripts")       
Set objFolderItem = objFolder.Self

objPath = objFolderItem.Path
objPath = Replace(objPath, "\", "\\")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("Select * from Win32_Directory where name = '" & objPath & "'")

For Each objFile in colFiles
    Wscript.Echo "Readable: " & objFile.Readable
Next
