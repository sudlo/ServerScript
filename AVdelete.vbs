'On Error Resume Next

LstFold = "\\jbs801\epsintel\chas\scripting\work\dra\Lists_Logs\"
LstFile = "IATAs.lst"
prog = "C:\program files\netiq\dra\ea.exe /server:jbs61006 "
AVNames = array("Users","Groups","Workstations","Servers","DCs")
	'IATAS = Array("TS1","TS2","TS3")

WScript.Echo "Path = " & LstFold & LstFile
Const ForReading = 1
Set objDictionary = CreateObject("Scripting.Dictionary")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(LstFold & LstFile, ForReading)
i = 0
Do Until objTextFile.AtEndOfStream
	strLine = objTextFile.ReadLine
	objDictionary.Add i, strLine
'		Wscript.Echo strLine'(i)
	i = i + 1
Loop
objTextFile.Close


	For Each objItem in objDictionary
		strIATA = objDictionary.Item(objItem)
		Wscript.Echo "strItem = " & strItem
		For Each AVName In AVNames
			Set objShell = WScript.CreateObject("WScript.Shell")
			CMDLine = "AV " & """" &  strIATA & " NSR " & AVName & """" & " DELETE  mode:b"
			WScript.Echo "Command: " & prog & CMDLine
			Set objExecObject = objShell.Exec(prog & CMDLine)'	execute Robocopy
			Set objStdOut = objExecObject.StdOut'	Wait for process to complete, collect results...
			Do Until objExecObject.StdOut.AtEndOfStream
			    strLine = objExecObject.StdOut.ReadLine()
			    WScript.Echo "strLine = " & strLine
			Loop
		Next
	Next
