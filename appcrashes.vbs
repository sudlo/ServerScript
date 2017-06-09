Set objArgs = WScript.Arguments

strComputer = objArgs(0)

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'System' and Type = 'information'")

Set RE = New RegExp
RE.IgnoreCase = True
RE.Pattern = "Exception"

For Each objEvent in colLoggedEvents

	Line = objEvent.Message

    If RE.Test(Line) Then  

 	Wscript.Echo "Time: " & objEvent.TimeWritten
       	Wscript.Echo "Computer Name: " & objEvent.ComputerName
	Wscript.Echo "Event Code: " & objEvent.EventCode
	Wscript.Echo "Message: " & objEvent.Message
	Wscript.Echo "Source Name: " & objEvent.SourceName

    End If
   
Next

