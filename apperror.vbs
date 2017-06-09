Set objArgs = WScript.Arguments

strComputer = objArgs(0)

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'Application' and Type = 'error'")

Set RE = New RegExp
RE.IgnoreCase = True
RE.Pattern = objArgs(1)

For Each objEvent in colLoggedEvents

    Line = objEvent.TimeWritten

    If RE.Test(Line) Then  

    Wscript.Echo "Computer Name: " & objEvent.ComputerName
    Wscript.Echo "Event Code: " & objEvent.EventCode
    Wscript.Echo "Message: " & objEvent.Message
    Wscript.Echo "Source Name: " & objEvent.SourceName
    Wscript.Echo "Time: " & objEvent.TimeWritten

  End If

Next

