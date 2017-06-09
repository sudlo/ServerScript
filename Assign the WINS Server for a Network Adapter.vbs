On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colNetCards = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objNetCard in colNetCards
    strPrimaryServer = "192.168.1.100"
    strSecondaryServer = "192.168.1.200"
    objNetCard.SetWINSServer strPrimaryServer, strSecondaryServer
Next
