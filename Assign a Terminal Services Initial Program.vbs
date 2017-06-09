Const STARTUP_PROGRAM = "c:\accounting\invoice.exe"
Const STARTUP_FOLDER = "c:\accounting\fy_2003"
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSEnvironmentSetting")

For Each objItem in colItems
    errResult = objItem.InitialProgram(STARTUP_PROGRAM, STARTUP_FOLDER)
Next
