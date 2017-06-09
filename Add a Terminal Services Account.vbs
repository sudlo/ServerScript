Const GUEST_ACCESS = 0
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSPermissionsSetting")

For Each objItem in colItems
    errResult = objItem.AddAccount("fabrikam\bob", GUEST_ACCESS)
Next
