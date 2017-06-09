strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Volume Where Name = 'D:\\'")

For Each objItem in colItems
    objItem.AddMountPoint("W:\\Scripts\\")
Next
