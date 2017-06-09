Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("c:\scripts\software.csv", True)

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery _
    ("Select * from Win32_Product")

objTextFile.WriteLine "Caption" & "," & _
    "Description" & "," & "Identifying Number" & "," & _
    "Install Date" & "," & "Install Location" & "," & _
    "Install State" & "," & "Name" & "," & _ 
    "Package Cache" & "," & "SKU Number" & "," & "Vendor" & "," _
        & "Version" 

For Each objSoftware in colSoftware
    objTextFile.WriteLine objSoftware.Caption & "," & _
    objSoftware.Description & "," & _
    objSoftware.IdentifyingNumber & "," & _
    objSoftware.InstallDate2 & "," & _
    objSoftware.InstallLocation & "," & _
    objSoftware.InstallState & "," & _
    objSoftware.Name & "," & _
    objSoftware.PackageCache & "," & _
    objSoftware.SKUNumber & "," & _
    objSoftware.Vendor & "," & _
    objSoftware.Version
Next
objTextFile.Close
	