On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colFloppyDrives = objVM.FloppyDrives
    For Each objDrive in colFloppyDrives
        errReturn = objDrive.AttachHostDrive("A")
    Next
Next
