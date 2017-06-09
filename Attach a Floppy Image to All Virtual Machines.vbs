On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colFloppyDrives = objVM.FloppyDrives
    For Each objDrive in colFloppyDrives
        errReturn = objDrive.AttachImage _
            ("C:\Virtual Machines\Additions\Dos Virtual Machine Additions.vfd")
    Next
Next
