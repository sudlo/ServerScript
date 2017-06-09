On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colDVDDrives = objVM.DVDROMDrives
    For Each objDrive in colDVDDrives
        errReturn = objDrive.AttachImage _
            ("C:\Virtual Machines\Images\VMAdditions013107.iso")
    Next
Next
