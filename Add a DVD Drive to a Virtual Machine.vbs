On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

errReturn = objVM.AddDVDROMDrive(0,1,0)
