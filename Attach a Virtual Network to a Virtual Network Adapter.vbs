On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

Set objNetwork = objVS.FindVirtualNetwork("Internal Network")


Set colNetworkAdapters = objVM.NetworkAdapters
For Each objNetworkAdapter in colNetworkAdapters
    errReturn = objNetworkAdapter.AttachToVirtualNetwork(objNetwork)
Next
