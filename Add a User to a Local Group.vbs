strComputer = "atl-ws-01"
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group")

Set objUser = GetObject("WinNT://" & strComputer & "/kenmyer,user")
objGroup.Add(objUser.ADsPath)
