strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebService")

For Each objItem in colItems
    objItem.AddExtensionFile _
        "C:\WINDOWS\system32\bits_update.dll", False, _
            "BITSEXT", True, "BITS Update"
Next
