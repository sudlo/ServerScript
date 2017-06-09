Set WshNetwork = CreateObject("WScript.Network")

WshNetwork.AddWindowsPrinterConnection "\\PrintServer1\Xerox300"
WshNetwork.SetDefaultPrinter "\\PrintServer1\Xerox300"
