'Created by: 
'Basheer Ahmed 
'Delivery Lead
'ITO-GCI, Wintel DTS.
'basheer.ahmed@hp.com
'Created date: 10th May, 2007.
'Version 1.0

'Overview
'This check all the server listed in the list.txt file and outputs the status of all auto-running services in the Output.txt file.

'Input
'The input to this script is List.txt file containing the servers each server in one line.

'Output 
'Output.txt is the output of the script contains the status of the Auto-running services.

' List Wins Service Status

' Windows Server 2003 : Yes
' Windows XP : Yes
' Windows 2000 : Yes


Set FSO = CreateObject("Scripting.FileSystemObject")
Set oFso=FSO.CreateTextFile("Auto.txt")

'Set iFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set iFile = FSO.OpenTextFile("List.txt")
 oFso.writeline "[Computer]" &vbtab&"[Service DisplayName]"  & VbTab & "[StartMode]" &Vbtab& "[State]"
On error resume next

Do while Not iFile.AtEndOfStream
strComputer = iFile.ReadLine

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

IF Err.Number <> 0 Then
oFso.writeline strComputer &vbtab& Err.number &vbtab& Err.Description
Err.Clear

Else

Set colRunningServices = objWMIService.ExecQuery("Select * from Win32_Service Where StartMode =  'Auto' ")

For Each objService in colRunningServices 
    oFso.writeline strComputer &vbtab& objService.DisplayName  & VbTab & objService.StartMode &Vbtab& objService.State
Next

End if
Loop