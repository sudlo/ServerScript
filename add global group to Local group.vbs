'Created by: Basheer Ahmed
'ITO-GCI
'Wintel-DTS
'Email: basheer.ahmed@hp.com
'Creation date: 19/04/2007

'Overview
'Add the given global group in to the Local computers

'Input:
'Input to this script is Computers.txt file containing the server names per line one server.

'Processing:
'Upon execution of the script the given Global group is added in to all computer respective Local group.

'Output:
'The output of this script is:
'1. The given Global group is added in to all computer respective Local group.
'2. The log of the activity is stored in the AddGlobal2Local.log file
'3. For error in the execution can be viewed by using the command "findstr /i Error > Error.txt" Now open the Error.txt file and check for the errors if any.

Const strDomain = "KPMG"    ' Enter your NetBIOS domain name here
Const strGlobalGroup = "GO-SG HP Support Local Power Users"  ' Enter the domain global group name here
Const strLocalGroup = "Remote Desktop Users"  ' Enter the local group name here
Const inFilename = "Computers.txt"    ' Input file namecontaining list of computers
Const outFilename = "AddGlobal2Local.log" ' Log file namecontaining results of operation

Set iFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set inFile = iFSO.OpenTextFile(inFilename)

Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set outFile = oFSO.CreateTextFile(outFilename, 8)

Set objGlobalGroup = GetObject("WinNT://" & strDomain & "/" & strGlobalGroup & ",group")

If Err.Number <> 0 Then
outFile.writeline Now & vbTab & " Error connecting to " & strDomain & "/" & _
strGlobalGroup & " --- " & Err.Description
   Err.Clear
 Else

Do while Not inFile.AtEndOfStream
strComputerName = inFile.ReadLine

'While Not inFile.AtEndOfStream
  ' arrayAccountNames = Split(inFile.Readline, vbTab, -1, 1)
   '       arrayAccountNames(0) contains the computer account name (to modify)
  ' strComputerName = arrayAccountNames(0)
   '       Connect to the computer's local group
   

	Set objLocalGroup = GetObject("WinNT://" & strComputerName & "/" & strLocalGroup & ",group")

	If Err.Number <> 0 Then
        outFile.writeline Now & vbTab & "Error connecting to " & _
        strComputerName & "/" & strLocalGroup & " --- " & Err.Description
         Err.Clear
   Else
         ' Add the global group to the local group on the computer
         objLocalGroup.Add(objGlobalGroup.ADsPath)
         If Err.Number <> 0 Then
               outFile.writeline Now & vbTab & _
                     "Error adding the global group to the local group on " & _
                     strComputerName & " --- " & Err.Description
               Err.Clear
         Else
               outFile.writeline (Now & vbTab & _
                     "Global group successfully added to local group on " & _
                     strComputerName & ".")
 End If
   End If
'Wend

Loop



End If