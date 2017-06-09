Const ForReading = 1
Const ForAppending = 8
Const ForWriting = 2
Const ReadOnly = 1
Const OverwriteExisting = TRUE


'the following host file captures the information of servers that are not reachable

Set  objErrorFSO = CreateObject("Scripting.FileSystemObject")
Set objErrorFile = objErrorFSO.OpenTextFile("C:\Scripts\Errors.txt", ForWriting, True)

'the following contains the details of Members of Local Admin Groups on each Computer
Set objLAdminFSO = CreateObject("Scripting.FileSystemObject")
Set objLAdminFile = ObjLAdminFSO.OpenTextFile("C:\Scripts\LocalAdminMembers.csv", ForAppending, True)


' the objServerFSO & objServerFile are to meant for reading server names from server list
Set  objServerFSO = CreateObject("Scripting.FileSystemObject")
Set objServerFile = objServerFSO.OpenTextFile("C:\Scripts\servers.txt", ForReading)



Do Until objServerFile.AtEndOfStream 
    strComputer = objServerFile.Readline

    Set objShell = CreateObject("WScript.Shell") 
    strCommand = "%comspec% /c ping -n 3 -w 1000 " & strComputer & ""
    Set objExecObject = objShell.Exec(strCommand)

    Do While Not objExecObject.StdOut.AtEndOfStream
        strText = objExecObject.StdOut.ReadAll()
        If Instr(strText, "Reply") > 0 Then

				Call LocalAdminGroupMembers(strComputer)		
		
        Else
            ObjErrorFile.WriteLine strComputer & " could not be reached."
        End If
    Loop
Loop




objServerFile.Close
objErrorFile.Close
objLAdminFile.Close

WScript.Echo "You are done with script...."


Function LocalAdminGroupMembers(strComputer)

		StrGroup ="Administrators"

	objLAdminFile.WriteLine vbCRLF & strComputer & " , " & "Administrator Group Members" & vbCRLF

		Set objGroup = GetObject("WinNT://" & strComputer & "/" & strGroup & ",group")
		If Err <> 0 Then
			
					
		else
			For Each objMember In objGroup.Members
			
								
				If(Instr (ObjMember.Name, "S-")=0) Then
				
					If (objMember.Class ="User") then
				
					 	objLAdminFile.WriteLine  objMember.Name & " , " & objMember.Class & " , " & objMember.Description & " , " & objMember.FullName
					Else
						objLAdminFile.WriteLine  objMember.Name & " , " & objMember.Class & " , " & objMember.Description
					End If				
				Else 
					  objLAdminFile.WriteLine  objMember.Name & " , " & objMember.Class & " , " & "Could Not find reference to the object"
				End If
			
			Next
		End If

End Function