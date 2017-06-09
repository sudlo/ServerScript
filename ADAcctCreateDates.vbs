	On Error Resume Next
	
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
		objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
		objCommand.Properties("Sort On") = "Name"
	
	objCommand.CommandText = _
	    "<LDAP://dc=na,dc=corp,dc=clorox,dc=com>;" & _
	        "(&(objectCategory=User));" & _
	            "Name,displayName,Description,whenCreated;Subtree"
	Set objRecordSet = objCommand.Execute
	
	objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	Err.Clear
	If InStr(objRecordSet.Fields("displayName").Value , ",") then
		TestArray = Split(objRecordSet.Fields("displayName").Value , ",")
		LastName = trim(TestArray(0))
	 	FirstName = trim(TestArray(1))

		If Err = 0 Then
		 	LastName = Split(LastName," ")
			If Err = 0 then
		 	 	WLName = ""
		 	 	For i = LBound(LastName) to UBound(LastName)
		 	    	WLName = WLName+LastName(i)
		 	    	'Wscript.Echo "	LastName: " & LastName(i) & vbTab & i
		 	    	'Wscript.Echo "	WLName: " & WLName
		 		Next
		 		LastName = WLName
			Else
				LastName = trim(TestArray(0))
		 	End If
			Err.Clear		
		 	FirstName = Split(FirstName," ")

			If Err = 0 Then
		 	 	WFName = ""			
		 	 	For i = LBound(FirstName) to UBound(FirstName)
					WFName = WFName + FirstName(i)
			 	    'Wscript.Echo "	WFName: " & WFName
		 		Next
		 		FirstName = WFName
		 	Else
		 		FirstName = trim(TestArray(1))
		 	End If

			DisplayName = LastName & "," & FirstName
		Else

		End If
	ElseIf InStr(objRecordSet.Fields("displayName").Value , " ") Then
		WDName = ""
		TestArray = Split(objRecordSet.Fields("displayName").Value , " ")
 	 	For i = LBound(TestArray) to UBound(TestArray)
 	    	WDName = WDName + TestArray(i) + "^"
 		Next
		If InStr(WDname,objRecordSet.Fields("Name").Value) Then
	 		DisplayName = WDName & " ??? "		
		else
	 		DisplayName = WDName
		End if
	Else
		If IsNull(objRecordSet.Fields("displayName").Value) Then
			DisplayName = "++Null++"
		ElseIf lcase(objRecordSet.Fields("displayName").Value) = LCase(objRecordSet.Fields("Name").Value) then
			DisplayName = "[CHECK]-" & objRecordSet.Fields("displayName").Value
		else
			DisplayName = objRecordSet.Fields("displayName").Value
		End if
	End if


	Result = objRecordSet.Fields("Name").Value & vbTab & DisplayName & vbTab & objRecordSet.Fields("whenCreated").Value
	
'	If InStr(Result,"[") Or InStr(Result,"^") or InStr(Result,"[") or InStr(Result,"?") or InStr(Result,"!") Then
		Wscript.Echo Result
'	End If
    objRecordSet.MoveNext
Loop

'''''''''''''''''''''''''''''''
' Set objUser = GetObject("LDAP://cn=ken myer, ou=Finance, dc=fabrikam, dc=com")
' Wscript.Echo objUser.WhenCreated

' Wscript.Echo "err1 = " & Err.Number & vbTab & Err.Description
' Err.Clear


