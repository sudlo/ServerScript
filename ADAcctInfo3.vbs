'written by Charles Ross
'searches current domain for user accounts based on firstName &/or lastName
'returns status of multiple account attributes

On Error GoTo 0
	Input = InputBox("Enter user's First name" & VbCrLf & "to search for:") 
	If Input = "" Then Input = "*"
	FirstName = Input
	Input = InputBox("Enter user's last name" & VbCrLf & "to search for:") 
	If Input = "" Then Input = "*"
	LastName = Input
	If FirstName = LastName Then WScript.Quit
	t1 = vbTab
	t2 = t1 & t1	
	Const TIMEOUT = 0
	Set objShell = WScript.CreateObject("WScript.Shell")
	Const ADS_SCOPE_SUBTREE = 2
	Const ADS_UF_ACCOUNTDISABLE = 2
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	On Error GoTo 0	
	objCommand.CommandText = _
	    "SELECT sAMAccountName FROM 'LDAP://dc=na,dc=corp,dc=clorox,dc=com' WHERE objectCategory='user' " & _
 	    "AND givenName='" & Firstname & "' AND sn='" & LastName & "'" '& _
 	    '"AND DisplayName='" & LastName & ",*" & "'" & _
 	    '"OR sn<>'" & "*" & "'"

	Set objRecordSet = objCommand.Execute
	On Error Resume Next	
	objRecordSet.MoveFirst
WScript.StdOut.WriteLine "Err = " & vbTab & Err

	If Err<> 0 Then
		On Error GoTo 0
		WScript.StdOut.WriteLine "Running error routine"
		objCommand.CommandText = _
		    "SELECT sAMAccountName FROM 'LDAP://dc=na,dc=corp,dc=clorox,dc=com' WHERE objectCategory='user' " & _
	 	    "AND displayName='" & "*" & LastName & "*" & "' AND displayName='" & "*" & FirstName & "*" & "'"

		Set objRecordSet = objCommand.Execute
		'On Error Resume Next	
		objRecordSet.MoveFirst
	Else
	End If

	Do Until objRecordSet.EOF
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		WScript.StdOut.WriteLine VbCrLf & "*******************************************" & vbcrlf
		strname = objRecordSet.Fields("sAMAccountName").Value


		'WScript.StdOut.WriteLine "ItemCount: " & objRecordSet.Bookmark + 1
		'WScript.StdOut.WriteLine "RecordCount: " & objRecordSet.recordcount
		item = objRecordSet.Bookmark + 1
		count = objRecordSet.recordcount
		'WScript.Echo "User's sam name is: " & strName	
		strContainer = "cn=users"
 		Set objRootDSE = GetObject("LDAP://rootDSE")
 	  	Set objUser = GetObject("LDAP://cn=" & strName & "," & strContainer & "," & objRootDSE.Get("defaultNamingContext"))
'		Set objUser = GetObject("LDAP://cn=" & objRecordSet.Fields("Name").Value& ",cn=users,dc=na,dc=corp,dc=clorox,dc=com")
'WScript.StdOut.WriteLine "UPN: " & objRecordSet.Fields("userPrincipalName").Value
		Set objUser = GetObject("LDAP://cn=" & strname & ",cn=users,dc=na,dc=corp,dc=clorox,dc=com")

		WScript.StdOut.WriteLine objUser.Get("displayName") & " - " & objUser.Get("description") & VbCrLf
		'''''''''''''''''''''''''
		arrMemberOf = objUser.GetEx("memberOf")
		If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
	    	'WScript.stdout.WriteLine t1 &  "The memberOf attribute is not set."
		Else
	    	'WScript.StdOut.WriteLine t1 & "Member of: "
			grps = ""
	    	For each Group in arrMemberOf
				grps = grps & VbCrLf & t2 & group
	        	'WScript.StdOut.WriteLine t2 & Group
    		Next
    		'WScript.StdOut.writeline grps
		End If
		'''''''''''''''''''''''''
		intUAC = objUser.Get("userAccountControl")
		If intUAC AND ADS_UF_ACCOUNTDISABLE Then
		    'Wscript.Echo "The account is disabled"
		    status = "Disabled"
		Else
		    'Wscript.Echo "The account is enabled"
		    status = "Active"
		End If
	'''''''''''''''''''''''''
	    objShell.Popup objRecordSet.Fields("sAMAccountName").Value & " (" & item & " of " & count & " matches)" & VbCrLf & _
		t1 & "displayName: " & objUser.Get("displayName") & VbCrLf & _
		t1 & "status: " & status & VbCrLf & _
		t1 & "givenName: " & objUser.Get("givenName") & VbCrLf & _
		t1 & "sn: " & objUser.Get("sn") & VbCrLf & _
		t1 & "description: " & objUser.Get("description") & VbCrLf & _
		t1 & "whenCreated: " & objUser.Get("whenCreated") & VbCrLf & _
		t1 & "whenChanged: " & objUser.Get("whenChanged") & VbCrLf & _
		t1 & "pwdLastChanged: " & objUser.PasswordLastChanged & VbCrLf & _
		t1 & "name: " & objUser.Get("name") & VbCrLf & _
 		t1 & "sAMAccountName: " & objUser.Get("sAMAccountName") & VbCrLf & _
		t1 & "userPrincipalName: " & objUser.Get("userPrincipalName")& VbCrLf & _
		t1 & "memberOf:" & _
		grps,TIMEOUT,item & " of " & count & " matches for: " & firstName & " " & lastName
		objRecordSet.MoveNext
		
	Loop
' 		t1 & "comment: " & objUser.Get("comment") & VbCrLf,0 & _
' 		t1 & "employeeNumber: " & objUser.Get("employeeNumber") & VbCrLf & _
' 		t1 & "homeDirectory: " & objUser.Get("homeDirectory") & VbCrLf & _
' 		t1 & "homeDrive: " & objUser.Get("homeDrive") & VbCrLf & _
' 		t1 & "initial: " & objUser.Get("initials") & VbCrLf & _
' 		t1 & "mail: " & objUser.Get("mail") & VbCrLf & _
' 		t1 & "roomNumber: " & objUser.Get("roomNumber") & VbCrLf & _
' 		t1 & "scriptPath: " & objUser.Get("scriptPath") & VbCrLf & _
' 		t1 & "telephoneNumber: " & objUser.Get("telephoneNumber") & VbCrLf & _
' 		t1 & "title: " & objUser.Get("title") & VbCrLf & _
' 		t1 & "objectGUID: " & objUser.Get("objectGUID") & VbCrLf & _
'		t1 & "distinguishedName: " & objUser.Get("distinguishedName") & VbCrLf & _
