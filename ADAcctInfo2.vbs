'written by Charles Ross
'searches current domain for user accounts based on firstName &/or lastName
'returns status of multiple account attributes

	Input = InputBox("Enter user's First name" & VbCrLf & "to search for:") 
	If Input = "" Then Input = "*"
	FirstName = Input
	Input = InputBox("Enter user's last name" & VbCrLf & "to search for:") 
	If Input = "" Then Input = "*"
	LastName = Input
' 	WScript.Echo "FirstName = " & firstname
' 	WScript.Echo "LastName = " & lastname

	If FirstName = LastName Then WScript.Quit
	t1 = vbTab
	t2 = t1 & t1	
	t3 = t2 & t1
	t4 = t2 & t2
	Const ADS_SCOPE_SUBTREE = 2
	Const ADS_UF_ACCOUNTDISABLE = 2
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	
	objCommand.CommandText = _
	    "SELECT sAMAccountName FROM 'LDAP://dc=na,dc=corp,dc=clorox,dc=com' WHERE objectCategory='user' " & _
 	    "AND givenName='" & Firstname & "' AND sn='" & LastName & "'" '& _
 	    '"AND DisplayName='" & LastName & ",*" & "'" & _
 	    '"OR sn<>'" & "*" & "'"

	Set objRecordSet = objCommand.Execute
	On Error Resume Next	
	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
	    'Wscript.Echo objRecordSet.Fields("sAMAccountName").Value
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		WScript.StdOut.WriteLine VbCrLf & "*******************************************" & vbcrlf
		strname = objRecordSet.Fields("sAMAccountName").Value
		'WScript.Echo "User's sam name is: " & strName	
		strContainer = "cn=users"
		Set objRootDSE = GetObject("LDAP://rootDSE")
	  	Set objUser = GetObject("LDAP://cn=" & strName & "," & strContainer & "," & objRootDSE.Get("defaultNamingContext"))
		WScript.StdOut.WriteLine objUser.Get("displayName") & " - " & objUser.Get("description") & VbCrLf
	'''''''''''''''''''''''''
		Set objUser = GetObject("LDAP://cn=" & objRecordSet.Fields("Name").Value& ",cn=users,dc=na,dc=corp,dc=clorox,dc=com")
		intUAC = objUser.Get("userAccountControl")
		
		If intUAC AND ADS_UF_ACCOUNTDISABLE Then
		    'Wscript.Echo "The account is disabled"
		    status = "Disabled"
		Else
		    'Wscript.Echo "The account is enabled"
		    status = "Active"
		End If
	'''''''''''''''''''''''''
		'wscript.stdout.writeline t1 &  "info: " & objUser.Get("info")
		wscript.stdout.writeline t1 & "status: " & status
		wscript.stdout.writeline t1 & "cn: " & objUser.Get("cn")
		wscript.stdout.writeline t1 & "comment: " & objUser.Get("comment")
		wscript.stdout.writeline t1 & "description: " & objUser.Get("description")
		wscript.stdout.writeline t1 & "displayName: " & objUser.Get("displayName")
		wscript.stdout.writeline t1 & "whenCreated: " & objUser.Get("whenCreated")
		wscript.stdout.writeline t1 & "whenChanged: " & objUser.Get("whenChanged")
		WScript.stdout.writeline t1 & "pwdLastChanged: " & objUser.PasswordLastChanged
		wscript.stdout.writeline t1 & "distinguishedName: " & objUser.Get("distinguishedName")
		wscript.stdout.writeline t1 & "employeeNumber: " & objUser.Get("employeeNumber")
		wscript.stdout.writeline t1 & "homeDirectory: " & objUser.Get("homeDirectory")
		wscript.stdout.writeline t1 & "homeDrive: " & objUser.Get("homeDrive")
		wscript.stdout.writeline t1 & "initial: " & objUser.Get("initials")
		wscript.stdout.writeline t1 & "mail: " & objUser.Get("mail")
		wscript.stdout.writeline t1 & "name: " & objUser.Get("name")
 		wscript.stdout.writeline t1 & "sAMAccountName: " & objUser.Get("sAMAccountName")
		wscript.stdout.writeline t1 & "roomNumber: " & objUser.Get("roomNumber")
		wscript.stdout.writeline t1 & "scriptPath: " & objUser.Get("scriptPath")
		wscript.stdout.writeline t1 & "givenName: " & objUser.Get("givenName")
		wscript.stdout.writeline t1 & "sn: " & objUser.Get("sn")
		wscript.stdout.writeline t1 & "telephoneNumber: " & objUser.Get("telephoneNumber")
		wscript.stdout.writeline t1 & "title: " & objUser.Get("title")
		wscript.stdout.writeline t1 & "objectGUID: " & objUser.Get("objectGUID")
		wscript.stdout.writeline t1 & "userPrincipalName: " & objUser.Get("userPrincipalName")
		'''''''''''''''''''''''''
		arrMemberOf = objUser.GetEx("memberOf")
		If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
	    	WScript.stdout.WriteLine t1 &  "The memberOf attribute is not set."
		Else
	    	WScript.StdOut.WriteLine t1 & "Member of: "
	    	For each Group in arrMemberOf
	        	WScript.StdOut.WriteLine t2 & Group
    		Next
		End If
		'On Error GoTo 0
		objRecordSet.MoveNext
	Loop
