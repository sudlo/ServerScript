	Dim Input
	t1 = vbTab
	t2 = t1 & t1	
	t3 = t2 & t1
	t4 = t2 & t2
	
	Input = InputBox("Enter user's First name" & VbCrLf & "to search for:") 
	If Input = "" Then Input = "*"
	FirstName = Input
	Input = InputBox("Enter user's last name" & VbCrLf & "to search for:") 
	If Input = "" Then Input = "*"
	LastName = Input
	If FirstName = LastName Then WScript.Quit
	
	Const ADS_SCOPE_SUBTREE = 2
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	
	objCommand.CommandText = _
	    "SELECT sAMAccountName FROM 'LDAP://dc=na,dc=corp,dc=clorox,dc=com' WHERE objectCategory='user' " & _
	        "AND givenName='" & Firstname & "' AND sn='" & LastName & "'"
	Set objRecordSet = objCommand.Execute
	
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
 
	'''''''''''''''''''''''''
		On Error Resume next
		WScript.StdOut.WriteLine objUser.Get("displayName") & " - " & objUser.Get("description") & VbCrLf
		strInfo = objUser.Get("info")
		strCN = objUser.Get("cn")
		strcomment = objUser.Get("comment")
		strdescription = objUser.Get("description")
		strdisplayName = objUser.Get("displayName")
		strwhenCreated = objUser.Get("whenCreated")
		strwhenChanged = objUser.Get("whenChanged")
WScript.echo "pwdLastSet is: " & objUser.PasswordLastChanged
		strdistinguishedName = objUser.Get("distinguishedName")
		stremployeeNumber = objUser.Get("employeeNumber")
		strgivenName = objUser.Get("givenName")
		strhomeDirectory = objUser.Get("homeDirectory")
		strhomeDrive = objUser.Get("homeDrive")
		strinitials = objUser.Get("initials")
		strmail = objUser.Get("mail")
		strname = objUser.Get("name")
		strroomNumber = objUser.Get("roomNumber")
		strSAMAccountName = objUser.Get("sAMAccountName")
		strscriptPath = objUser.Get("scriptPath")
		strsn = objUser.Get("sn")
		strtelephoneNumber = objUser.Get("telephoneNumber")
		strtitle = objUser.Get("title")
		strobjectGUID = objUser.Get("objectGUID")
		strUPN = objUser.Get("userPrincipalName")
	'''''''''''''''''''''''''
		'wscript.stdout.writeline t1 &  "info: " & strInfo
		wscript.stdout.writeline t1 &  "cn: " & strCN
		wscript.stdout.writeline t1 &  "comment: " & strcomment
		wscript.stdout.writeline t1 &  "description: " & strdescription
		wscript.stdout.writeline t1 &  "displayName: " & strdisplayName
		wscript.stdout.writeline t1 &  "whenCreated: " & strwhenCreated
		wscript.stdout.writeline t1 &  "whenChanged: " & strwhenChanged
		wscript.stdout.writeline t1 &  "distinguishedName: " & strdistinguishedName
		wscript.stdout.writeline t1 &  "employeeNumber: " & stremployeeNumber
		wscript.stdout.writeline t1 &  "homeDirectory: " & strhomeDirectory
		wscript.stdout.writeline t1 &  "homeDrive: " & strhomeDrive
		wscript.stdout.writeline t1 &  "initial: " & strinitials
		wscript.stdout.writeline t1 &  "mail: " & strmail
		wscript.stdout.writeline t1 &  "name: " & strname
 		wscript.stdout.writeline t1 &  "sAMAccountName: " & strSAMAccountName
		wscript.stdout.writeline t1 &  "roomNumber: " & strroomNumber
		wscript.stdout.writeline t1 &  "scriptPath: " & strscriptPath
		wscript.stdout.writeline t1 &  "givenName: " & strgivenName
		wscript.stdout.writeline t1 &  "sn: " & strsn
		wscript.stdout.writeline t1 &  "telephoneNumber: " & strtelephoneNumber
		wscript.stdout.writeline t1 &  "title: " & strtitle
		wscript.stdout.writeline t1 &  "objectGUID: " & strobjectGUID
		wscript.stdout.writeline t1 &  "userPrincipalName: " & strUPN
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
		On Error GoTo 0
		refresh
		objRecordSet.MoveNext
	Loop

Sub refresh
'WScript.Echo "Reset sub"
		strInfo = "" 
		strCN = ""
		strcomment = ""
		strdescription = ""
		strdisplayName = ""
		strwhenCreated = ""
		strwhenChanged = ""
		strdistinguishedName = ""
		stremployeeNumber = ""
		strgivenName = ""
		strhomeDirectory = ""
		strhomeDrive = ""
		strinitials = ""
		strmail = ""
		strname = ""
		strroomNumber = ""
		strSAMAccountName = ""
		strscriptPath = ""
		strsn = ""
		strtelephoneNumber = ""
		strtitle = ""
		strobjectGUID = ""
		strUPN = ""
End sub


