	Option explicit
	Dim strGrpName,arrMemberOf,strMember,role,clear
	Dim objConnection,objCommand,objRecordSet,objGroup,objUser
	Const ADS_PROPERTY_CLEAR = 1
	
	clear = "n"
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	'Search for groups with wildcard "ccm*:\"
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = _
		"<LDAP://dc=na,dc=corp,dc=clorox,dc=com>;" & _
		"(&(objectCategory=Group)(cn=ccm*));" & "Name"
	Set objRecordSet = objCommand.Execute
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Process found groups
	While Not objRecordSet.EOF
		strGrpName = objRecordSet.Fields("Name")
		Call CaseRole ' subroutime
		WScript.stdout.Write "Group: '" & strGrpName
	    WScript.echo "'." & vbTab & "Group's assigned CCM Role: '" & role & "'."
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	Query each found group one-at-a-time / identify group members
		'On Error Resume Next
		Set objGroup = GetObject("LDAP://cn=" & strGrpName & ",cn=Users,dc=na,dc=corp,dc=clorox,dc=com")
		objGroup.GetInfo
		arrMemberOf = objGroup.GetEx("member")
		WScript.Echo vbTab & "Members:"
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	Process each group member one-at-a-Time
		If clear = "n" Then
			For Each strMember in arrMemberOf
				'WScript.Echo "strMember is: " & strMember
 				'On Error Resume Next
 				Set objUser = GetObject("LDAP://" & strMember)
 			    WScript.echo vbTab & vbTab & objUser.cn
				WScript.Echo vbTab & vbTab & vbTab &  "CN: " & objUser.cn 
				WScript.Echo vbTab & vbTab & vbTab & objUser.cn &  "'s CCM role should be: " & role 
				If objUser.info = "" Then
					WScript.Echo vbTab & vbTab & vbTab & objUser.cn & "'s CCM role is: <not set>"
				Else
					WScript.Echo vbTab & vbTab & vbTab & objUser.cn & "'s CCM role is: " & objUser.info
				End If
				
				If lcase(objUser.info) = lcase(role) Then
					WScript.Echo vbTab & vbTab & vbTab & "OK - Correct role confirmed"
				Else
					WScript.Echo vbTab & vbTab & vbTab & ">>> Configuring CCM role..."
					Set objUser = GetObject("LDAP://" & strMember)
					objUser.Put "info" , role
					objUser.SetInfo
				End If
			Next

		ElseIf clear = "y" Then

 			For Each strMember in arrMemberOf
 				On Error Resume Next
				WScript.Echo "strMember is: " & strMember
 				Set objUser = GetObject("LDAP://" & strMember)
				wscript.Echo "Checking 'info' attrib for acct: " & objUser.cn
				WScript.Echo vbTab & vbTab & vbTab &  "CN: " & objUser.cn 
				If objUser.info = "" Then
					WScript.Echo vbTab & vbTab & vbTab & objUser.cn & "'s CCM role was: <not set>"
				Else
					WScript.Echo vbTab & vbTab & vbTab & objUser.cn &   "'s CCM role should was: " & objUser.info 
				End if
	
				objUser.PutEx ADS_PROPERTY_CLEAR, "info", 0
				objUser.SetInfo
		
				If objUser.info = role Then
					WScript.Echo vbTab & vbTab & vbTab &  "User's CCM role is still set to: " & objUser.info
				Else
					WScript.Echo vbTab & vbTab & vbTab &  "User's CCM role is: <not set>"
				End If
			Next
 		Else
 		End If


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		WScript.Echo vbcrlf
	    objRecordSet.MoveNext
	Wend

	Wscript.Echo VbCrLf & "# matching groups found: " & objRecordSet.RecordCount
	objConnection.Close

	Sub procUsers
	End sub


	Sub CaseRole
		Select Case strGrpName
			Case ("ccm_project_users")
				role = "editor"
			Case ("ccm_brand")
				role = "editor"
			Case ("ccm_sales")
				role = "editor"
			Case ("ccm_mfg_planner")
				role = "editor"
			Case ("ccm_consumer_svcs")
				role = "editor"
			Case ("ccm_packaging")
				role = "editor"
			Case ("ccm_product_dev")
				role = "editor"
			Case ("ccm_pserc")
				role = "editor"
			Case ("ccm_legal")
				role = "editor"
			Case ("ccm_project_mgrs")
				role = "editor"
			Case ("ccm_artists")
				role = "editor"
			Case ("ccm_artists_ext")
				role = "editor"
			Case ("ccm_coordinators")
				role = "editor"
			Case ("ccm_separator_print")
				role = "editor"
			Case ("ccm_senior_mgrs")
				role = "editor"
			Case ("ccm_database_mgrs")
				role = "admin"
			Case ("ccm_is_admin")
				role = "master"
			Case Else
				role = "not recognized"
		End Select
	End Sub
	
	