
	Const adVarChar = 200
	Const MaxCharacters = 255
	Const ForReading = 1
	Const ForWriting = 2
	Const ADS_SCOPE_SUBTREE = 2

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	Set objFileLog = objFSO.CreateTextFile(".\ADGroupInfo.txt, ForWriting)
	objFileLog.WriteLine Now & VbCrLf & VbCrLf 
 	objFileLog.WriteLine "-------------------------------------------------------------------------------------------------------"

	objFileLog.WriteLine "First attibute shown = ""Name""" & VbCrLf & VbCrLf 
	
	Set datalist = CreateObject("ADOR.Recordset")
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	
	
	
	objCommand.CommandText = "SELECT Name FROM 'LDAP://dc=avagotech,dc=net' " & "WHERE objectCategory='group'and GroupType <0"
	Set objRecordSet = objCommand.Execute
	objRecordSet.MoveFirst

	datalist.Fields.Append "Names", adVarChar, MaxCharacters
	datalist.Open
	Do Until objRecordSet.EOF
		DataList.AddNew
		DataList("Names") = objRecordSet.Fields("Name").Value' & "-" & objRecordSet.Fields("whenCreated").Value
		DataList.Update
	    objRecordSet.MoveNext
	Loop
	DataList.Sort = "Names"
	DataList.MoveFirst

	Do Until DataList.EOF
		DatFile = DataList.Fields.Item("Names")


			On Error Resume Next
			Set objGroup = GetObject ("LDAP://cn=" & datFile & ",dc=avagotech, dc=net")
			If Err.Number = -2147016656 Then
			Else
	   	        objFileLog.WriteLine VbCrLf & "------------------------------------------------------------------------------" & VbCrLf
				strWhenCreated = objGroup.Get("whenCreated")
				strWhenChanged = objGroup.Get("whenChanged")
				strDN = objGroup.Get("distinguishedName")				
				strCN = objGroup.Get("CN")				
				strPreWin = objGroup.Get("sAMAccountName")				
'				
				objFileLog.WriteLine datFile
				objFileLog.WriteLine vbTab & "DN: " & strDN
				objFileLog.WriteLine vbTab & "CN: " & strCN
				objFileLog.WriteLine vbTab & "sAMAccountName: " & strPreWin
				objFileLog.WriteLine vbTab & "whenCreated: " & strWhenCreated & " (Created - GMT)"
				objFileLog.WriteLine vbTab & "whenChanged: " & strWhenChanged & " (Modified - GMT)"
	   	        objFileLog.WriteLine VbCrLf
	   	        objFileLog.WriteLine "***** Group MemberShip Info for Group" & strPrewin & "*****"
					
				'Call the Funcntion DisplayMembers to display members of each group
				DisplayMembers "LDAP://" & strDN	   	        	
	   	        On Error goto 0
			End If 
		DataList.MoveNext ' always the last line before Loop		
	Loop 

	objFileLog.WriteLine VbCrLf & VbCrLf & Now 
	WScript.Quit



Function DisplayMembers ( strGroupADsPath)

	strSpaces  = " "
	set dicSeenGroupMember = CreateObject("Scripting.Dictionary")
   
   set objGroup = GetObject(strGroupADsPath)
   WScript.Echo objGroup
   for each objMember In objGroup.Members
      	objLogFile.WriteLine  strSpaces & objMember.Name & vbTab & objMember.DisplayName & vbTab & objMember.Description
      if objMember.Class = "group" then
         if dicSeenGroupMember.Exists(objMember.ADsPath) then
            objLogFile.WriteLine strSpaces & "   ^ already seen group member " & _
                                     "(stopping to avoid loop)"
         else
            dicSeenGroupMember.Add objMember.ADsPath, 1
            DisplayMembers objMember.ADsPath, strSpaces & " ", _
                           dicSeenGroupMember
         end if
      end if
   next
   
End Function
