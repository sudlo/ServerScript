on error resume next

'* Open a text file for recording AD Account Details
set fs=CreateObject("Scripting.FileSystemObject")
set u=fs.OpenTextFile("C:\ADScripts\ADLogs\ADAccounts.csv",2,true)

' Create a heading row in the csv file
u.writeline "Display Name" & " , " & "SAM Accountname" & " , " & _
				"Description" & " , " & "Account Creation" & " , " & " Expiration Date" & " , " & "Account Status" & " , " &_
				"Employee Number" & " , " & "Home Directory" & " , " & _
				"Home Drive" & " , " & "Mail" & " , " & "Telephone" & " , " & "Group Information"
	


set usr=GetObject("LDAP://cn=users,dc=na, dc=corp,dc=Clorox, dc=com")

For each member in usr


	displayname = " "
	samaccountname = " "
    Description = " "
    Account Creation = " "
    Account ExpireDate = " "
    EmployeeNumber = " "
    homeDirectory = " "
    homeDrive = " "
    mail = " "
    telephoneNumber = " "
    status = " "
    strGroups=" "
     
    displayname = member.get("displayname")
    samaccountname = member.get("samaccountname")
    Description=member.Get("description")
	Creation = member.Get("whenCreated")
	
	'Account Expiration Status
	er = member.get("accountExpirationdate") 
	If Err.Number = -2147467259 OR  er = #1/1/1970# Then
    	ExpireDate = "Not set"
	Else
    	ExpireDate = member.AccountExpirationDate
	End If
	
	'Retrieve account status
    intUAC = member.Get("userAccountControl")
		
	If intUAC AND ADS_UF_ACCOUNTDISABLE Then
	       status = "Disabled"
	Else
	       status = "Active"
	End If
	
	EmployeeNumber = member.Get("employeeNumber")
	homeDirectory= member.Get("homeDirectory")
	homeDrive= member.Get("homeDrive")
	mail=member.Get("mail")
	telephoneNumber = member.Get("telephoneNumber")
		
		

	' This code displays the group membership of a user.
	' It avoids infinite loops due to circular group nesting by 
	' keeping track of the groups that have already been seen.
	' ------ SCRIPT CONFIGURATION ------
	strUserDN = member. distinguishedname
	' ------ END CONFIGURATION ---------
     
		set objUser = GetObject("LDAP://" & strUserDN)
		strSpaces = ""
		set dicSeenGroup = CreateObject("Scripting.Dictionary")
				
		DisplayGroups "LDAP://" & strUserDN, strSpaces, dicSeenGroup
				
		u.writeline displayname & " , " & samaccountname & " , " & _
				Description & " , " & Creation & " , " & ExpireDate & " , " & status & " , " &_
				EmployeeNumber & " , " & homeDirectory & " , " & _
				homeDrive & " , " & mail & " , " & telephoneNumber & " , "  & strGroups		
				
Next

u.close

Wscript.Echo "Script Done..."	

Function DisplayGroups ( strObjectADsPath, strSpaces, dicSeenGroup)
     
   set objObject = GetObject(strObjectADsPath)
   strGroups = strGroups & strSpaces & objObject.Name & ";"
   on error resume next ' Doing this to avoid an error when memberOf is empty
   if IsArray( objObject.Get("memberOf") ) then
      colGroups = objObject.Get("memberOf")
   else
      colGroups = Array( objObject.Get("memberOf") )
   end if
   
   for each strGroupDN In colGroups
      if Not dicSeenGroup.Exists(strGroupDN) then
         dicSeenGroup.Add strGroupDN, 1
         DisplayGroups "LDAP://" & strGroupDN, strSpaces & " ", dicSeenGroup
      end if
   next
     
End Function			
