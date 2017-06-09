

'	'''''''''''''''''''''''''''
	'Generic List User Account-page values
On Error Resume Next

Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
 
WScript.Echo "User Principal Name: " & objUser.userPrincipalName
WScript.Echo "SAM Account Name: " & objUser.sAMAccountName
WScript.Echo "User Workstations: " & objUser.userWorkstations

Set objDomain = GetObject("LDAP://dc=fabrikam,dc=com")
WScript.Echo "Domain controller: " & objDomain.dc



'	'''''''''''''''''''''''''''
	'Generic Mod User Account-page values
Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
objUser.Put "userPrincipalName", "MyerKen@fabrikam.com"
objUser.Put "sAMAccountName", "MyerKen01"
objUser.Put "userWorkstations","wks1,wks2,wks3"
objUser.SetInfo

