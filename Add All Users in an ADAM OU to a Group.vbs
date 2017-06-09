On Error Resume Next

Set objOU = GetObject("LDAP://localhost:389/ou=Accounting,dc=fabrikam,dc=com")
Set objGroup = GetObject _
    ("LDAP://localhost:389/cn=Accountants,ou=Accounting,dc=fabrikam,dc=com")
objOU.Filter = Array("user")

For Each objUser in objOU
    objGroup.Add objUser.AdsPath
Next
