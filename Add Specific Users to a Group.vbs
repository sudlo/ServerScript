On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
    "SELECT ADsPath FROM 'LDAP://DC=fabrikam,DC=com' " _
        & "WHERE objectCategory='User' AND Title='Accountant'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute

Set objGroup = GetObject _
    ("LDAP://cn=Accountants,ou=NA,dc=fabrikam,dc=com")
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    objGroup.Add(objRecordSet.Fields("ADsPath").Value)
    objRecordSet.MoveNext
Loop
