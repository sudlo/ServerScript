Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = inventory.mdb" 

objRecordSet.Open "SELECT * FROM GeneralProperties" , _
    objConnection, adOpenStatic, adLockOptimistic

objRecordSet.AddNew
objRecordSet("ComputerName") = "atl-ws-01"
objRecordSet("Department") = "Human Resources"
objRecordSet("OSName") = "Microsoft Windows XP Professional"
objRecordSet("OSVersion") = "5.1.2600"
objRecordSet("OSManufacturer") = "Microsoft Corporation"
objRecordSet.Update

objRecordSet.Close
objConnection.Close
