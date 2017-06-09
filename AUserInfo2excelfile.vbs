Dim ObjWb 
Dim ObjExcel 
Dim x, zz, PasswordExpiry   
Set objRoot = GetObject("LDAP://RootDSE") 
strDNC = objRoot.Get("DefaultNamingContext") 
Set objDomain = GetObject("LDAP://" & strDNC) ' Bind to the top of the Domain using LDAP using ROotDSE 
Call ExcelSetup("Sheet1") ' Sub to make Excel Document 
x = 1 
Call enummembers(objDomain) 
Sub enumMembers(objDomain) 
PasswordExpiry = 60 
On Error Resume Next 

For Each objMember In objDomain ' go through the collection 


If ObjMember.Class = "user" Then ' if not User object, move on. 
x = x +1 ' counter used to increment the cells in Excel 


  objwb.Cells(x, 1).Value = objMember.Class 


SamAccountName = ObjMember.samAccountName 
EmailAddr = objMember.mail 
WhenCreated = ObjMember.WhenCreated 
PasswordLastChanged = Objmember.PasswordLastChanged 
UserAccountControl = objMember.UserAccountControl 
PassAge = DateDiff("d", PasswordLastChanged, Now) 


If objMember.UserAccountControl = 544 Then Status = "PWExpires" 
If objMember.UserAccountControl = 512 Then Status = "PWExpires" 
If objMember.UserAccountControl = 514 Then Status = "AccountDisabled" 
If objMember.UserAccountControl = 32 Then Status = "PWnotRequired" 
If objMember.UserAccountControl = 66048 Then Status = "PWdoesnotExpire" 
If objMember.UserAccountControl = 66080 Then Status = "PWdoesnotExpire" 
If objMember.UserAccountControl = 66082 Then Status = "PWdoesnotExpire" 
If objMember.UserAccountControl = 546 Then Status = "AccountDisabled" 
If objMember.UserAccountControl = 66050 Then Status = "AccountDisabled" 


If PassAge = 39419 Then PassAge = "Unknown" 
If PassAge = 39420 Then PassAge = "Never" 


set objLogon = objMember.Get("lastLogonTimestamp") 
intLogonTime = objLogon.HighPart * (2^32) + objLogon.LowPart 
intLogonTime = intLogonTime / (60 * 10000000) 
intLogonTime = intLogonTime / 1440 
intLogonTime = intLogonTime + #1/1/1601# 


' Write the values to Excel, using the X counter to increment the rows. 


objwb.Cells(x, 1).Value = SamAccountName 
objwb.Cells(x, 2).Value = EmailAddr 
objwb.Cells(x, 3).Value = WhenCreated 
objwb.Cells(x, 4).Value = PasswordLastChanged 
objwb.Cells(x, 5).Value = intLogonTime 
objwb.Cells(x, 6).Value = UserAccountControl 
objwb.Cells(x, 7).Value = Status 
objwb.Cells(x, 8).Value = PassAge 
objwb.Cells(x, 9).Value = ExpireWarning 


' Blank out Variables in case the next object doesn't have a value for the 
property 
SamAccountName = "-" 
EmailAddr = "-" 
WhenCreated = "-" 
PasswordLastChanged = "-" 


For ll = 1 To 20 
Secondary(ll) = "" 
Next 
  End If 


  ' If the AD enumeration runs into an OU object, call the Sub again to 
itinerate 


  If objMember.Class = "organizationalUnit" or OBjMember.Class = "container" Then 
      enumMembers (objMember) 
  End If 
Next 
End Sub 
Sub ExcelSetup(shtName) ' This sub creates an Excel worksheet and adds Column heads to the 1st row 
Set objExcel = CreateObject("Excel.Application") 
Set objwb = objExcel.Workbooks.Add 
Set objwb = objExcel.ActiveWorkbook.Worksheets(shtName) 
Objwb.Name = "All AD Users" ' name the sheet 
objwb.Activate 
objExcel.Visible = True 
objwb.Cells(1, 1).Value = "SamAccountName" 
objwb.Cells(1, 2).Value = "Emailaddress" 
objwb.Cells(1, 3).Value = "WhenCreated" 
objwb.Cells(1, 4).Value = "PaswordLastChanged" 
objwb.Cells(1, 5).Value = "LastLogonTime" 
objwb.Cells(1, 6).Value = "AccControl" 
objwb.Cells(1, 7).Value = "Status" 
objwb.Cells(1, 8).Value = "PassAge" 


End Sub 
MsgBox "Script Completed" ' show that script is complete 

