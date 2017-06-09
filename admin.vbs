'Listing All the Local Groups a User Belongs To

'Returns a list of all the local groups on the computer atl-win2k-01 that a user named kenmyer belongs to. 

Set FSO = CreateObject("Scripting.FileSystemObject")
Set Exc = CreateObject("Excel.Application")
Exc.Visible = True
set wb = Exc.Workbooks.Add
set ws = wb.WorkSheets(1)
Exc.Cells(1,1).Value = "hi"
Exc.SaveWorkSpace = "c:\ff.xls"
Exc.Quit
Set Exc = Notihing
Set outFile = FSO.CreateTextFile("Output.txt",8)
Set inFile = FSO.OpenTextFile("List.txt")

Do While Not inFile.AtEndOfStream

strComputer = inFile.ReadLine
Set colGroups = GetObject("WinNT://" & strComputer & "/Administrators")
'colGroups.Filter = Array("group")

outFile.WriteLine "Computer"&vbtab&"Group" &vbtab& "Member of"&vbtab&"IsDisabled"

'For Each objGroup In colGroups
  

   For Each objUser in colGroups.Members
i= strComp(objUser.name, "hpadmin",1)
        If  i=0 Then
            outFile.WriteLine strComputer &vbtab& objUser.name &vbtab& colGroups.Name 
        End If
    Next
    


'Next

Loop
