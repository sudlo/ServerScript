Dim arruDispName,uDFN,uDLN
Set CNUsers = GetObject ("LDAP://CN=Users,DC=na,DC=corp,DC=clorox,DC=com")
CNUsers.Filter = Array("user")
For Each User in CNUsers
    SAM = lcase(User.sAMAccountName)
    DispName = lcase(User.displayName)
	If InStr(sam,"admin") And InStr(DispName,"* service account")=0 Then
		WScript.StdOut.writeline "AdminAccount/Display Name is: " & sam & " / " & Dispname
		If InStr(DispName,",") then
			arrDispName = Split(DispName, ",")
			tfn = trim(arrDispName(1))
			xfn = left(tfn,3) 'changed this to check only first 3 char
			tfn = left(tfn,1)

			xln = arrDispName(0)
			tln = left(arrDispName(0),7)

			fn = trim(tfn)
			ln = trim(tln)

			lenln = Len(ln)
			ProbAN = fn & ln
			uDispName = " "
			On Error Resume Next
			Set objUser = GetObject("LDAP://cn=" & ProbAN & ",cn=users,dc=na,dc=corp,dc=clorox,dc=com")

			if Err = 0 Then
  				uDispName = lcase(objUser.displayName)
				If InStr(uDispName,",") Then
					pDN(uDispName) ' call sub 
				End If
			End If

			if Err = 0 and InStr(uDLN,xln) and InStr(uDFN,xfn) Then
				cA(ProbAN) 'call sub
				On Error goto 0
			Else
	 			if lenln => "7" Then
					ProbAN = Left(ProbAN,7)
				Elseif lenln < "7" Then
				Else
				End If
				i = 1
				TempAN = ProbAN
				Do Until i = 10
					Err = 0
					ProbAN = TempAN & i
					On Error Resume next
					Set objUser = GetObject("LDAP://cn=" & ProbAN & ",cn=users,dc=na,dc=corp,dc=clorox,dc=com")
					If Err = 0 Then
	     				uDispName = lcase(objUser.displayName)
						If InStr(uDispName,",") Then
							pDN(uDispName) ' call sub
						End If
     				Else
     					uDispName = " "
     				End if
					If Err = 0 And (InStr(uDLN,xln) And InStr(uDFN,xfn)) Then
						cA (ProbAN) 'call Sub
 						Exit Do
					Else
					End If
					i = i + 1
				Loop

				If Err = -2147016656 Then
					WScript.StdOut.writeline vbTab & ucase(">>> Problem Found - Full Name match could not be found for: ") & sam
					WScript.StdOut.writeblanklines (1)
				End If	
			End If ' Err = 0 and InStr(uDLN,xln) and InStr(uDFN,xfn) Then
  		End If ' If InStr(DispName,",") then
 	End If 'If InStr(sam,"admin") And InStr(DispName,"* service account")=0 Then
On Error goto 0
Next
'**********************************************
Sub cA (ProbAN)
						If objUser.AccountDisabled = FALSE Then
						      WScript.StdOut.writeline vbTab & "okay - matching user account/FullName is VALID & enabled: " & ProbAN & " / " & uDispName
						      WScript.StdOut.writeblanklines (1)
						Else
						      WScript.StdOut.writeline vbTab & UCase(">>> Problem Found - matching user account is VALID, but DISABLED: ") & ProbAN & " / " & uDispName
						      WScript.StdOut.writeblanklines (1)
						End If
End Sub

Sub pDN(uDispName)
					arruDispName = Split(uDispName, ",")
					uDFN = trim(arruDispName(1))
					uDFN = Left(uDFN,3) 'changed this to check only first 3 char
					uDLN = trim(arruDispName(0))
End Sub