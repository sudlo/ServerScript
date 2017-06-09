Set CNUsers = GetObject ("LDAP://CN=Users,DC=na,DC=corp,DC=clorox,DC=com")
CNUsers.Filter = Array("user")

For Each User in CNUsers
     SAM = lcase(User.sAMAccountName)
     DispName = lcase(User.displayName)
	On Error Resume Next
	If InStr(sam,"admin") And InStr(DispName,"* service account")=0 Then
	
		WScript.StdOut.writeline "AdminAccount/Display Name is: " & sam & " / " & Dispname
		If InStr(DispName,",") then
			arrDispName = Split(DispName, ",")
			'WScript.StdOut.writeline "FirstName = " & arrDispName(1)
			'WScript.StdOut.writeline "LastName = " & arrDispName(0)
			'WScript.StdOut.writeline "1stInit = " & left(arrDispName(1),1)
			'WScript.StdOut.writeline "2ndInit = " & left(arrDispName(0),7)
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
			'WScript.Echo "	uDispName = " & uDispName
			'WScript.Echo "	xfn = " & xfn
			'WScript.Echo "	xln = " & xln

			'WScript.StdOut.writeline vbTab & "Probable user acct is: " & """" & ProbAN & """"

			Set objUser = GetObject("LDAP://cn=" & ProbAN & ",cn=users,dc=na,dc=corp,dc=clorox,dc=com")
			'check error returned
			'WScript.Echo "	err = " & err     				
			if Err = 0 Then

   				uDispName = lcase(objUser.displayName)

				'WScript.Echo vbTab & "UserAccount/Display Name = " & ProbAN & " / " & uDispName

				If InStr(uDispName,",") Then
					arruDispName = Split(uDispName, ",")
					uDFN = trim(arruDispName(1))
					uDFN = Left(uDFN,3) 'changed this to check only first 3 char
					uDLN = trim(arruDispName(0))
					'WScript.Echo "	uDFN = " & uDFN
					'WScript.Echo "	uDLN = " & uDLN
				End If
				
				if Err = 0 and InStr(uDLN,xln) and InStr(uDFN,xfn) Then
					'WScript.Echo vbTab & "GOOD - Full DisplayNames match for both: " & ProbAN & " & " & sam
					
					If Err = 0 And (InStr(uDLN,xln) And InStr(uDFN,xfn)) Then
						'WScript.Echo vbTab & "Admin account Full Name matches this User account Full Name: " & ProbAN
						If objUser.AccountDisabled = FALSE Then
						      WScript.StdOut.writeline vbTab & "okay - matching user account is VALID & enabled: " & """" & ProbAN & """"
						      WScript.StdOut.writeblanklines (1)
						Else
						      WScript.StdOut.writeline vbTab & UCase(">>> Problem Found - matching user account is VALID, but DISABLED: ") & """" & ProbAN & """"
						      WScript.StdOut.writeblanklines (1)
						End If
					Else
						Err = 1
					End If
				Else
					'WScript.Echo vbTab & "BAD - DisplayNames DO NOT match for: " & ProbAN & " & " & sam
					'WScript.Echo vbTab & "      Try variations on acct name: "
					'WScript.Echo "	Lenln = " & Lenln
		 			if lenln => "7" Then
						ProbAN = Left(ProbAN,7)
					Elseif lenln < "7" Then
					Else
					End if

					i = 1
					Do Until i = 10
						Err = 0
						ProbANx = ProbAN & i
						'WScript.Echo "	checking ProbANx = " & ProbANx							
						'WScript.Echo "	i = " & i
						Set objUser = GetObject("LDAP://cn=" & ProbANx & ",cn=users,dc=na,dc=corp,dc=clorox,dc=com")
	     				uDispName = lcase(objUser.displayName)
						'WScript.Echo "	err = " & err     				
						If InStr(uDispName,",") Then
							arruDispName = Split(uDispName, ",")
							uDFN = trim(arruDispName(1))
							uDFN = Left(uDFN,3) 'changed this to check only first 3 char
							uDLN = trim(arruDispName(0))
							'WScript.Echo "	New uDFN = " & uDFN
							'WScript.Echo "	New uDLN = " & uDLN
						End If
						If Err = 0 And (InStr(uDLN,xln) And InStr(uDFN,xfn)) Then
							'WScript.Echo vbTab & "Admin account Full Name matches this User account Full Name: " & ProbANx
							If objUser.AccountDisabled = FALSE Then
							      WScript.StdOut.writeline vbTab & "okay - matching user account is VALID & enabled: " & """" & ProbANx & """"
							      WScript.StdOut.writeblanklines (1)
							Else
							      WScript.StdOut.writeline vbTab & UCase(">>> Problem Found - matching user account is VALID, but DISABLED: ") & """" & ProbANx & """"
							      WScript.StdOut.writeblanklines (1)
							End If
							Exit Do
						Else

						End If
						i = i + 1
					Loop

					If Err = -2147016656 Then
						WScript.StdOut.writeline vbTab & ucase(">>> Problem Found - Full Name match could not be found")
						WScript.StdOut.writeline vbTab & UCase("    Compare 'Full Name' fields for both: ") & lcase(sam & " & " & ProbAN)
						WScript.StdOut.writeblanklines (1)
					End If	
				End If


 			Elseif Err <> 0 Then 
				'WScript.Echo "	Search failed for: " & ProbAN
				'WScript.Echo "	Trying variations on name."
				'WScript.Echo "	Lenln = " & Lenln
 				if lenln => "7" Then
					ProbAN = Left(ProbAN,7)
				End If
				i = 1
				Do Until Err = 0 Or i = 10
					Err = 0
					ProbANx = ProbAN & i
					'WScript.Echo "	Trying: " & ProbANx
					Set objUser = GetObject("LDAP://cn=" & ProbANx & ",cn=users,dc=na,dc=corp,dc=clorox,dc=com")
     				uDispName = lcase(objUser.displayName)

					If InStr(uDispName,",") Then
						arruDispName = Split(uDispName, ",")
						uDFN = trim(arruDispName(1))
						uDFN = Left(uDFN,3) 'changed this to check only first 3 char
						uDLN = trim(arruDispName(0))
						'WScript.Echo "	uDFN = " & uDFN
						'WScript.Echo "	uDLN = " & uDLN
					End If
					'WScript.Echo "	Err = " & Err
					If Err = 0 And (InStr(uDFN,xfn) And InStr(uDLN,xln)) Then
						'WScript.Echo vbTab & "Matches 'err = 1' account name = " & ProbANx
						If Err = 0 And (InStr(uDLN,xln) And InStr(uDFN,xfn)) Then
							'WScript.Echo vbTab & "Admin account Full Name matches this User account Full Name: " & ProbANx
							If objUser.AccountDisabled = FALSE Then
							      WScript.StdOut.writeline vbTab & "okay - matching user account is VALID & enabled: " & """" & ProbANx & """"
							      WScript.StdOut.writeblanklines (1)
							Else
							      WScript.StdOut.writeline vbTab & UCase(">>> Problem Found - matching user account is VALID, but DISABLED: ") & """" & ProbANx & """"
							      WScript.StdOut.writeblanklines (1)
							End If
							Exit Do
						Else
							Err = 1
						End If
					Else
						Err = 1
					End If
					i = i + 1
				Loop

				If i > 9 And Err <> 0 then
	 				WScript.StdOut.writeline vbTab & UCase(">>> Problem Found - no match found for admin account: ") & sam & " (Probable user acct: " & """" & ProbAN & """" & ")"
	 				WScript.StdOut.writeblanklines (1)
 				End if
 			End If
' 
 		End If ' If InStr(DispName,",") then
		
	End If 'If InStr(sam,"admin") And InStr(DispName,"* service account")=0 Then

Next

