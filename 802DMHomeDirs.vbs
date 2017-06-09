	option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	Written/Maintined by EPS/Intel - C. Ross - 6/06
'	Script requires copy of ROBOCOPY.EXE (2003 Resource Kit) in D:\Util\
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim vol, svol, dvol, src, SrcUNC, dest, destUNC, i, j, k, l, w, y, n, objargs
	Dim LogName, LogPath, LogOut, TS, Robo, RoboArg, RoboArgs, exclDirs
	Dim path, objlog, objFileLog, cmdline, objshell, objExecObject
	Dim objStdOut, line, strOut, objfso, objfolder, colsubfolders, objsubfolder
	Dim paths, dpath, exclds, excld, sfnames

	Const GroupDigits = True
	Const Default = False
	Const NoDecimals = 0
	Const Decimals = 1
	Const ForReading = 1
	Const ForAppending = 8
	Const ForWriting = 2

	Set objArgs = WScript.Arguments
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	vol = lcase(objArgs(0))
	TS = Year(Now) & MonthName(Month(Now),True) & Day(Now)
	src = "\\10.8.8.2"
	dest = "\\10.8.8.11"
	LogName = "JBS802" & "_" & vol & "_" & ts & ".log"
	LogPath = "d:\util\FAS3020Migration\Logs\HomeDirs\"
	LogOut = LogPath & LogName
	Robo = "d:\util\robocopy.exe "
	path = "_home$\homedirs\"
'	WScript.Echo "Logout: " & LogOut
'	Usage syntax:
	if objArgs.Count <> 1 Then
		WScript.stdout.WriteLine "Parameter missing - script is halted - " & ts
		WScript.stdout.WriteLine vbTab & "Usage:  cscript 802DMGroup.vbs argument"
		WScript.Quit
	End If

	Set objLog = CreateObject("Scripting.FileSystemObject")
	If objLog.FileExists(logout) Then
		Set objFileLog = objLog.OpenTextFile(LogOut, forappending)
	Else
		Set objFileLog = objLog.CreateTextFile(LogOut, false)
	End If

	Select case vol
		Case "h1"
			svol = "\1"
			dvol = "\001"
			w = "n"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.2\1_home$\homedirs\escheric \\10.8.8.2\1_home$\homedirs\bwillard \\10.8.8.2\1_home$\homedirs\jdavidso \\10.8.8.2\1_home$\homedirs\tbruguie"
		Case "h2"
			svol = "\2"
			dvol = "\002"
			w = "n"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.2\2_home$\homedirs\mfoley \\10.8.8.2\2_home$\homedirs\ashenoy \\10.8.8.2\2_home$\homedirs\fhopson \\10.8.8.2\2_home$\homedirs\stomkins \\10.8.8.2\2_home$\homedirs\descudie \\10.8.8.2\2_home$\homedirs\rcefalo"
		Case "h3"
			svol = "\3"
			dvol = "\003"
			w = "n"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD"
		Case "h4"
			svol = "\6"
			dvol = "\004"
			w = "y"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.11\004_home$\homedirs\fhopson \\10.8.8.11\004_home$\homedirs\stomkins \\10.8.8.11\004_home$\homedirs\descudie \\10.8.8.11\004_home$\homedirs\rcefalo"
			exclds = Array("k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z")
		Case "h5"
			svol = "\6"
			dvol = "\005"
			w = "y"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.11\005_home$\homedirs\escheric \\10.8.8.11\005_home$\homedirs\bwillard \\10.8.8.11\005_home$\homedirs\mfoley \\10.8.8.11\005_home$\homedirs\ashenoy \\10.8.8.11\005_home$\homedirs\jdavidso \\10.8.8.11\005_home$\homedirs\tbruguie"
			exclds = Array("a","b","c","d","e","f","g","h","i","j")
		Case "h6"
			svol = "\1"
			dvol = "\005"
			w = "n"
			paths = array(path & "escheric\",path & "bwillard\",path & "jdavidso\",path & "tbruguie\")
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD"
		Case "h7"
			svol = "\2"
			dvol = "\005"
			w = "n"
			paths = array(path & "mfoley\",path & "ashenoy\")
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD"
		Case "h8"
			svol = "\2"
			dvol = "\004"
			w = "n"
			paths = array(path & "fhopson\",path & "stomkins\",path & "descudie\",path & "rcefalo\")
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD"
	End Select

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	WScript.Echo "w = " & w
	If w = "y" Then
		Set objFolder = objFSO.GetFolder(src & svol & path)
		Set colSubfolders = objFolder.Subfolders

		For Each objSubfolder in colSubfolders
		    sfnames = lcase(objSubfolder.Name)'
			'WScript.Echo "sfnames: " & sfnames
			For Each excld In exclds
				If left(sfnames,1) = excld Then
					RoboArg = RoboArg & " " & src & svol & path & sfnames
					'WScript.Echo sfnames
				End If
			Next
		Next
	Else
	'WScript.Echo "Didn't Run"
	End If

' WScript.Echo "RoboArg: " & RoboArg
' WScript.Quit

	'Setup RoboCopy to copy source to destination - capture output
	Set objShell = WScript.CreateObject("WScript.Shell")
	
	For Each dpath In paths
		SrcUNC = src & svol & dpath
		DestUNC = dest & dvol & dpath
		CMDLine = SrcUNC & " " & destUNC & RoboArg
		i = 0
		j = 0
		k = 0
		l = 0
		Set objExecObject = objShell.Exec(Robo & CMDLine)
	
		Do Until objExecObject.StdOut.AtEndOfStream
			strOut = objExecObject.StdOut.ReadLine
			WScript.stdout.WriteLine strOut
		   	objFileLog.WriteLine strOut
		   	If InStr(strOut,"ERROR 5 (0x00000005)") Then
		   		i = (i + 1)
		   	ElseIf InStr(strOut,"ERROR 1338 (0x0000053A)") Then
		   		j = (j + 1)
		   	elseif InStr(strOut,"ERROR 2 (0x00000002)") Then
		   		k = (k + 1)
		   	elseif InStr(strOut,"(0x0") and Not InStr(strOut,"0x00000002") And Not InStr(strOut,"0x00000005") And Not InStr(strOut,"0x0000053A") Then
		   		l = (l + 1)
		   	Else
		   		'no action
		   	End If
		Loop

	    objFileLog.WriteBlankLines (1)
	    objFileLog.WriteLine "------------------------------------------------------------------------------"
	    objFileLog.WriteLine vbTab & "Robo Copy process: Ended at: " & Now & VbCrLf
	    If i <> 0 Then
	    	objFileLog.WriteLine "!!!  ATTENTION >>> " & i & " '...Access is denied...' ERRORS DETECTED"
	    	wscript.stdout.WriteLine "!!!  ATTENTION >>> " & i & " '...Access is denied...' ERRORS DETECTED"
	    Else
	    	objFileLog.WriteLine vbTab & "Zero 'I' errors detected: " & i
	    End if
	    If j <> 0 Then
	    	objFileLog.WriteLine "!!!  ATTENTION >>> " & j & " '...security descriptor structure is invalid...' ERRORS DETECTED"
	    	wscript.stdout.WriteLine "!!!  ATTENTION >>> " & j & " '...security descriptor structure is invalid...' ERRORS DETECTED"
	    Else
	    	objFileLog.WriteLine vbTab & "Zero 'J' errors detected: " & j
	    End if
	    If k <> 0 Then
	    	objFileLog.WriteLine "!!!  ATTENTION >>> " & k & " '...system cannot find the file specified...' ERRORS DETECTED"
	    	wscript.stdout.WriteLine "!!!  ATTENTION >>> " & k & " ERRORS DETECTED"
	    Else
	    	objFileLog.WriteLine vbTab & "Zero 'K' errors detected: " & k
	    End If
	    If l <> 0 Then
	    	objFileLog.WriteLine "!!!  ATTENTION >>> " & l & " '...NOT YET DEFINED...' ERRORS DETECTED"
	    	wscript.stdout.WriteLine "!!!  ATTENTION >>> " & l & " '...NOT YET DEFINED...' ERRORS DETECTED"
	    Else
	    	objFileLog.WriteLine vbTab & "Zero errors detected"
	    End if
	    objFileLog.WriteLine "------------------------------------------------------------------------------"
	    objFileLog.WriteLine "<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>" & VbCrLf
	Next 
	'Close the Log file
         objFileLog.Close
         WScript.Quit
