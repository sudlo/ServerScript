	option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	Written/Maintined by EPS/Intel - C. Ross - 6/06
'	Script requires copy of ROBOCOPY.EXE (2003 Resource Kit) in D:\Util\
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim vol, svol, dvol, src, SrcUNC, dest, destUNC, i, j, k, l, objargs
	Dim LogName, LogPath, LogOut, TS, Robo, RoboArg, RoboArgs, exclDirs
	Dim path, objlog, objFileLog, cmdline, objshell, objExecObject
	Dim objStdOut, line, strOut
	Dim paths, dpath

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
	src = "\\10.8.8.4"
	dest = "\\10.8.8.10"
	LogName = "JBS801" & "_" & vol & "_" & ts & ".log"
	LogPath = "d:\util\FAS3020Migration\Logs\Groups\"
	LogOut = LogPath & LogName
	Robo = "d:\util\robocopy.exe "
	path = "_dept$\groups\"
'	WScript.Echo "Logout: " & LogOut
'	Usage syntax:
	if objArgs.Count <> 1 Then
		WScript.stdout.WriteLine "Parameter missing - script is halted - " & ts
		WScript.stdout.WriteLine vbTab & "Usage:  cscript 801DMGroup.vbs argument"
		WScript.Quit
	End If

	Set objLog = CreateObject("Scripting.FileSystemObject")
	If objLog.FileExists(logout) Then
		Set objFileLog = objLog.OpenTextFile(LogOut, forappending)
	Else
		Set objFileLog = objLog.CreateTextFile(LogOut, false)
	End If

	Select case vol
		Case "g1"
			svol = "\4"
			dvol = "\006"
			paths = Array(path)
			RoboArg = " /copyall /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.4\4_dept$\groups\armorall \\10.8.8.4\4_dept$\groups\legal \\10.8.8.4\4_dept$\groups\britaf&a \\10.8.8.4\4_dept$\groups\tax"
		Case "g2"
			svol = "\4"
			dvol = "\008"
			paths = Array(path & "armorall\",path & "legal\",path & "britaf&a\")
			RoboArg = " /copyall /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP"
		Case "g3"
			svol = "\4"
			dvol = "\009"
			paths = array(path & "tax\")
			RoboArg = " /copyall /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP"
		Case "g4"
			svol = "\5"
			dvol = "\007"
			paths = Array(path)
			RoboArg = " /copyall /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.4\5_dept$\groups\ctcgroup \\10.8.8.4\5_dept$\groups\cppcdivs \\10.8.8.4\5_dept$\groups\forecast"
		Case "g5"
			svol = "\5"
			dvol = "\009"
			paths = array(path & "forecast\",path & "cppcdivs\",path & "ctcgroup\")
			RoboArg = " /copyall /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP"
		Case "g6"
			svol = "\7"
			dvol = "\008"
			paths = Array(path)
			RoboArg = " /copyall /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.4\7_dept$\groups\docsys \\10.8.8.10\008_dept$\groups\armorall \\10.8.8.10\008_dept$\groups\legal \\10.8.8.10\008_dept$\groups\britaf&a"
		Case "g7"
			svol = "\7"
			dvol = "\010"
			paths = Array(path & "docsys\")
			RoboArg = " /copyall /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP"
	End Select


	'Setup RoboCopy to copy source to destination - capture output
	Set objShell = WScript.CreateObject("WScript.Shell")
	WScript.stdout.WriteLine "Robo & CMDLine = " & Robo & CMDLine	
	
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
