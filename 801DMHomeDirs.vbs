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
	src = "\\10.8.8.4"
	dest = "\\10.8.8.10"
	LogName = "JBS801" & "_" & vol & "_" & ts & ".log"
	LogPath = "d:\util\FAS3020Migration\Logs\HomeDirs\"
	LogOut = LogPath & LogName
	Robo = "d:\util\robocopy.exe "
	path = "_home$\homedirs\"

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
		Case "h1"
			svol = "\1"
			dvol = "\001"
			w = "n"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.4\1_home$\homedirs\rlee \\10.8.8.4\1_home$\homedirs\j4chan \\10.8.8.4\1_home$\homedirs\jkikuchi \\10.8.8.4\1_home$\homedirs\cbennet1"
		Case "h2"
			svol = "\2"
			dvol = "\002"
			w = "n"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.4\2_home$\homedirs\mdeleo \\10.8.8.4\2_home$\homedirs\mmcgowan \\10.8.8.4\2_home$\homedirs\mvilla \\10.8.8.4\2_home$\homedirs\numethgr"
		Case "h3"
			svol = "\3"
			dvol = "\003"
			w = "n"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.4\3_home$\homedirs\lcannady \\10.8.8.4\3_home$\homedirs\paguilar \\10.8.8.4\3_home$\homedirs\idesousa \\10.8.8.4\3_home$\homedirs\schang \\10.8.8.4\3_home$\homedirs\wgriffin \\10.8.8.4\3_home$\homedirs\dsavanna"
		Case "h4"
			svol = "\6"
			dvol = "\004"
			w = "y"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.10\004_home$\homedirs\rlee \\10.8.8.10\004_home$\homedirs\j4chan \\10.8.8.10\004_home$\homedirs\lcannady \\10.8.8.10\004_home$\homedirs\paguilar \\10.8.8.10\004_home$\homedirs\jkikuchi \\10.8.8.10\004_home$\homedirs\cbennet1"
			exclds = Array("a","b","c","d","e","f","g","h","n","o","p","q","r")
		Case "h5"
			svol = "\6"
			dvol = "\005"
			w = "y"
			paths = Array(path)
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD \\10.8.8.10\005_home$\homedirs\mdeleo \\10.8.8.10\005_home$\homedirs\mmcgowan \\10.8.8.10\005_home$\homedirs\mvilla \\10.8.8.10\005_home$\homedirs\numethgr \\10.8.8.10\005_home$\homedirs\idesousa \\10.8.8.10\005_home$\homedirs\schang \\10.8.8.10\005_home$\homedirs\wgriffin \\10.8.8.10\005_home$\homedirs\dsavanna"
			exclds = Array("i","j","k","l","m","s","t","u","v","w","x","y","z")
		Case "h6"
			svol = "\1"
			dvol = "\004"
			w = "n"
			paths = array(path & "rlee\",path & "j4chan\",path & "jkikuchi\",path & "cbennet1\")
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD"
		Case "h7"
			svol = "\3"
			dvol = "\004"
			w = "n"
			paths = array(path & "lcannady\",path & "paguilar\")
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD"
		Case "h8"
			svol = "\2"
			dvol = "\005"
			w = "n"
			paths = array(path & "mdeleo\",path & "mmcgowan\",path & "mvilla\",path & "numethgr\")
			RoboArg = " /copy:DAT /NFL /NDL /MIR /IPG:0 /r:0 /w:0 /NP /XD"
		Case "h9"
			svol = "\3"
			dvol = "\005"
			w = "n"
			paths = array(path & "idesousa\",path & "schang\",path & "wgriffin\",path & "dsavanna\")
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
	End If
' WScript.Echo "RoboArg: " & RoboArg
 'WScript.Quit

	'Setup RoboCopy to copy source to destination - capture output
	Set objShell = WScript.CreateObject("WScript.Shell")
'WScript.Echo "RoboArg = " & RoboArg	
	For Each dpath In paths
		SrcUNC = src & svol & dpath

		DestUNC = dest & dvol & dpath

'		DestUNC = dest & dvol & dpath
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

Sub chew
	For Each excld In exclds
		If left(sfnames,1) = excld Then
			RoboArg = RoboArg & excld
		End If
	next
End sub







	'Close the Log file
         objFileLog.Close
         WScript.Quit
