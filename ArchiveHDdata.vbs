'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	Written/Maintined by EPS/Intel - C. Ross - 12/03

'	Script will return either HomeDir UNC, or "Creation Failure"

'	Script requires copy of RMTSHARE.EXE (NT4 Resource Kit) in D:\Gateway\adScripts\Util\
'	Script requires copy of ROBOCOPY.EXE (2003 Resource Kit) in D:\Gateway\adScripts\Util\

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	Usage syntax:
  WScript.Sleep 5000 ' to help avoid conflict writing to log file
  Set objArgs = WScript.Arguments
  if objArgs.Count < 6 Then
     WScript.Echo "Usage:  "
     WScript.Echo "    ArchiveHDdata.vbs UserID SourceUNC TargetUNC SourcePath LogName LogPath"
   End if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     UserID = ucase(objArgs(0))
     SrcUNC = ucase(objArgs(1))
     DestUNC = ucase(objArgs(2))
     SrcPath = ucase(objArgs(3))
     LogName = objArgs(4)
     LogPath = objArgs(5)
     Robo = "d:\gateway\adScripts\util\robocopy.exe "
     RoboArgs = Array(" /create /copy:DAT /e /r:2 /w:0 /NP /NFL /NDL /XD ~snapshot"," /copy:DAT /IS /A-:A /e /r:2 /NP /w:0 /MOVE /XD ~snapshot")
     exclDirs = Array("\Clxflags","\Internet","\Notedata","\Remedy","\Rwin","\Rwin611","\Win95","\Winnt")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	Connect to existing Log File:

   Const ForReading = 1
   Const ForAppending = 8
   Const DeleteReadOnly = True
   Const DeleteHidden = True

   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objFileLog = objFSO.OpenTextFile(LogPath & LogName, ForAppending)

    objFileLog.WriteLine "<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>"
    objFileLog.WriteLine now 	'adds timestamp to start of Log

  if objArgs.Count < 6 Then
     objFileLog.WriteLine "Usage:  "
     objFileLog.WriteLine "    MoveData UserID UserID SourcePath TargetPath PhysicalPath LogPath LogName"
     objFileLog.close
     WScript.Quit(1)
  End If

    objFileLog.WriteBlankLines (1)
    objFileLog.WriteLine "Starting Archive HD Data script:"
    objFileLog.WriteLine "    - From Trigger File: UserID = " & ucase(UserID)
    objFileLog.WriteLine "    - From Trigger File: Source UNC = " & ucase(SrcUNC)
    objFileLog.WriteLine "    - From Trigger File: Destination UNC = " & ucase(DestUNC)
    objFileLog.WriteLine "    - From Trigger File: Source Path = " & ucase(SrcPath)
    objFileLog.WriteLine "    - From Trigger File: Log Name = " & LogName
    objFileLog.WriteLine "    - From Trigger File: Log Path = " & LogPath
    objFileLog.WriteBlankLines (1)


'   Delete "excluded" folders in source
    objFileLog.WriteLine "------------------------------------------------------------------------------"
    objFileLog.WriteLine "Check for and delete any ""excluded"" folders"
    For Each dir In ExclDirs

         Set objFSO = CreateObject("Scripting.FileSystemObject")
     If objFSO.FolderExists(SrcPath & Ucase(dir)) Then
'         objFileLog.WriteBlankLines (1)
         objFileLog.WriteLine "    ! Deleting folder: " & SrcPath & ucase(dir)
         objFSO.DeleteFolder(SrcPath & ucase(dir)), true
         objFileLog.WriteLine "        * Error level reported: " & Err 
     Else
'         objFileLog.WriteBlankLines (1)
         objFileLog.WriteLine "    - Checked - folder does not exist: " & SrcPath & ucase(dir)
     End If

     Next
    objFileLog.WriteLine "    - Completed checking for excluded folders"
    objFileLog.WriteBlankLines (1)
    objFileLog.WriteLine "------------------------------------------------------------------------------"
    objFileLog.WriteLine "Starting RoboCopy process:"
    objFileLog.WriteLine "    - RoboCopy will make 2 passes."
    objFileLog.WriteLine "      1st pass creates dirs and 0 length file names,"
    objFileLog.WriteLine "      2nd pass copies actual data, and deletes source upon completion."
    objFileLog.WriteBlankLines (1)
    objFileLog.WriteLine "    - Depending on size, this may take a while..."
    objFileLog.WriteBlankLines (1)
    objFileLog.WriteLine "    - Check final log entry to confirm deletion of source folder(s)."
    objFileLog.WriteLine "      If source not automatically deleted, it must be determined why, then cleaned up manually."
WScript.Sleep 1000

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Setup RoboCopy to "MOVE" HD source to destination.

       For Each arg In RoboArgs
          Set objShell = WScript.CreateObject("WScript.Shell")
          CMDLine = SrcPath & " " & destUNC & "\ArchivedHomeDirs\" & ucase(UserID) & " " & arg
          Set objExecObject = objShell.Exec(Robo & CMDLine)
'   Wait for RoboCopy to complete, then write to log...
          Set objStdOut = objExecObject.StdOut
          strOutput = objStdOut.ReadAll
          objFileLog.Write strOutput
       Next

    objFileLog.WriteBlankLines (1)
    objFileLog.WriteLine "------------------------------------------------------------------------------"
    objFileLog.WriteLine "    Robo Copy process reports: Ended"
    objFileLog.WriteLine "------------------------------------------------------------------------------"
         objFileLog.WriteBlankLines (1)

'''''''''''''''''''''''''''''''''''''''
    objFileLog.WriteBlankLines (1)
'	'''''''''''''''''''''''''''''''''''''''''''''''''''
'	Verify Folder Deleted

    WScript.Sleep 100
    objFileLog.WriteLine "Confirm Source Folder deletion."
        LoopVariable = 1
        Do Until LoopVariable > 5
           Set objFSO = CreateObject("Scripting.FileSystemObject")
         If objFSO.FolderExists(SrcPath) = False Then
           objFileLog.WriteLine "    - Success - Source Folder Deleted: " & srcpath
           Wscript.StdOut.WriteLine "Success:"
           Exit Do
          Else
           objFileLog.WriteLine "    - Waiting 2.5 sec for deletion of folder... Loop Number = " & LoopVariable 
         End If
           LoopVariable = LoopVariable + 1
           wscript.sleep 2500 '2.5 Seconds
        Loop 

'	'''''''''''''''
'	Report Script Failed or end script
 
          Set objFSO = CreateObject("Scripting.FileSystemObject")
          If objFSO.FolderExists(SrcPath) = True Then
'           Set objFolder = objFSO.GetFolder(SrcPath & ucase(UserID))
           objFileLog.WriteBlankLines (1)
           objFileLog.WriteLine "*** - Fatel Error - Source Folder NOT deleted: " & SrcPath
           objFileLog.WriteLine "      It must be determined why (not all files copied?), then cleaned up manually."

           Wscript.StdOut.WriteLine "Error:"
          Else
           objFileLog.WriteLine "    - Echo'd to stdOut: ""Success:"""
           objFileLog.WriteBlankLines (1)
           objFileLog.WriteLine "Script completed successfully"
          End If

''''''''''''''''''''''''''''''''''''
'	Close the Log file

         objFileLog.WriteBlankLines (1)
         objFileLog.WriteLine now
         objFileLog.Close
         WScript.Quit
