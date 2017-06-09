' ====================================================================================
' TITLE:        	FolderSizeTracker.vbs
' AUTHOR:       	Greg Shultz
' DATE : 			07/18/2003
' URL:				http://www.thewinwiz.com
'
' REVISED BY:		Christian Sawyer
' EMAIL:			csawyer@implanciel.com
' REVISED DATE:		8/26/2004
' COMPANY: 			Implanciel Inc.
'
' PURPOSE:  		The Folder Size Tracker Tool offers a more efficient alternative 
'					to the manual approach. To find out which folders on a hard disk
'					are consuming the most space, you just point Folder Size Tracker 
'					to the path you want to investigate and let it go to work. 
'					The results are then compiled in an Excel worksheet and displayed 
'					in a rotating 3-D pie chart, making it easy to spot which folders 
'					are hogging the most disk space.
'
' HOW TO USE:		Run it from Windows Explorer or Dos Prompt.
'					
' NOTES:			The Folder Size Tracker Tool works in all versions of Windows and 
'					with Excel 2000/2002/2003.
'
' Added by Revisor:	All comments and all headers.
'
'					GetFolder and GoSubFolders fonctionnality to allow extract
'					all subfolders in each folder. In original release, it only
'					list first level from the choosen path.
'
'					BrowseForFolderDialogBox to replace InputBox.
'
'					IsNoData to replace all verification for empty string.
'
'					GetScriptPath to replace objShell.CurrentDirectory.
'					.CurrentDirectory seem to point at last open directory.
'
'					
' ====================================================================================
'Main Sub
Option Explicit
Dim objShell
Dim objFSO
Dim objXL
Dim objDir
Dim intRow
Dim strPath
Dim strFile
Dim strCurDir
 
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objXL = WScript.CreateObject("Excel.Application")
' Modified by Christian Sawyer. Replace objshell.CurrentDirectory by
' GetScriptPath.
strCurDir = GetScriptPath
strFile = strCurDir + "FolderSizeTracker.ini"

GetPath				'Ask user to choose which folder to analyse.
SplashScreen		'Show to user what's happen and who build this script.
CreateSpreadSheet	'Create Excel file.
'*****************************************************************************
' Added by Christian Sawyer to recursively find all folders and Subfolders.
Set objDir = GetFolder(strPath)
BuildSpreadSheet objDir
GoSubFolders objDir
' End addition.
'*****************************************************************************
SortData			'Sort data in Excel file based on folder size.
CreateChart			'Create pie chart based on sheet Folder Size Data.
RotateChart			'Rotate chart to place biggest folder at bottom.
'If strFile doesn't exist, Call instructions to show to user.
'Modified by Christian Sawyer. Replace If Then Else Endif by one line IF.
If Not objFSO.FileExists(strFile) Then Call Instructions
  
'End Main Sub
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	SplashScreen
'
' PURPOSE:  		Show message box to user what is happen and brief presentation
'					of original author, Greg Shultz and the Revisor.
'
' HOW TO USE:		SplashScreen
'
' ====================================================================================
Sub SplashScreen
	Dim intRetVal
	Dim strSplash
  	
	strSplash = "Please wait while Folder Size Tracker analyzes the folder contents" & vbCrLf & _
  				"and adds data to an Excel spreadsheet." & vbCrLf & vbCrLf & _
  				"Created for TechRepublic      " & vbCrLf & _
  				"by Greg Shultz" & vbCrLf & _
  				"www.TheWinWiz.com" & vbCrLf & _
      			"Revised By Christian Sawyer"		'Added by Christian Sawyer.
  	intRetVal = objShell.Popup(strSplash, 2, "Folder Size Tracker", 64)

End Sub
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	GetPath
'
' PURPOSE:  		Ask user to choose a folder to analyze.
'					Here, strPath is global.
'
' HOW TO USE:		GetPath
'
' ====================================================================================
Sub GetPath
	Dim strHeader
	strHeader = "Enter the drive or path in which you want to track folder size." & vbCrLf & vbCrLf &_
  			    "For example:  C:\   or   C:\Documents and Settings"
	'Modified by Christian Sawyer. Replace InputBox by BrowseForFolderDialogBox.
	strPath = BrowseForFolderDialogBox(strHeader)
	'Modified by Christian Sawyer. Introduce IsNoData.
	If IsNoData(strPath) Then WScript.Quit
End Sub
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	CreateSpreadSheet
'
' PURPOSE:  		Create the spreadsheet with two columns (Folder, Total in MB)
'					and set the header. Name also the active sheet (Folder Size Data).
'
' HOW TO USE:		CreateSpreadSheet
'
' ====================================================================================
Sub CreateSpreadSheet
   	objXL.Visible = True
   	objXL.WorkBooks.Add
   	objXL.Worksheets(2).visible = False
   	objXL.Worksheets(3).visible = False
   	objXL.ActiveSheet.Name = "Folder Size Data"
   	objXL.Range("A1").Select
   	objXL.Columns(1).ColumnWidth = 30
   	objXL.Columns(2).ColumnWidth = 10
   	objXL.Columns(2).NumberFormat = "#,##0.0"
   	objXL.ActiveSheet.Cells(1,1).Value = "Folders in: " & strPath
   	objXL.Range("A1:A1").Select
   	objXL.Selection.Font.Size = 14
   	objXL.Range("A2:B2").Select
   	objXL.Selection.Font.Size = 12
   	objXL.ActiveSheet.Cells(3,1).Value = "Folder"
   	objXL.ActiveSheet.Cells(3,2).Value = "Total (MB)"
   	objXL.ActiveSheet.Cells(3,2).AddComment "Folder Size Tracker doesn't display folders containing 1MB or less."
   	objXL.Range("A1:B3").Select
    objXL.Selection.Font.Bold = True
   	objXL.Range("A5:B5").Select
   	intRow = 5
End Sub
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	BuildSpreadSheet
'
' PURPOSE:  		Fill the spreadsheet with folder name and folder size
'					for each folder passed in parameter.
'
' HOW TO USE:		BuildSpreadSheet objFolderName
'
' ====================================================================================
Sub BuildSpreadSheet(strSubFolder)
	Dim objFolder
	Dim intFolderSize
	'Modified by Christian Sawyer to be recalled by GoSubFolders.
	For Each objFolder In strSubFolder.SubFolders
		'Verify if current folder is a system folder or not.
		If objFolder.Name <> "System Volume Information" Then
			intFolderSize = objFolder.Size
			intFolderSize = (FormatNumber(intFolderSize, 0, , , 0)/1024)/1024 
			If intFolderSize > 1 Then
				objXL.ActiveSheet.Cells(intRow,1).Value = objFolder.Name
				objXL.ActiveSheet.Cells(intRow,2).Value = intFolderSize
				intRow = intRow + 1
			End If
		End If
	Next
End Sub
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	GetFolder
'
' PURPOSE:  		Verify if passed directory in parameter exist or not.
'					If exist, return the folder name. Otherwise, display
'					a message and quit. Need one parameter, objFolder to check.
'
' HOW TO USE:		GetFolder objFolderName
'
' NOTE:				This function has been added by Christian Sawyer.
' ====================================================================================
Function GetFolder(strFolder)
On Error Resume Next
	Set GetFolder = objFSO.GetFolder(strFolder)
	If Err.Number <> 0 Then
		Wscript.Echo "Error connecting to folder: " & strFolder & vbCrLf & _
					 "[" & Err.Number & "] " & Err.Description
		Wscript.Quit Err.Number
	End If
End Function
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	GoSubFolders
'
' PURPOSE:		  	Loop in all subdirectories under objDIR passed in parameters.
'					Call BuildSpreadSheet with new subfolder to extract
'					folder size from it.
'
' HOW TO USE:		GoSubFolders objFolderName
'
' NOTE:				This function has been added by Christian Sawyer.
' ====================================================================================
Sub GoSubFolders (objDIR)
	Dim objFolder
	' Verify if objDIR is not a system directory.
	If objDIR <> "\System Volume Information" Then 
		For Each objFolder in objDIR.SubFolders
			BuildSpreadSheet objFolder	'Recall BuildSpreadSheet with new subfolder.
			GoSubFolders objFolder		'Recall itself recursively.
		Next
	End If   
End Sub
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	SortData
'
' PURPOSE:		  	Verify first if intRow is <=5. If yes, means that the choosen
'					folder is a system folder or you choose an empty folder.
'					If greater than 5, will sort data based on size in column 3.
'
' HOW TO USE:		SortData
'
' ====================================================================================
Sub SortData
	Dim strMsg
	'Value of 5 has been set in CreateSpreadSheet
	If intRow = 5 Then
		strMsg = "Folder Size Tracker was unable to find any folders in the path " & vbcrlf & vbcrlf &_ 
				 strPath & vbcrlf & vbcrlf &_
				 "over the size of 1MB!" & vbcrlf & vbcrlf &_
				 "Folder Size Tracker will now shut down."
		MsgBox strMsg, 16, "Folder Size Tracker"
		objXL.ActiveWorkBook.Saved = True
		objXL.WorkBooks.Close
		objXL.Application.Quit
		Wscript.Quit
	Else
		objXL.ActiveCell.CurrentRegion.Select
		objXL.Selection.Sort objXL.Worksheets(1).Range("B3"), 2, , , , , , 0, 1, False, 1
	End If
End Sub
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	CreateChart
'
' PURPOSE:		  	Create a chart in a new sheet with legend and center the chart.
'
' HOW TO USE:		CreateChart
'
' ====================================================================================
Sub CreateChart
	objXL.Charts.Add
	objXL.ActiveChart.ChartType = 70
	objXL.ActiveChart.Name = "Pie Chart"		
	objXL.ActiveChart.PlotArea.Select
	objXL.Selection.Left = 1
	objXL.Selection.Top = 1
	objXL.Selection.Height = 100
	objXL.Selection.Width = 675	

	objXL.ActiveChart.HasLegend = True	
	objXL.ActiveChart.Legend.Position = -4107 
End Sub
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	RotateChart
'
' PURPOSE:		  	Make the pie chart rotate in two ways. Will show biggest
'					folder at bottom of chart.
'
' HOW TO USE:		RotateChart
'
' ====================================================================================
Sub RotateChart
	Dim intRotate
	
	For intRotate = 10 To 360 Step 10
		WScript.Sleep(100)
    	objXL.ActiveChart.Rotation = intRotate
	Next

	For intRotate = 350 To 0 Step -10
		WScript.Sleep(100)
    	objXL.ActiveChart.Rotation = intRotate
	Next
	
	objXL.ActiveChart.ChartArea.Select
	objXL.ActiveChart.Deselect
End Sub
'-------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	Instructions
'
' PURPOSE:		  	Display instructions on how to use the chart, how to modify it,
'					after the spreadsheet is completed.
'
' HOW TO USE:		Instructions
'
' ====================================================================================
Sub Instructions
	Dim objFile
	Dim intRetVal
	Dim strMsg
	
	strMsg = "To better identify smaller slices, you can manually rotate the pie chart:" & vbcrlf & vbcrlf & _
	"1) Pull down the Chart menu" & vbcrlf &_
	"2) Select the 3-D View command" & vbcrlf &_
	"3) Click either of the Rotate buttons" & vbcrlf &_ 
	"4) Click the Apply button"

	intRetVal = MsgBox(strMsg, vbOkOnly, "Folder Size Tracker")

	strMsg = "To better identify the folder represented by a particular slice," & vbcrlf &_
	"simply hover your mouse pointer over that slice." & vbcrlf & vbcrlf & _
	"The Value figure is the size of the folder in MB."
	
	intRetVal = MsgBox(strMsg, vbOkOnly, "Folder Size Tracker")
	
	strMsg = "Keep in mind that Folder Size Tracker doesn't display folders" & vbcrlf &_ 
	"containing 1MB or less."
	
	intRetVal = MsgBox(strMsg, vbOkOnly, "Folder Size Tracker")
	
	strMsg = "Do you want to see these instructions again?"
	
	intRetVal = MsgBox(strMsg, vbYesNo, "Folder Size Tracker")
	
	If intRetVal = vbNo Then
		strMsg = "If at a later date you want to see these instructions again, simply locate and delete the FolderSizeTracker.ini file"
		intRetVal = MsgBox(strMsg, vbOkOnly, "Folder Size Tracker")
		Set objFile = objFSO.CreateTextFile(strFile, True)
	End If
End Sub
'--------------------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	BrowseForFolderDialogBox
'
' PURPOSE:		  	Open Window dialog box to choose directory to process.
'					Return the choosen path to a variable.
'
' HOW TO USE:		strResult = BrowseForFolderDialogBox("Which folder?")
'
' NOTE:				This function has been added by Christian Sawyer.
' ====================================================================================
Function BrowseForFolderDialogBox(strTitle)
	Const WINDOW_HANDLE = 0
	Const NO_OPTIONS = &H0001
	Dim objShellApp
	Dim objFolder
	Dim objFldrItem
	Dim objPath
	
	If IsNoData(objShellApp) Then
		Set objShellApp = WScript.CreateObject("Shell.Application")
	End If
	Set objFolder = objShellApp.BrowseForFolder(WINDOW_HANDLE, strTitle , NO_OPTIONS)
	If IsNoData(objFolder) Then
		Wscript.Echo "You choose to cancel. This will stop this script."
		Wscript.Quit
	Else
		Set objFldrItem = objFolder.Self
			objPath = objFldrItem.Path
			BrowseForFolderDialogBox = objPath
		Set objShellApp	= Nothing
		Set objFolder	= Nothing
		Set objFldrItem	= Nothing
	End If
End Function
' ====================================================================================
' TITLE:        	IsNoData
'
' PURPOSE:		  	Verify if passed parameter variable contain something.
'
' RETURN:			True if contain something, otherwise, False.
'
' HOW TO USE:		intResult = IsNoData(varSource) or
'					If IsNoData(varSource) Then do something.
'
' NOTE:				This function has been added by Christian Sawyer.
' ====================================================================================
Function IsNoData(varVal2Check)
	'Verify if varVal2Check contain something.
	On Error Resume Next
    If IsNull(varVal2Check) Or IsEmpty(varVal2Check) Then
		IsNoData = True
    Else
        If IsDate(varVal2Check) Then
			IsNoData = False
        ElseIf varVal2Check = "" Then
			IsNoData = True
		ElseIf Not IsObject(varVal2Check) Then
			IsNoData = False
		Else
            IsNoData = False
        End If
    End If
End Function
'--------------------------------------------------------------------------------------------------
' ====================================================================================
' TITLE:        	GetScriptPath
'
' PURPOSE:  		Find the path where the current script were executed from.
'					Used to save any file in same path where the script reside.
'
' HOW TO USE:		strResult = GetScriptPath()
'
' NOTE:				This function has been added by Christian Sawyer.
' ====================================================================================
Function GetScriptPath() ' Find path from where this script is executed.
	GetScriptPath = MID(Wscript.ScriptFullName, 1, InstrRev(Wscript.ScriptFullName,"\"))
End Function

