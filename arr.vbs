'This script is provided under the Creative Commons license located
'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
'be used for commercial purposes with out the expressed written consent
'of NateRice.com

vServer = "SERVERNAME"
vCommunityString = "YOURCONNECTIONSTRING"
 
'populate controllers
aControllers = Split(SNMPWALK(vServer, vCommunityString, _
".1.3.6.1.4.1.232.3.2.2.1.1.1", "0"), "|")
  
vNumberOfControllers = UBound(aControllers)

'We're creating an array that will store data for up to
'10 controllers and 100 drives on each controller.
'You can change this, if you have more.
Dim aDriveSizes(10, 100)

'This script will error every time for a couple reasons.
'First is how we parse information for display. Second is
'occasionally data will not be returned from the device
'you query.
On Error Resume Next
  
vControllerNumber = -1
Do Until vControllerNumber = UBound(aControllers) 'populate controller info
  vControllerNumber = vControllerNumber + 1
    
  'pull drive info for controller
  aSingleDriveSizes = Split(SNMPWALK(vServer, vCommunityString, _
  ".1.3.6.1.4.1.232.3.2.5.1.1.45." & aControllers(vControllerNumber), "0"), _
  "|") 
    
  vDriveLoop = 0
  For Each vSingleDriveSize In aSingleDriveSizes
    'populate drive sizes for current controller
    aDriveSizes(vControllerNumber, vDriveLoop) = vSingleDriveSize 
    vDriveLoop = vDriveLoop + 1
  Next
Loop

'/////DISPLAY INFO\\\\\\

vControllerNumber = 0
For Each vController In aControllers

  vDisplayString = vDisplayString & "Controler Information " & _
  vController & vbLF

  vDriveID = -1
  Do Until vDriveID = 100
    vDrive = aDriveSizes(vControllerNumber, vDriveID)
    vDriveID = vDriveID + 1
    If Len(vDrive) > 0 Then
      vDisplayString = vDisplayString & "  -- Drive " & vDriveID & " (" & _
      vDrive & "MB)" & vbLF
    End If
  Loop
   
  vControllerNumber = vControllerNumber + 1
Next
   
WScript.Echo vDisplayString