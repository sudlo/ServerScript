strComputer = "."
Set wmiServices = GetObject _
    ("winmgmts:{impersonationLevel=Impersonate}!//" & strComputer)

Set wmiDiskDrives = wmiServices.ExecQuery _
    ("SELECT Caption, DeviceID FROM Win32_DiskDrive")
 
For Each wmiDiskDrive In wmiDiskDrives
    WScript.Echo wmiDiskDrive.Caption & " (" & wmiDiskDrive.DeviceID & ")"
    strEscapedDeviceID = Replace _
        (wmiDiskDrive.DeviceID, "\", "\\", 1, -1, vbTextCompare)
    Set wmiDiskPartitions = wmiServices.ExecQuery _
        ("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & _
    strEscapedDeviceID & """} WHERE AssocClass = " & _
        "Win32_DiskDriveToDiskPartition")
 
    For Each wmiDiskPartition In wmiDiskPartitions
        WScript.Echo vbTab & wmiDiskPartition.DeviceID
        Set wmiLogicalDisks = wmiServices.ExecQuery _
            ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & _
                wmiDiskPartition.DeviceID & """} WHERE AssocClass = " & _
                    "Win32_LogicalDiskToPartition")
 
        For Each wmiLogicalDisk In wmiLogicalDisks
            WScript.Echo vbTab & vbTab & wmiLogicalDisk.DeviceID
        Next
    Next
Next
