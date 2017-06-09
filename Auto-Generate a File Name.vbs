Set objFSO = CreateObject("Scripting.FileSystemObject")

For i = 1 to 10
    strTempFile = objFSO.GetTempName
    Wscript.Echo strTempFile
Next
