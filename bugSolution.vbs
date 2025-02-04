Option Explicit

' Early binding for better error handling
Dim objFSO As Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Explicit type checking
Dim strFilePath As String
strFilePath = "C:\test.txt"

If objFSO.FileExists(strFilePath) Then
  ' Correctly handling file size as a number
  Dim fileSize As Long
  fileSize = objFSO.GetFile(strFilePath).Size
  MsgBox "File size: " & fileSize & " bytes"
Else
  MsgBox "File not found!"
End If

' Clean up object
Set objFSO = Nothing