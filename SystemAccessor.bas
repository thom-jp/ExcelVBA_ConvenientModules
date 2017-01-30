Attribute VB_Name = "SystemAccessor"
#Const REF = False
'If REF then require references below _
- Microsoft Scripting Runtime _
- Windows Script Host Object Model _
- Microsoft ActiveX Data Objects X.X Library

'If Not REF Then Requir Nothing

#If REF Then
Private fso_ As Scripting.FileSystemObject
Private shell_ As IWshRuntimeLibrary.WshShell
#Else
Private fso_ As Object
Private shell_ As Object
#End If

#If REF Then
Property Get SharedFSO() As Scripting.FileSystemObject
#Else
Property Get SharedFSO() As Object
#End If
    If fso_ Is Nothing Then Set fso_ = CreateObject("Scripting.FileSystemObject")
    Set SharedFSO = fso_
End Property

#If REF Then
Property Get SharedWshShell() As IWshRuntimeLibrary.WshShell
#Else
Property Get SharedWshShell() As Object
#End If
    If shell_ Is Nothing Then Set shell_ = CreateObject("WScript.Shell")
    Set SharedWshShell = shell_
End Property

Function GetTempFilePath(Optional create_file As Boolean = False) As String
    Dim ret As String
    ret = Environ$("temp") & "\" & SharedFSO.GetTempName
    If create_file Then
        Call SharedFSO.CreateTextFile(ret)
    End If
    GetTempFilePath = ret
End Function

#If REF Then
Function GetCommandResultAsTextStream(command_string, Optional temp_path) As Scripting.TextStream
#Else
Function GetCommandResultAsTextStream(command_string, Optional temp_path) As Object
#End If
    Dim tempPath As String
    If IsMissing(temp_path) Then
        tempPath = GetTempFilePath
    Else
        tempPath = temp_path
    End If
#If Not REF Then
    Const WshHide = 0
    Const ForReading = 1
#End If
    Call SharedWshShell.Run("cmd.exe /c " & command_string & " > " & tempPath, WshHide, True)
    Set GetCommandResultAsTextStream = SharedFSO.OpenTextFile(tempPath, ForReading)
End Function

Function GetCommandResult(command_string) As String
    Dim ret As String
#If REF Then
    Dim ts As Scripting.TextStream
#Else
    Dim ts As Object
#End If
    Dim tempPath As String: tempPath = GetTempFilePath
    Set ts = GetCommandResultAsTextStream(command_string, tempPath)
    If ts.AtEndOfStream Then
        ret = ""
    Else
        ret = ts.ReadAll
    End If
    ts.Close
    Call SharedFSO.DeleteFile(tempPath, True)
    GetCommandResult = ret
End Function

Function GetCommandResultAsArray(command_string) As String()
    Dim ret() As String
    ret = Split(GetCommandResult(command_string), vbNewLine)
    GetCommandResultAsArray = ret
End Function

#If REF Then
Function GetPSCommandResultAsTextStream(command_string, Optional temp_path) As Scripting.TextStream
#Else
Function GetPSCommandResultAsTextStream(command_string, Optional temp_path) As Object
#End If
    Dim tempPath As String
    If IsMissing(temp_path) Then
        tempPath = GetTempFilePath
    Else
        tempPath = temp_path
    End If
#If Not REF Then
    Const WshHide = 0
    Const ForReading = 1
#End If
    Call SharedWshShell.Run("powershell -ExecutionPolicy RemoteSigned -Command Invoke-Expression """ & command_string & " | Out-File -filePath " & tempPath & " -encoding Default""", WshHide, True)
    Set GetPSCommandResultAsTextStream = SharedFSO.OpenTextFile(tempPath, ForReading)
End Function

Function GetPSCommandResult(command_string) As String
    Dim ret As String
#If REF Then
    Dim ts As Scripting.TextStream
#Else
    Dim ts As Object
#End If
    Dim tempPath As String: tempPath = GetTempFilePath
    Set ts = GetPSCommandResultAsTextStream(command_string, tempPath)
    If ts.AtEndOfStream Then
        ret = ""
    Else
        ret = ts.ReadAll
    End If
    ts.Close
    Call SharedFSO.DeleteFile(tempPath, True)
    GetPSCommandResult = ret
End Function

Function GetPSCommandResultAsArray(command_string) As String()
    Dim ret() As String
    ret = Split(GetPSCommandResult(command_string), vbNewLine)
    GetPSCommandResultAsArray = ret
End Function
