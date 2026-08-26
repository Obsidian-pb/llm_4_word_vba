Attribute VB_Name = "m_FileIO"
' Module: m_FileIO - File I/O operations (read/write UTF-8, path resolution, folder creation)
Option Explicit

' ============================================================
'  File operations with UTF-8 encoding
' ============================================================

' Reads entire text file in UTF-8 encoding
Public Function ReadFileUTF8(ByVal filePath As String) As String
    Dim stream As Object
    Dim result As String
    Dim errNum As Long
    Dim errDesc As String
    
    On Error GoTo CleanFail
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                       ' adTypeText = 2
    stream.CharSet = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    result = stream.ReadText(-1)          ' adReadAll = -1
    
    GoTo Cleanup
    
CleanFail:
    errNum = Err.Number
    errDesc = Err.Description
    
Cleanup:
    ' IO-03: unified resource cleanup - the stream is closed and released
    ' even if an error occurred during the operation above.
    On Error Resume Next
    If Not stream Is Nothing Then
        If stream.State <> 0 Then stream.Close
        Set stream = Nothing
    End If
    On Error GoTo 0
    
    If errNum <> 0 Then
        Err.Raise errNum, "ReadFileUTF8", errDesc
    End If
    
    ReadFileUTF8 = result
End Function

' Writes text to a file in UTF-8 encoding (without BOM).
' Overwrites existing file, creates missing folders.
Public Sub WriteFileUTF8(ByVal filePath As String, ByVal content As String)
    Dim fso As Object
    Dim stream As Object
    Dim folderPath As String
    Dim errNum As Long
    Dim errDesc As String
    
    On Error GoTo CleanFail
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    folderPath = fso.GetParentFolderName(filePath)
    If folderPath <> "" Then
        EnsureFolderExists folderPath
    End If
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                       ' adTypeText = 2
    stream.CharSet = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2         ' adSaveCreateOverWrite = 2
    
    GoTo Cleanup
    
CleanFail:
    errNum = Err.Number
    errDesc = Err.Description
    
Cleanup:
    ' IO-03: unified resource cleanup - the stream is closed and released
    ' even if an error occurred during the operation above.
    On Error Resume Next
    If Not stream Is Nothing Then
        If stream.State <> 0 Then stream.Close
        Set stream = Nothing
    End If
    Set fso = Nothing
    On Error GoTo 0
    
    If errNum <> 0 Then
        Err.Raise errNum, "WriteFileUTF8", errDesc
    End If
End Sub

' Resolves a path to absolute form. A relative path
' is resolved relative to the active Visio document folder.
Public Function ResolveFilePath(ByVal filePath As String) As String
    filePath = Trim$(filePath)
    
    ' Remove enclosing quotes
    If Left$(filePath, 1) = """" Then filePath = Mid$(filePath, 2)
    If Right$(filePath, 1) = """" Then filePath = Left$(filePath, Len(filePath) - 1)
    
    ' Check if the path is absolute (drive: or \\server\share)
    Dim isAbsolute As Boolean
    isAbsolute = (InStr(1, filePath, ":", vbTextCompare) > 0) Or (Left$(filePath, 2) = "\\")
    
    If Not isAbsolute Then
        Dim basePath As String
        basePath = ""
        On Error Resume Next
        basePath = ActiveDocument.path
        On Error GoTo 0
        If basePath = "" Then basePath = CurDir$
        
        If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
        filePath = basePath & filePath
    End If
    
    ResolveFilePath = filePath
End Function

' Checks whether a file exists using FileSystemObject.FileExists.
' More reliable than Dir$ for special paths and files with non-standard attributes.
Public Function FileExists(ByVal filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    Set fso = Nothing
End Function

' Recursively creates a folder with all nested subfolders.
Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(folderPath) Then
        Set fso = Nothing
        Exit Sub
    End If
    
    Dim parent As String
    parent = fso.GetParentFolderName(folderPath)
    If parent <> "" And Not fso.FolderExists(parent) Then
        EnsureFolderExists parent
    End If
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    Set fso = Nothing
End Sub

