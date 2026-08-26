VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_WriteTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_WriteTextFile - Tool: write text content to a file (UTF-8)
Option Explicit
Implements clsToolBase

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "write_text_file"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Writes text content to a file in UTF-8 encoding. " & _
        "If the file already exists, it will be completely overwritten. Missing folders are created automatically. " & _
        "Use this tool when you need to save, write or create a text file."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """file_path"":{""type"":""string"",""description"":""Path to the text file (absolute or relative)""}," & _
            """content"":{""type"":""string"",""description"":""Text content to write to the file""}" & _
        "}," & _
        """required"":[""file_path"",""content""]" & _
    "}"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler
    
    Dim file_path As String
    Dim content As String
    file_path = ExtractJsonString(argsJson, "file_path")
    content = ExtractJsonString(argsJson, "content")
    
    If Trim$(file_path) = "" Then
        clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString("File path is not specified") & "}"
        Exit Function
    End If
    
    file_path = ResolveFilePath(file_path)
    
    Call WriteFileUTF8(file_path, content)
    
    clsToolBase_Execute = "{""status"":""ok"",""message"":" & JsonString("File successfully written: " & file_path) & "}"
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString(Err.Description) & "}"
End Function


