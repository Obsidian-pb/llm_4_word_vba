VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_ReadSelectedText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_InsertText - Tool: insert text to selected range
Option Explicit
Implements clsToolBase

' Tool name (what LLM sees in tools[])
Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "read_selected_text"
End Property

' Description for LLM
Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Читает текст выделенный в данный момент в докуменнте Word."
End Property

' JSON Schema of parameters
Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{}," & _
        """required"":[]" & _
    "}"
End Property

' Execute the tool
Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    
    On Error GoTo ErrHandler
    
    
    ' 1. Получаем выделенный текст
    Dim SelText As String
    If Selection.Type = wdSelectionNormal And Selection.Range.Characters.Count > 0 Then
        SelText = Selection.text
    Else
        SelText = ""
    End If
    SelText = Trim$(SelText)
    
    ' 2. Result
    Dim msg As String
    If SelText = "" Then
        msg = "В настоящий момент никакой текст в документе не выделен."
    Else
        msg = "В настоящий момент в документе не выделен следующий текст: '" & SelText & "'."
    End If
    clsToolBase_Execute = "{""status"":""ok"",""message"":" & JsonString(msg) & "}"
    
    Exit Function
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString(Err.Description) & "}"
End Function



