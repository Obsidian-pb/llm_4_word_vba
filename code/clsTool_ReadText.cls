VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_ReadText"
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
    clsToolBase_Description = "Читает текст расположенный в указанной области в докуменнте Word. " & _
                              "Если область (start_pos:end_pos) не задана, читается текст выделенный в данный момент."
End Property

' JSON Schema of parameters
Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """start_pos"":{""type"":""integer"",""description"":""Range start pos for text read.""}," & _
            """end_pos"":{""type"":""integer"",""description"":""Range end pos for text read.""}" & _
        "}," & _
        """required"":[]" & _
    "}"
End Property

' Execute the tool
Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    
    On Error GoTo ErrHandler
    
    ' 0. Получение аргументов
    Dim start_pos_s As String
    Dim end_pos_s As String
    start_pos_s = ExtractJsonString(argsJson, "start_pos")
    end_pos_s = ExtractJsonString(argsJson, "end_pos")
    
    
    ' 1. Получаем выделенный текст
    Dim SelText As String
    Dim SelStart As Long
    Dim SelEnd As Long
    
    If Not start_pos_s = "" And Not end_pos_s = "" Then
        SelStart = CLng(start_pos_s)
        SelEnd = CLng(end_pos_s)
        SelText = ActiveDocument.Range(SelStart, SelEnd).text
    Else
        If Selection.Type = wdSelectionNormal And Selection.Range.Characters.Count > 0 Then
            SelText = Selection.text
            SelStart = Selection.Start
            SelEnd = Selection.End
        Else
            SelText = ""
        End If
    End If
    

    SelText = Trim$(SelText)
    
    ' 2. Result
    Dim msg As String
    If SelText = "" Then
        msg = "В настоящий момент никакой текст в документе не выделен."
    Else
        msg = "В настоящий момент в документе, диапазоне (" & SelStart & ":" & SelEnd & ") выделен следующий текст: '" & SelText & "'."
    End If
    clsToolBase_Execute = "{""status"":""ok"",""message"":" & JsonString(msg) & "}"
    
    Exit Function
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString(Err.Description) & "}"
End Function



