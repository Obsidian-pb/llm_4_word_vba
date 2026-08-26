VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_InsertComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_InsertComment - Tool: insert comment to selected range
Option Explicit
Implements clsToolBase

' Tool name (what LLM sees in tools[])
Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "insert_comment"
End Property

' Description for LLM
Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Вставляет комментарий к выделенному фрагменту текста."
End Property

' JSON Schema of parameters
Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """comment"":{""type"":""string"",""description"":""Комментарий который будет добавлен к выделенному тексту.""}" & _
        "}," & _
        """required"":[""comment""]" & _
    "}"
End Property

' Execute the tool
Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    
    On Error GoTo ErrHandler
    
    Dim comment As String
    comment = ExtractJsonString(argsJson, "comment")
    
    If Trim$(comment) = "" Then
        clsToolBase_Execute = "{""status"":""error"",""message"":""Комментарий не может быть пустым.""}"
    End If
    
    ' 1. Получаем выделенный текст
    Dim SelText As String
    If Selection.Type = wdSelectionNormal And Selection.Range.Characters.Count > 0 Then
        SelText = Selection.text
    Else
        SelText = ""
    End If
       
    ' 2. Insertion
    Selection.Comments.Add Range:=Selection.Range
    Selection.TypeText text:=comment
    
    Dim rng As Word.Range
    Set rng = Selection.Range
    
    ' 3. Result
    Dim msg As String
    If SelText = "" Then
        msg = "Комментарий '" & comment & "' вставлен в место расположения курсора."
    Else
        msg = "Текст '" & SelText & "' снабжен комментарием '" & comment & "'."
    End If
    clsToolBase_Execute = "{""status"":""ok"",""message"":" & JsonString(msg) & "}"
    
    
    Exit Function
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString(Err.Description) & "}"
End Function



