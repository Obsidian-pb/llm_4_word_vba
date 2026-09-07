VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_InsertText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsTool_InsertText - Tool: insert text to selected range
Option Explicit
Implements clsToolBase

' Tool name (what LLM sees in tools[])
Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "insert_text"
End Property

' Description for LLM
Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "«амен€ет текст Word в выделеннной области переданным текстом. " & _
        "¬ызывай его, когда пользователь просит вставить в документ некоторый текст. " & _
        "≈сли агрументы 'start_pos' и 'end_pos' не переданы, текст будет вставлен в место в котором сейчас находитс€ курсор. " & _
        "»нструмент может быть использован и дл€ удалени€ текста - дл€ этого следует вернуть пустую строку, в этом случае " & _
        "все содерржимое области в диапазоне от 'start_pos' ло 'end_pos' будет заменено ей, то есть стерто."
End Property

' JSON Schema of parameters
Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """text_for_insert"":{""type"":""string"",""description"":""Text which will be inserted into Word document.""}," & _
            """start_pos"":{""type"":""integer"",""description"":""Range start pos for insert.""}," & _
            """end_pos"":{""type"":""integer"",""description"":""Range end pos for insert.""}" & _
        "}," & _
        """required"":[""text_for_insert""]" & _
    "}"
End Property

' Execute the tool
Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    
    On Error GoTo ErrHandler
    
    Dim text_for_insert As String
    Dim start_pos_s As String
    Dim end_pos_s As String
    text_for_insert = ExtractJsonString(argsJson, "text_for_insert")
    start_pos_s = ExtractJsonString(argsJson, "start_pos")
    end_pos_s = ExtractJsonString(argsJson, "end_pos")
        
    ' 1. ѕолучаем выделенный текст
    Dim SelText As String
'    If Selection.Type = wdSelectionNormal And Selection.Range.Characters.Count > 0 Then
'        SelText = Selection.text
'    Else
'        SelText = ""
'    End If
    
    ' 2. Insertion
    Dim rng As Word.Range
    If Not start_pos_s = "" And Not end_pos_s = "" Then
        Set rng = ActiveDocument.Range(CLng(start_pos_s), CLng(end_pos_s))
    Else
        Set rng = Selection.Range
    End If
    SelText = rng.text
    rng.text = text_for_insert
    
    ' 3. Result
    Dim msg As String
    If SelText = "" Then
        msg = "“екст '" & text_for_insert & "' вставлен в область (" & start_pos_s & ":" & end_pos_s & ")."
    Else
        msg = "“екст '" & SelText & "' заменен текстом '" & text_for_insert & "' в области (" & start_pos_s & ":" & end_pos_s & ")."
    End If
    clsToolBase_Execute = "{""status"":""ok"",""message"":" & JsonString(msg) & "}"
    
    
    Exit Function
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString(Err.Description) & "}"
End Function

