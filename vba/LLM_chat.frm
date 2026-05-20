VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LLM_chat 
   Caption         =   "Чат с LLM"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10410
   OleObjectBlob   =   "LLM_chat.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LLM_chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'====  Публичные переменные ====
Public answer As String


Private Sub CB_Close_Click()
    Me.Hide
End Sub

Private Sub CB_InsertLast_Click()
    Selection.Range.Text = answer
End Sub

Private Sub CB_Send_Click()
    Dim prompt As String
    Dim selText As String
    Dim fullPrompt As String
    Dim ok As Boolean
    Dim lines() As String
    
    ' 0. Загрузка настроек из реестра
    InitLLMSettings
    If LLM_API_URL = "" Or LLM_MODEL_ID = "" Or LLM_API_KEY = "" Then
        MsgBox "Не заданы настройки LLM (URL, MODEL, KEY). Сначала выполните ConfigureLLMSettings.", vbExclamation
        ConfigureLLMSettings
    End If
    
    ' 1. Получение запроса
    lines = Split(TB_Chat.Text, vbCrLf)
    prompt = lines(UBound(lines))
    If Trim(prompt) = "" Then
        MsgBox "Запрос содержит пустую строку!"
        Exit Sub
    End If
    Debug.Print prompt
    Me.CB_Send.Enabled = False
    
    
    ' 2. Добавляем выделенный текст (если есть)
    If Selection.Type = wdSelectionNormal And Selection.Range.Characters.Count > 0 Then
        selText = Selection.Text
    Else
        selText = ""
    End If

    If selText <> "" Then
        fullPrompt = prompt & vbCrLf & vbCrLf & _
                     "Выделенный фрагмент документа:" & vbCrLf & selText
    Else
        fullPrompt = prompt
    End If

    ' 3. Вызов модели
    ok = CallLLM(fullPrompt, answer)
    If Not ok Then
        MsgBox "Ошибка при обращении к LLM. Проверьте ключ и параметры API в коде макроса.", vbExclamation
        Exit Sub
    End If

    ' 5. Вставка ответа в форму
    Me.TB_Chat.Text = Me.TB_Chat.Text & vbCrLf & vbCrLf & answer & vbCrLf & vbCrLf

    ' 6. Блокирвоание кнопки отправки запроса (для защиты от случайной отправки)
    Me.CB_Send.Enabled = True
End Sub



