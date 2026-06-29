Attribute VB_Name = "LLM_work"
Option Explicit

'==== Настройки API (переменные хранятся в реестре) ====
Public LLM_API_URL As String
Public LLM_MODEL_ID As String
Public LLM_API_KEY As String
Public LLM_SYSTEM_PROMPT As String

' Ключи реестра (HKCU)
Public Const REG_ROOT As String = "HKEY_CURRENT_USER\Software\LLMWordMacro\"
Public Const REG_LLM_API_URL As String = REG_ROOT & "LLM_API_URL"
Public Const REG_LLM_MODEL_ID As String = REG_ROOT & "LLM_MODEL_ID"
Public Const REG_LLM_API_KEY As String = REG_ROOT & "LLM_API_KEY"
Public Const REG_LLM_SYSTEM_PROMPT As String = REG_ROOT & "LLM_SYSTEM_PROMPT"

' Заголовок формы
Public Const CHAT_LLM_FORM_CAPTION = "Чат с LLM"



'==== Основной макрос ====
Public Sub RunLLMQuery()
    Dim prompt As String
    Dim selText As String
    Dim fullPrompt As String
    Dim answer As String
    Dim ok As Boolean
    
    ' 0. Загрузка настроек из реестра
    InitLLMSettings
    If LLM_API_URL = "" Or LLM_MODEL_ID = "" Or LLM_API_KEY = "" Then
        MsgBox "Не заданы настройки LLM (URL, MODEL, KEY). Сначала выполните ConfigureLLMSettings.", vbExclamation
        ConfigureLLMSettings
'        Exit Sub
    End If
    
    ' 1. Диалог ввода запроса
    prompt = InputBox("Введите запрос для LLM:", "LLM чат-бот")
    If Trim(prompt) = "" Then Exit Sub
    
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
    
    ' 4. Подстановка ответа в документ
    Selection.Range.Text = answer
End Sub

'==== Показ формы чата ====
Public Sub RunLLMChat()
    LLM_chat.Show
End Sub

'==== Вызов LLM через HTTP ====
Public Function CallLLM(ByVal prompt As String, ByRef answer As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim payload As String
    Dim responseText As String
    Dim startTime As Single
    Dim waitMs As Long
    
    payload = BuildJsonPayload(prompt)
    
    ' ServerXMLHTTP avoids the nested COM message-pump problem
    ' that sync XMLHTTP has — no re-entry into Visio/LLM_chat events.
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "POST", LLM_API_URL, True
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Authorization", "Bearer " & LLM_API_KEY
    
    ' Timeouts (ms): resolve, connect, send, receive
    http.setTimeouts 10000, 10000, 30000, 120000
    
    ' Send payload
    http.send payload
    
    ' Poll readyState without freezing the UI + spinner
    Dim pollCounter As Long
    startTime = Timer
    Do While http.readyState <> 4
        DoEvents
'        Debug.Print http.responseText
        pollCounter = pollCounter + 1
        UpdateSpinner pollCounter
        ' Absolute ceiling — 3 minutes, in case server hangs
        If Timer - startTime > 180 Then
            Debug.Print "CallLLM: timeout (180s)"
            CallLLM = False
            Exit Function
        End If
    Loop

    If http.Status <> 200 Then
        Log "CallLLM: HTTP " & http.Status & " " & http.StatusText
        CallLLM = False
        Exit Function
    End If
    
    ' Get Response
    responseText = http.responseText
    answer = ExtractContentFromJson(responseText)
    
    CallLLM = (answer <> "")
    Exit Function
    
ErrHandler:
    CallLLM = False
End Function




' Формирование JSON-тела запроса
'Private Function BuildJsonPayload(ByVal prompt As String) As String
'    Dim esc As String
'    esc = JsonEscape(prompt)
'
'    ' Формат под ваш API (OpenAI-совместимый)
'    BuildJsonPayload = _
'        "{" & _
'        """model"":""" & LLM_MODEL_ID & """," & _
'        """messages"":[" & _
'            "{""role"":""user"",""content"":""" & esc & """}" & _
'        "]" & _
'        "}"
'End Function
Private Function BuildJsonPayload(ByVal prompt As String) As String
    Dim escUser As String
    Dim escSystem As String
    Dim messagesJson As String

    escUser = JsonEscape(prompt)
    escSystem = JsonEscape(LLM_SYSTEM_PROMPT)

    ' Формируем массв messages, опционально добавляя системный промпт
    If Trim(LLM_SYSTEM_PROMPT) <> "" Then
        messagesJson = _
            "{""role"":""system"",""content"":""" & escSystem & """}," & _
            "{""role"":""user"",""content"":""" & escUser & """}"
    Else
        messagesJson = _
            "{""role"":""user"",""content"":""" & escUser & """}"
    End If

    ' Формат под ваш API (OpenAI-совместимый)
    BuildJsonPayload = _
        "{" & _
        """model"":""" & LLM_MODEL_ID & """," & _
        """messages"":[" & messagesJson & "]" & _
        "}"
End Function

' Простейший экранировщик для JSON-строки
Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, Chr(34), "\" & Chr(34))
    s = Replace(s, vbBack, "\b")
    s = Replace(s, vbFormFeed, "\f")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    ' Удаляем остальные управляющие символы (ASCII 0–31, кроме вышеперечисленных)
    Dim i As Long
    For i = 0 To 31
        If InStr(s, Chr(i)) > 0 Then
            s = Replace(s, Chr(i), "\u" & Right$("0000" & Hex(AscW(Chr(i))), 4))
        End If
    Next i
    JsonEscape = s
End Function

' ==== Разбор JSON-ответа ====
' Ожидаем структуру как в вашем примере:
' {
'   "choices":[
'     {
'       "message":{
'         "content":"...текст ответа..."
'       }
'     }
'   ]
'   ...
' }
Private Function ExtractContentFromJson(ByVal json As String) As String
    Dim key As String
    Dim pos As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim tmp As String
    
    ' 1. Находим блок "message":{"role":...,"content":"..."}
    key = """message"":{"
    pos = InStr(1, json, key, vbTextCompare)
    If pos = 0 Then
        ExtractContentFromJson = ""
        Exit Function
    End If
    
    ' 2. Отрезаем всё до "message":{, чтобы сократить строку
    tmp = Mid$(json, pos + Len(key))
    
    ' 3. Внутри этого блока ищем "content":"..."
    key = """content"":"""
    pos = InStr(1, tmp, key, vbTextCompare)
    If pos = 0 Then
        ExtractContentFromJson = ""
        Exit Function
    End If
    
    startPos = pos + Len(key)
    endPos = startPos
    
    ' 4. Ищем завершающую кавычку, учитывая возможные экранированные \"
    Do While endPos <= Len(tmp)
        If Mid$(tmp, endPos, 1) = """" Then
            ' Проверяем, не экранирована ли кавычка
            If Mid$(tmp, endPos - 1, 1) <> "\" Then
                Exit Do
            End If
        End If
        endPos = endPos + 1
    Loop
    
    If endPos > Len(tmp) Then
        ExtractContentFromJson = ""
        Exit Function
    End If
    
    ExtractContentFromJson = JsonUnescape(Mid$(tmp, startPos, endPos - startPos))
End Function

' Обратное преобразование для \n, \" и \\
Private Function JsonUnescape(ByVal s As String) As String
    ' \\ -> \
    s = Replace(s, "\\", Chr(92))
    ' \" -> "
    s = Replace(s, "\" & Chr(34), Chr(34))
    ' \n -> CRLF
    s = Replace(s, "\n", vbCrLf)
    JsonUnescape = s
End Function


'==== Spinner indicator for LLM_chat form ====
' Cycles the last character of the form caption through [/ \ - |]
' each time `counter` is a multiple of couunter_step.
Public Sub UpdateSpinner(ByVal counter As Long, Optional ByVal couunter_step As Integer = 1000)
    Static idx As Long
    Const chars = "/-\|/-\|"

    ' Advance only on couunter_step-multiples of counter
    If counter Mod couunter_step = 0 Then
        idx = (idx + 1) Mod Len(chars)
        LLM_chat.Caption = CHAT_LLM_FORM_CAPTION & Mid$(chars, idx + 1, 1)
    End If
End Sub







'==== Работа с данными настроек ====
' Чтение настроек из реестра
Public Sub InitLLMSettings()
    On Error Resume Next
    LLM_API_URL = CStr(GetSettingFromRegistry(REG_LLM_API_URL, "https://api.aitunnel.ru/v1/chat/completions"))
    LLM_MODEL_ID = CStr(GetSettingFromRegistry(REG_LLM_MODEL_ID, "gpt-5.1"))
    LLM_API_KEY = CStr(GetSettingFromRegistry(REG_LLM_API_KEY, ""))
    LLM_SYSTEM_PROMPT = CStr(GetSettingFromRegistry(REG_LLM_SYSTEM_PROMPT, ""))
    On Error GoTo 0
End Sub

' Универсальное чтение значения (если нет - возвращает defaultValue)
Private Function GetSettingFromRegistry(ByVal fullKey As String, ByVal defaultValue As String) As String
    Dim shell As Object
    Dim val As Variant

    On Error Resume Next
    Set shell = CreateObject("WScript.Shell")
    val = shell.RegRead(fullKey)
    If Err.Number <> 0 Then
        Err.Clear
        GetSettingFromRegistry = defaultValue
    Else
        GetSettingFromRegistry = CStr(val)
    End If
    On Error GoTo 0
End Function

' Универсальная запись значения в реестр (строка, REG_SZ)
Private Sub SaveSettingToRegistry(ByVal fullKey As String, ByVal value As String)
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.RegWrite fullKey, value, "REG_SZ"
End Sub



' ==== Настройка параметров LLM через диалоги ====
' Основная функция конфигурации: спрашивает у пользователя значения и
' записывает их в реестр HKCU\Software\LLMWordMacro
Public Sub ConfigureLLMSettings()
    Dim url As String
    Dim modelId As String
    Dim apiKey As String
    Dim sysPrompt As String

    ' текущие значения (если уже заданы)
    InitLLMSettings

    url = InputBox( _
        prompt:="Введите URL API LLM (например, https://api.openai.com/v1/chat/completions):", _
        Title:="Настройка LLM_API_URL", _
        Default:=IIf(LLM_API_URL <> "", LLM_API_URL, "https://api.openai.com/v1/chat/completions") _
    )
    If url = "" Then Exit Sub

    modelId = InputBox( _
        prompt:="Введите идентификатор модели (например, gpt-5.1):", _
        Title:="Настройка LLM_MODEL_ID", _
        Default:=IIf(LLM_MODEL_ID <> "", LLM_MODEL_ID, "gpt-5.1") _
    )
    If modelId = "" Then Exit Sub

    apiKey = InputBox( _
        prompt:="Введите API-ключ (Bearer токен) для LLM:", _
        Title:="Настройка LLM_API_KEY", _
        Default:=LLM_API_KEY _
    )
    If apiKey = "" Then Exit Sub

    sysPrompt = InputBox( _
        prompt:="Введите системный промпт (необязательно). Он будет отправляться как роль system перед сообщением пользователя.", _
        Title:="Настройка LLM_SYSTEM_PROMPT", _
        Default:=LLM_SYSTEM_PROMPT _
    )
    ' sysPrompt может быть пустым — это допустимо

    ' сохраняем в реестр текущего пользователя
    SaveSettingToRegistry REG_LLM_API_URL, url
    SaveSettingToRegistry REG_LLM_MODEL_ID, modelId
    SaveSettingToRegistry REG_LLM_API_KEY, apiKey
    SaveSettingToRegistry REG_LLM_SYSTEM_PROMPT, sysPrompt

    ' сразу обновляем внутренние переменные
    InitLLMSettings

    MsgBox "Настройки LLM сохранены в реестре:" & vbCrLf & _
           "HKCU\Software\LLMWordMacro\" & vbCrLf & _
           "LLM_API_URL, LLM_MODEL_ID, LLM_API_KEY, LLM_SYSTEM_PROMPT", _
           vbInformation
End Sub

