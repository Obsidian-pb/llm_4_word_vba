VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHarness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsHarness - Main LLM agent class (message history, tool calling loop, HTTP communication)
Option Explicit

' ============================================================
'  Harness single class
' ============================================================

' ============================================================
'  PROPERTIES
' ============================================================
Private history As Collection       ' Message history — collection of JSON strings (each string = one message).
Private agent_max_turns As Integer  ' Maximum number of Agent turns
Private tick_interval As Long       ' Work indicator event interval
'==== API settings (variables stored in registry) ====
Public llm_api_url As String
Public llm_model_id As String
Public llm_api_key As String
Public llm_system_prompt As String

' ============================================================
'  EVENTS
' ============================================================
' Agent is thinking
Public Event OnThinking()
' Agent's answer
Public Event OnAnswer(ByVal s As String)
' Agent's reasoning
Public Event OnReasoning(ByVal s As String)
' Result of tool calls
Public Event OnToolCalling(ByVal s As String)
' System message
Public Event OnSystem(ByVal s As String)
' Work indicator event
Public Event OnTick()

' ============================================================
'  MEMBERS
' ============================================================
Private tools_handler As clsHarnessToolsHandler


' ============================================================
'  AGENT INIT
' ============================================================
Public Function Init(Optional ByRef t_handler As clsHarnessToolsHandler = Nothing) As clsHarness
    
    ' Tools handler class
    If t_handler Is Nothing Then
        Set tools_handler = New clsHarnessToolsHandler
    Else
        Set tools_handler = t_handler
    End If
    tools_handler.Init
    
    ' Message history collection
    Set history = New Collection

Set Init = Me
End Function


' ============================================================
'  AGENT CONFIGURATION
' ============================================================
Public Function Set_LLM(ByVal api_url As String, _
                ByVal model_id As String, _
                ByVal api_key As String, _
                ByVal system_prompt As String, _
                Optional ByVal max_turns As Integer = 30, _
                Optional ByVal tick_intrvl As Integer = 10000) As clsHarness
    llm_api_url = api_url
    
    llm_model_id = model_id
    llm_api_key = api_key
    If llm_system_prompt = "" Then
        llm_system_prompt = system_prompt
    End If
    agent_max_turns = max_turns
    tick_interval = tick_intrvl
    
    history.Add BuildMessageJson("system", llm_system_prompt)
    
    Set Set_LLM = Me
End Function


' ============================================================
'  MAIN AGENT LOOP
'  Stores message history and loops until LLM produces
'  a regular text response (without tool_calls).
' ============================================================
Public Sub AgentLoop(ByVal userPrompt As String)
Dim turn As Long
Dim messagesJson As String
Dim responseJson As String
Dim toolCalls As Collection
Dim finalReasoning As String
Dim assistantMsg As String
Dim tc As Variant


    ' If message history has not been created yet — create it
    If history Is Nothing Then
        Set history = New Collection
    End If

    ' 0. Add the system prompt
    If history.Count = 0 Then
        history.Add BuildMessageJson("system", llm_system_prompt)
    End If

    ' 1. Add the first user message.
    history.Add BuildMessageJson("user", userPrompt)

    For turn = 1 To agent_max_turns
        RaiseEvent OnThinking
        ' 2. Collect history into a JSON array and send to LLM.
        messagesJson = CollectionToJsonArray(history)

        '!!! Agent call here!
        responseJson = CallLLMOnce(messagesJson)

        ' 2.0. If the model returned an empty response
        If responseJson = "" Then
            RaiseEvent OnAnswer("LLM returned an empty response. Aborting loop.")
            Exit Sub
        End If

        ' 3. Try to extract tool_calls from the response.
        Set toolCalls = ParseToolCalls(responseJson)

        If toolCalls Is Nothing Then  'Or toolCalls.Count = 0
            ' No tools called — this is the final text response.
            finalReasoning = ExtractReasoning(responseJson)
            If finalReasoning <> "" Then RaiseEvent OnReasoning(finalReasoning)

            RaiseEvent OnAnswer(ExtractAssistantContent(responseJson))
            
            assistantMsg = ExtractAssistantMessageJson(responseJson)
            history.Add assistantMsg
            Exit Sub
        End If

        ' 4. Tools are present. First, add the full assistant message
        '    (with all tool_calls) to history so that LLM remembers the context.
        
        assistantMsg = ExtractAssistantMessageJson(responseJson)
        history.Add assistantMsg

        ' 5. Execute each tool call in turn.
        For Each tc In toolCalls
            Dim callId   As String: callId = tc("id")
            Dim funcName As String: funcName = tc("name")
            Dim funcArgs As String: funcArgs = tc("arguments")

            ' Log that a tool is being called
            Dim reasoningText As String
            reasoningText = ExtractReasoning(responseJson)
            If reasoningText <> "" Then RaiseEvent OnReasoning(reasoningText)

            ' Dispatcher: find the function by name and execute it.
            Dim toolResult As String
            toolResult = tools_handler.Execute(funcName, funcArgs)
            RaiseEvent OnToolCalling(toolResult)

            ' 6. Add the result back to history as a role=tool message.
            history.Add BuildToolMessageJson(callId, toolResult)
        Next tc

        ' The loop continues: resend history to LLM and see
        ' if it wants to call something else or is ready to answer with text.
    Next turn

    RaiseEvent OnAnswer("[Agent] Turn limit exceeded (" & agent_max_turns & ").")
End Sub


' ============================================================
'  SINGLE HTTP REQUEST TO LLM
'  Differs from regular CallLLM in that it passes the "tools" field
'  and does NOT include streaming (for the agent, streaming only
'  interferes with parsing tool_calls).
' ============================================================
Private Function CallLLMOnce(ByVal messagesJson As String) As String
Dim startTime As Variant

    On Error GoTo ErrHandler

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP")

    http.Open "POST", llm_api_url, True   ' synchronous mode is simpler for the agent, but currently async mode
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Authorization", "Bearer " & llm_api_key
    http.setTimeouts 10000, 10000, 30000, 120000

    ' Build the request body: model + messages + tool descriptions.
    ' "stream":false is explicitly set to get a complete JSON.
    Dim payload As String
    payload = "{" & _
        """model"":""" & llm_model_id & """," & _
        """stream"":false," & _
        """messages"":" & messagesJson & "," & _
        """tools"":" & tools_handler.GetToolsJson() & _
    "}"

    http.send payload

    Dim i As Long: i = 0
    startTime = Now
    Do While http.readyState <> 4
        DoEvents

        ' HI-05: use DateDiff on the full date-time value (Now) instead of
        ' Timer-arithmetic difference, which turns negative (and is unreliable)
        ' when the system clock crosses midnight during the wait.
        If DateDiff("s", startTime, Now) > 180 Then
            RaiseEvent OnSystem("Timeout 180s")
            CallLLMOnce = ""
            Exit Function
        End If
        
        i = i + 1
        If i > tick_interval Then
            i = 0
            RaiseEvent OnTick
        End If
    Loop

    If http.status <> 200 Then
        Log "[CallLLMOnce] HTTP " & http.status & " " & http.statusText
        Log http.responseText
        RaiseEvent OnSystem("HTTP " & http.status & " " & http.statusText)
        CallLLMOnce = ""
        Exit Function
    End If

    CallLLMOnce = http.responseText
    Exit Function

ErrHandler:
    Log "[CallLLMOnce] Error: " & Err.Description
    CallLLMOnce = ""
End Function

' ============================================================
'      CONTEXT CLEARING
' ============================================================
Public Sub ClearContext()
    Set history = New Collection
    RaiseEvent OnSystem("Context cleared")
End Sub

' ============================================================
'  ADDING REFERENCE CONTEXT (RAG)
'  Document content is added to history with role "system".
'  The model will use it as a source of facts when answering.
' ============================================================
Public Sub AddContext(ByVal sourceName As String, ByVal content As String)
    Dim contextMsg As String
    contextMsg = "Below is a reference document (source: " & sourceName & ")." & vbCrLf & _
                 "Use it as a source of factual data when answering user questions. " & _
                 "If the required answer is not in the document — honestly say so, do not make things up." & vbCrLf & vbCrLf & _
                 content
    history.Add BuildMessageJson("system", contextMsg)
    RaiseEvent OnSystem("Reference context added: " & sourceName)
End Sub


' ============================================================
'      HELPER FUNCTIONS FOR WORKING WITH LLM RESPONSES
' ============================================================

' Extracts choices[0].message.content (final text) from the LLM response.
' MD-02: primary extraction uses structural JSON parsing (JsonConverter),
' which is immune to nested fields, whitespace, escaped backslashes and
' alternative API layouts. The old InStr manual search remains only as
' an emergency fallback if the JSON cannot be parsed.
Private Function ExtractAssistantContent(ByVal json As String, Optional ByVal key As String = """content"":") As String
    On Error GoTo Fallback

    Dim root As Object
    Dim choices As Collection
    Dim choice As Object
    Dim message As Object
    Dim val As Variant

    Set root = JsonConverter.ParseJson(json)
    If root Is Nothing Then GoTo Fallback

    If root.Exists("choices") Then
        Set choices = root("choices")
        If choices.Count > 0 Then
            Set choice = choices(1)
            If choice.Exists("message") Then
                Set message = choice("message")
                If message.Exists("content") Then
                    val = message("content")
                    If IsNull(val) Or IsEmpty(val) Then
                        ExtractAssistantContent = ""
                    Else
                        ExtractAssistantContent = CStr(val)
                    End If
                    Exit Function
                End If
            End If
        End If
    End If

    ' No content found structurally — try the manual fallback.
    GoTo Fallback

Fallback:
    ExtractAssistantContent = ExtractContent_Manual(json, key)
End Function


' Manual InStr fallback used only when structural JSON parsing fails
' (previous logic preserved as an emergency path).
Private Function ExtractContent_Manual(ByVal json As String, ByVal key As String) As String
    Dim pos As Long, startPos As Long, endPos As Long
    pos = InStr(1, json, key, vbTextCompare)
    If pos = 0 Then ExtractContent_Manual = "": Exit Function

    startPos = pos + Len(key)
    ' Skip spaces and opening quote
    Do While startPos <= Len(json) And (Mid$(json, startPos, 1) = " " Or Mid$(json, startPos, 1) = """")
        If Mid$(json, startPos, 1) = """" Then startPos = startPos + 1: Exit Do
        startPos = startPos + 1
    Loop

    ' If content was null — return empty string
    If LCase$(Mid$(json, startPos - 1, 4)) = "null" Then
        ExtractContent_Manual = ""
        Exit Function
    End If

    endPos = startPos
    Do While endPos <= Len(json)
        If Mid$(json, endPos, 1) = """" And Mid$(json, endPos - 1, 1) <> "\" Then Exit Do
        endPos = endPos + 1
    Loop

    ExtractContent_Manual = JsonUnescape(Mid$(json, startPos, endPos - startPos))
End Function


' Extracts the reasoning field ("reasoning_content" for DeepSeek,
' "reasoning" for other models).
' MD-02: primary extraction uses structural JSON parsing and checks both
' the standard location choices[0].message.<key> and top-level <key>.
' The old InStr manual search remains only as a fallback if parsing fails.
Private Function ExtractReasoning(ByVal json As String) As String
    On Error GoTo Fallback

    Dim root As Object
    Dim choices As Collection
    Dim choice As Object
    Dim message As Object
    Dim candidates As Variant
    Dim i As Long
    Dim val As Variant

    Set root = JsonConverter.ParseJson(json)
    If root Is Nothing Then GoTo Fallback

    candidates = Array("reasoning_content", "reasoning")

    For i = LBound(candidates) To UBound(candidates)
        ' Some APIs emit the reasoning field at the top level of the response.
        If root.Exists(candidates(i)) Then
            val = root(candidates(i))
            If Not (IsNull(val) Or IsEmpty(val)) Then
                ExtractReasoning = CStr(val)
                Exit Function
            End If
        End If

        ' Standard OpenAI-compatible location: choices[0].message.<key>
        If root.Exists("choices") Then
            Set choices = root("choices")
            If choices.Count > 0 Then
                Set choice = choices(1)
                If choice.Exists("message") Then
                    Set message = choice("message")
                    If message.Exists(candidates(i)) Then
                        val = message(candidates(i))
                        If Not (IsNull(val) Or IsEmpty(val)) Then
                            ExtractReasoning = CStr(val)
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' Not found structurally — try the manual fallback.
    GoTo Fallback

Fallback:
    ExtractReasoning = ExtractReasoning_Manual(json)
End Function


' Manual InStr fallback used only when structural JSON parsing fails
' (previous logic preserved as an emergency path).
Private Function ExtractReasoning_Manual(ByVal json As String) As String
    Dim candidates As Variant
    Dim i As Long
    candidates = Array("""reasoning_content"":", """reasoning"":")

    For i = LBound(candidates) To UBound(candidates)
        Dim key As String
        key = candidates(i)

        Dim pos As Long
        pos = InStr(1, json, key, vbTextCompare)
        If pos > 0 Then
            Dim startPos As Long, endPos As Long
            startPos = pos + Len(key)

            ' Skip spaces
            Do While startPos <= Len(json) And Mid$(json, startPos, 1) = " "
                startPos = startPos + 1
            Loop

            ' If null — skip
            If LCase$(Mid$(json, startPos, 4)) = "null" Then
                ExtractReasoning_Manual = ""
                Exit Function
            End If

            ' If it starts with a quote — extract the string
            If Mid$(json, startPos, 1) = """" Then
                startPos = startPos + 1
                endPos = startPos
                Do While endPos <= Len(json)
                    If Mid$(json, endPos, 1) = """" And Mid$(json, endPos - 1, 1) <> "\" Then Exit Do
                    endPos = endPos + 1
                Loop
                ExtractReasoning_Manual = JsonUnescape(Mid$(json, startPos, endPos - startPos))
                Exit Function
            End If
        End If
    Next i

    ExtractReasoning_Manual = ""
End Function


' Parses the tool_calls array from the LLM response.
' Returns a Collection where each element is a Dictionary
' with keys "id", "name", "arguments".
' Parses the tool_calls array from a complete LLM response.
' The response is parsed as JSON first; brackets inside JSON strings
' therefore cannot affect the traversal.
' Returns a Collection where each element is a Dictionary
' with keys "id", "name", "arguments".
Private Function ParseToolCalls(ByVal json As String) As Collection
    Dim root As Object
    Dim choices As Collection
    Dim choice As Object
    Dim message As Object
    Dim toolCalls As Collection
    Dim callValue As Variant
    Dim toolCall As Object
    Dim func As Object
    Dim result As Collection
    Dim item As Object
    Dim value As Variant
    Dim s As String

    On Error GoTo ParseError

    ' Parse the complete response. Do not search for delimiters in the text.
    Set root = JsonConverter.ParseJson(json)

    If Not root.Exists("choices") Then GoTo NoToolCalls
    Set choices = root("choices")
    If choices.Count = 0 Then GoTo NoToolCalls

    Set choice = choices(1)
    If Not choice.Exists("message") Then GoTo NoToolCalls
    Set message = choice("message")
    If Not message.Exists("tool_calls") Then GoTo NoToolCalls

    Set toolCalls = message("tool_calls")
    If toolCalls.Count = 0 Then GoTo NoToolCalls

    Set result = New Collection
    For Each callValue In toolCalls
        Set toolCall = callValue
        If toolCall.Exists("function") Then
            Set func = toolCall("function")
            If func.Exists("name") Then
                Set item = CreateObject("Scripting.Dictionary")

                If toolCall.Exists("id") Then
                    value = toolCall("id")
                    If IsNull(value) Then item("id") = "" Else item("id") = CStr(value)
                Else
                    item("id") = ""
                End If

                value = func("name")
                If IsNull(value) Then item("name") = "" Else item("name") = CStr(value)

                ' arguments is a JSON string and must remain a string.
                If func.Exists("arguments") Then
                    value = func("arguments")
                    If IsNull(value) Then
                        s = ""
                    Else
                        s = CStr(value)
                    End If
                Else
                    s = ""
                End If
                item("arguments") = s
                result.Add item
            End If
        End If
    Next callValue

    If result.Count = 0 Then GoTo NoToolCalls
    Set ParseToolCalls = result
    Exit Function

NoToolCalls:
    Set ParseToolCalls = Nothing
    Exit Function

ParseError:
    Log "[ParseToolCalls] Invalid LLM JSON: " & Err.Description
    Set ParseToolCalls = Nothing
End Function
' ============================================================
' Builds a JSON message {role, content}.
Private Function BuildMessageJson(ByVal role As String, ByVal content As String) As String
    BuildMessageJson = "{" & _
        """role"":""" & role & """," & _
        """content"":" & JsonString(content) & _
    "}"
End Function

' Builds a JSON message with role=tool and the tool result.
' MD-03: the tool_call_id is escaped via JsonString so that any
' non-standard identifier (quotes, backslashes, etc.) cannot break JSON.
Private Function BuildToolMessageJson(ByVal toolCallId As String, ByVal content As String) As String
    BuildToolMessageJson = "{" & _
        """role"":""tool""," & _
        """tool_call_id"":" & JsonString(toolCallId) & "," & _
        """content"":" & JsonString(content) & _
    "}"
End Function

' Extracts the complete assistant message JSON from the response
' (to save it to history along with tool_calls).
Private Function ExtractAssistantMessageJson(ByVal json As String) As String
    Dim key As String, pos As Long
    key = """message"":"
    pos = InStr(1, json, key, vbTextCompare)
    If pos = 0 Then ExtractAssistantMessageJson = "{}": Exit Function

    Dim startObj As Long
    startObj = InStr(pos + Len(key), json, "{")
    If startObj = 0 Then ExtractAssistantMessageJson = "{}": Exit Function

    ' Find the matching closing brace, accounting for nesting.
    Dim depth As Long, i As Long
    depth = 0
    For i = startObj To Len(json)
        Select Case Mid$(json, i, 1)
            Case "{": depth = depth + 1
            Case "}": depth = depth - 1
                If depth = 0 Then
                    ExtractAssistantMessageJson = Mid$(json, startObj, i - startObj + 1)
                    Exit Function
                End If
        End Select
    Next i
    ExtractAssistantMessageJson = "{}"
End Function

