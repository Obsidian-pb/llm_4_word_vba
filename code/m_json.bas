Attribute VB_Name = "m_json"
' Module: m_json - JSON utility functions based on VBA-JSON (JsonConverter)
Option Explicit

' Converts a Collection of JSON strings into a JSON array [...].
Public Function CollectionToJsonArray(ByVal col As Collection) As String
    Dim i As Long
    Dim result As String

    result = "["
    For i = 1 To col.Count
        result = result & col.item(i)
        If i < col.Count Then result = result & ","
    Next i
    result = result & "]"

    CollectionToJsonArray = result
End Function

' Wraps a string in JSON quotes with proper escaping.
' Requires JsonConverter module from VBA-JSON.
Public Function JsonString(ByVal s As String) As String
    JsonString = JsonConverter.ConvertToJson(s)
End Function

' Builds a JSON response object {"status":"ok","message":<escaped msg>}.
' The message is escaped via JsonString, so quotes, backslashes,
' and newlines inside the message cannot break the JSON structure.
Public Function MakeOk(ByVal msg As String) As String
    MakeOk = "{""status"":""ok"",""message"":" & JsonString(msg) & "}"
End Function

' Builds a JSON response object {"status":"error","message":<escaped msg>}.
' The message is escaped via JsonString.
Public Function MakeError(ByVal msg As String) As String
    MakeError = "{""status"":""error"",""message"":" & JsonString(msg) & "}"
End Function

' Extracts a field value from a JSON object and returns it as string.
' Supports string/number/boolean/null values.
' For object/array values returns their JSON representation.
' Search is recursive: if the key is nested, the first found value is returned.
Public Function ExtractJsonString(ByVal json As String, ByVal key As String) As String
    On Error GoTo ParseError

    Dim root As Object
    Dim found As Boolean
    Dim foundValue As Variant

    Set root = JsonConverter.ParseJson(json)
    If root Is Nothing Then Exit Function

    Call TryFindKeyValue(root, key, foundValue, found)
    If Not found Then Exit Function

    ExtractJsonString = JsonValueToString(foundValue)
    Exit Function

ParseError:
    ExtractJsonString = ""
End Function

Private Sub TryFindKeyValue(ByVal node As Variant, ByVal key As String, ByRef foundValue As Variant, ByRef found As Boolean)
    If found Then Exit Sub

    If Not IsObject(node) Then Exit Sub

    Select Case typeName(node)
        Case "Dictionary"
            Dim dict As Object
            Dim dictKey As Variant

            Set dict = node
            If dict.Exists(key) Then
                foundValue = dict.item(key)
                found = True
                Exit Sub
            End If

            For Each dictKey In dict.Keys
                Call TryFindKeyValue(dict.item(dictKey), key, foundValue, found)
                If found Then Exit Sub
            Next dictKey

        Case "Collection"
            Dim item As Variant
            For Each item In node
                Call TryFindKeyValue(item, key, foundValue, found)
                If found Then Exit Sub
            Next item
    End Select
End Sub

' Unescapes JSON string content (reverse escaping).
' Input is expected without outer quotes, e.g. line\ntext
Public Function JsonUnescape(ByVal s As String) As String
    On Error GoTo Fallback

    Dim wrapper As String
    Dim root As Object

    wrapper = "{""v"":""" & s & """}"
    Set root = JsonConverter.ParseJson(wrapper)

    If Not root Is Nothing Then
        If root.Exists("v") Then
            If IsNull(root.item("v")) Then
                JsonUnescape = ""
            Else
                JsonUnescape = CStr(root.item("v"))
            End If
            Exit Function
        End If
    End If

Fallback:
    JsonUnescape = FallbackJsonUnescape(s)
End Function

Private Function JsonValueToString(ByVal value As Variant) As String
    If IsObject(value) Then
        If value Is Nothing Then
            JsonValueToString = ""
        Else
            JsonValueToString = JsonConverter.ConvertToJson(value)
        End If
        Exit Function
    End If

    If IsNull(value) Or IsEmpty(value) Then
        JsonValueToString = ""
        Exit Function
    End If

    Select Case VarType(value)
        Case vbBoolean
            If CBool(value) Then
                JsonValueToString = "true"
            Else
                JsonValueToString = "false"
            End If
        Case vbString
            JsonValueToString = CStr(value)
        Case vbDate
            JsonValueToString = Format$(CDate(value), "yyyy-mm-dd\THH:nn:ss")
        Case Else
            JsonValueToString = Replace$(CStr(value), ",", ".")
    End Select
End Function

Private Function FallbackJsonUnescape(ByVal s As String) As String
    Const MARKER As String = "??BS??"

    s = Replace(s, "\\", MARKER)
    s = Replace(s, "\" & Chr$(34), Chr$(34))
    s = Replace(s, "\n", vbCrLf)
    s = Replace(s, "\r", vbCr)
    s = Replace(s, "\t", vbTab)
    s = Replace(s, MARKER, "\")

    FallbackJsonUnescape = s
End Function
