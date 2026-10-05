Attribute VB_Name = "m_Profiles"
' Module: m_Profiles - LLM connection profiles stored in a single YAML file
' The profile file is located in the same folder as the document that hosts
' the VBA project. Each profile stores:
'   - name          : profile name
'   - url           : LLM access URL
'   - api_key       : LLM API key
'   - model         : model name
'   - system_prompt : system prompt
' The file is a valid YAML document and is written/read by this module only.
Option Explicit

' Name of the YAML file that keeps all LLM connection profiles.
Private Const PROFILES_FILE_NAME As String = "llm_profiles.yaml"


' ============================================================
'  PUBLIC API
' ============================================================

' Returns the full path to the profiles YAML file.
' The file is placed in the same folder as the active document
' (the document that hosts this VBA project).
Public Function GetProfilesFilePath() As String
    Dim basePath As String

    On Error Resume Next
    basePath = ThisDocument.path
    On Error GoTo 0

    ' Fallback to the current directory if the document has not been saved yet.
    If basePath = "" Then basePath = CurDir$

    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"

    GetProfilesFilePath = basePath & PROFILES_FILE_NAME
End Function

' Saves profile settings to the YAML file.
' If a profile with the same name already exists, it is updated in place;
' otherwise a new profile is appended to the file.
Public Sub SaveLLMProfile(ByVal profileName As String, _
                          ByVal apiUrl As String, _
                          ByVal apiKey As String, _
                          ByVal modelId As String, _
                          ByVal maxTokens As String, _
                          ByVal temperature As String, _
                          ByVal systemPrompt As String)
    Dim filePath As String
    Dim profiles As Collection
    Dim prof As Object
    Dim idx As Long

    profileName = Trim$(profileName)
    If profileName = "" Then
        Err.Raise vbObjectError + 1001, "SaveLLMProfile", "Profile name must not be empty"
    End If

    filePath = GetProfilesFilePath()

    ' Load the existing profiles (if the file exists) to keep other profiles intact.
    If FileExists(filePath) Then
        Set profiles = ParseProfilesYaml(ReadFileUTF8(filePath))
    Else
        Set profiles = New Collection
    End If

    idx = FindProfileIndex(profiles, profileName)

    If idx > 0 Then
        ' Update the existing profile (its position in the file is preserved).
        Set prof = profiles(idx)
    Else
        ' Create a new profile and append it to the end of the list.
        Set prof = CreateObject("Scripting.Dictionary")
        profiles.Add prof
    End If

    prof("name") = profileName
    prof("url") = apiUrl
    prof("api_key") = apiKey
    prof("model") = modelId
    prof("max_tokens") = maxTokens
    prof("temperature") = temperature
    prof("system_prompt") = systemPrompt

    WriteFileUTF8 filePath, SerializeProfilesYaml(profiles)
End Sub

' Returns a profile as a Dictionary with the keys:
' "name", "url", "api_key", "model", "system_prompt".
' Returns Nothing if the profile is not found.
Public Function GetLLMProfile(ByVal profileName As String) As Object
    Dim filePath As String
    Dim profiles As Collection
    Dim idx As Long

    filePath = GetProfilesFilePath()
    If Not FileExists(filePath) Then Exit Function

    Set profiles = ParseProfilesYaml(ReadFileUTF8(filePath))
    idx = FindProfileIndex(profiles, profileName)
    If idx > 0 Then
        Set GetLLMProfile = profiles(idx)
    End If
End Function

' Convenience wrapper that returns the settings of a profile field by field.
' Returns True if the profile was found, False otherwise.
Public Function GetLLMProfileParams(ByVal profileName As String, _
                                    ByRef apiUrl As String, _
                                    ByRef apiKey As String, _
                                    ByRef modelId As String, _
                                    ByRef maxTokens As String, _
                                    ByRef temperature As String, _
                                    ByRef systemPrompt As String) As Boolean
    Dim prof As Object

    apiUrl = ""
    apiKey = ""
    modelId = ""
    maxTokens = ""
    temperature = ""
    systemPrompt = ""

    Set prof = GetLLMProfile(profileName)
    If prof Is Nothing Then Exit Function

    apiUrl = ProfileValue(prof, "url")
    apiKey = ProfileValue(prof, "api_key")
    modelId = ProfileValue(prof, "model")
    maxTokens = ProfileValue(prof, "max_tokens")
    temperature = ProfileValue(prof, "temperature")
    systemPrompt = ProfileValue(prof, "system_prompt")

    GetLLMProfileParams = True
End Function

' Returns a Collection with the names of all saved profiles (in file order).
Public Function GetProfileNames() As Collection
    Dim filePath As String
    Dim profiles As Collection
    Dim names As Collection
    Dim prof As Object

    Set names = New Collection

    filePath = GetProfilesFilePath()
    If FileExists(filePath) Then
        Set profiles = ParseProfilesYaml(ReadFileUTF8(filePath))
        For Each prof In profiles
            names.Add ProfileValue(prof, "name")
        Next prof
    End If

    Set GetProfileNames = names
End Function

' Deletes a profile by name. Returns True if a profile was removed.
Public Function DeleteLLMProfile(ByVal profileName As String) As Boolean
    Dim filePath As String
    Dim profiles As Collection
    Dim idx As Long

    filePath = GetProfilesFilePath()
    If Not FileExists(filePath) Then Exit Function

    Set profiles = ParseProfilesYaml(ReadFileUTF8(filePath))
    idx = FindProfileIndex(profiles, profileName)
    If idx = 0 Then Exit Function

    profiles.Remove idx
    WriteFileUTF8 filePath, SerializeProfilesYaml(profiles)
    DeleteLLMProfile = True
End Function


' ============================================================
'  YAML PARSING / SERIALIZATION (private helpers)
' ============================================================

' Parses the YAML content and returns a Collection of Dictionaries.
' Each Dictionary represents a single profile.
' The parser is line based: a line starting with "-" begins a new profile,
' every following "key: value" line fills the current profile.
Private Function ParseProfilesYaml(ByVal content As String) As Collection
    Dim profiles As Collection
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim current As Object
    Dim colonPos As Long
    Dim key As String
    Dim val As String

    Set profiles = New Collection
    Set current = Nothing

    ' Normalize line endings so the split below is predictable.
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    lines = Split(content, vbLf)

    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))

        ' Skip empty lines and comments.
        If line = "" Then GoTo NextLine
        If Left$(line, 1) = "#" Then GoTo NextLine

        ' A list item starts a new profile.
        If Left$(line, 1) = "-" Then
            Set current = CreateObject("Scripting.Dictionary")
            profiles.Add current
            line = Trim$(Mid$(line, 2))
            If line = "" Then GoTo NextLine
        End If

        ' Split "key: value" at the first colon.
        colonPos = InStr(line, ":")
        If colonPos = 0 Then GoTo NextLine

        key = LCase$(Trim$(Left$(line, colonPos - 1)))
        val = ParseYamlValue(Trim$(Mid$(line, colonPos + 1)))

        ' Store only known keys and only when a profile is already open.
        If Not current Is Nothing Then
            Select Case key
                Case "name", "url", "api_key", "model", "max_tokens", "temperature", "system_prompt"
                    current(key) = val
            End Select
        End If

NextLine:
    Next i

    Set ParseProfilesYaml = profiles
End Function

' Converts a raw YAML scalar into a plain string.
' Values produced by this module are JSON-escaped double-quoted strings.
Private Function ParseYamlValue(ByVal raw As String) As String
    If raw = "" Then Exit Function

    ' Null-like scalars map to an empty string.
    If raw = "~" Then Exit Function
    If LCase$(raw) = "null" Then Exit Function

    ' Double-quoted scalar: strip the outer quotes and unescape the content.
    If Len(raw) >= 2 Then
        If Left$(raw, 1) = """" And Right$(raw, 1) = """" Then
            ParseYamlValue = JsonUnescape(Mid$(raw, 2, Len(raw) - 2))
            Exit Function
        End If

        ' Single-quoted scalar: a doubled '' represents a single quote.
        If Left$(raw, 1) = "'" And Right$(raw, 1) = "'" Then
            ParseYamlValue = Replace(Mid$(raw, 2, Len(raw) - 2), "''", "'")
            Exit Function
        End If
    End If

    ' Plain (unquoted) scalar.
    ParseYamlValue = raw
End Function

' Serializes the profiles Collection back into YAML text.
Private Function SerializeProfilesYaml(ByVal profiles As Collection) As String
    Dim sb As String
    Dim prof As Object

    sb = "# LLM connection profiles" & vbCrLf & _
         "# This file is generated automatically by the HARNESS_N project." & vbCrLf & _
         "profiles:" & vbCrLf

    For Each prof In profiles
        sb = sb & "  - name: " & QuoteYamlValue(ProfileValue(prof, "name")) & vbCrLf
        sb = sb & "    url: " & QuoteYamlValue(ProfileValue(prof, "url")) & vbCrLf
        sb = sb & "    api_key: " & QuoteYamlValue(ProfileValue(prof, "api_key")) & vbCrLf
        sb = sb & "    model: " & QuoteYamlValue(ProfileValue(prof, "model")) & vbCrLf
        sb = sb & "    max_tokens: " & QuoteYamlValue(ProfileValue(prof, "max_tokens")) & vbCrLf
        sb = sb & "    temperature: " & QuoteYamlValue(ProfileValue(prof, "temperature")) & vbCrLf
        sb = sb & "    system_prompt: " & QuoteYamlValue(ProfileValue(prof, "system_prompt")) & vbCrLf
    Next prof

    SerializeProfilesYaml = sb
End Function

' Quotes and escapes a value as a JSON/YAML double-quoted scalar.
' JSON escaping is a valid subset of YAML double-quoted scalar escaping,
' so the same string can be safely parsed back by ParseYamlValue.
Private Function QuoteYamlValue(ByVal s As String) As String
    QuoteYamlValue = JsonString(s)
End Function

' Safely reads a value from a profile Dictionary and returns it as a string.
Private Function ProfileValue(ByVal prof As Object, ByVal key As String) As String
    If prof Is Nothing Then Exit Function
    If Not prof.Exists(key) Then Exit Function
    If IsNull(prof(key)) Then Exit Function
    ProfileValue = CStr(prof(key))
End Function

' Returns the 1-based index of the profile with the given name
' (case-insensitive comparison), or 0 when the profile is not found.
Private Function FindProfileIndex(ByVal profiles As Collection, ByVal profileName As String) As Long
    Dim i As Long
    Dim prof As Object

    For i = 1 To profiles.Count
        Set prof = profiles(i)
        If StrComp(ProfileValue(prof, "name"), profileName, vbTextCompare) = 0 Then
            FindProfileIndex = i
            Exit Function
        End If
    Next i
End Function


