Attribute VB_Name = "m_Utils"
' Module: m_Utils - Utility functions (logging, registry read/write)
Option Explicit

' =============================================
' Log output
' =============================================
Public Sub Log(msg)
    If DEBUG_MODE = True Then
        Debug.Print msg
    End If
End Sub



' Universal registry value reader (returns defaultValue if missing)
Public Function GetSettingFromRegistry(ByVal fullKey As String, ByVal defaultValue As String) As String
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

' Universal registry value writer (string, REG_SZ)
Public Sub SaveSettingToRegistry(ByVal fullKey As String, ByVal value As String)
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.RegWrite fullKey, value, "REG_SZ"
End Sub
