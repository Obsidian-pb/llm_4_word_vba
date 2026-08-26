VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHarnessToolsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsHarnessToolsHandler - Tools registry and dispatcher for the LLM agent
Option Explicit

' ============================================================
'  Harness tools handler class
' ============================================================

Private tools As Collection   ' Collection of clsToolBase instances

' Initialization: add all tools from the list
Public Sub Init()
    Set tools = New Collection
    
    ' Adding new tools
    Dim toolInstance As clsToolBase

    ' Insert text
    Set toolInstance = New clsTool_InsertText
    tools.Add toolInstance, toolInstance.name
    ' Insert comment
    Set toolInstance = New clsTool_InsertComment
    tools.Add toolInstance, toolInstance.name
    ' Read selected text
    Set toolInstance = New clsTool_ReadSelectedText
    tools.Add toolInstance, toolInstance.name
    ' Write text file. Temporary off
'    Set toolInstance = New clsTool_WriteTextFile
'    tools.Add toolInstance, toolInstance.name
    ' Read text file
    Set toolInstance = New clsTool_ReadTextFile
    tools.Add toolInstance, toolInstance.name

End Sub


' Generate JSON for tools[]
Public Function GetToolsJson() As String
    If tools Is Nothing Or tools.Count = 0 Then
        GetToolsJson = "[]"
        Exit Function
    End If
    
    Dim i As Long
    Dim result As String
    result = "["
    
    For i = 1 To tools.Count
        Dim tool As clsToolBase
        Set tool = tools.item(i)
        
        result = result & "{" & _
            """type"":""function""," & _
            """function"":{" & _
                """name"":""" & tool.name & """," & _
                """description"":""" & tool.Description & """," & _
                """parameters"":" & tool.Schema & _
            "}" & _
        "}"
        
        If i < tools.Count Then result = result & ","
    Next i
    
    result = result & "]"
    GetToolsJson = result
End Function

' Execute tool by name
Public Function Execute(ByVal toolName As String, ByVal argsJson As String) As String
    If tools Is Nothing Then
        Execute = "{""error"":""Tool registry is not initialized""}"
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    
    Dim tool As clsToolBase
    Set tool = tools.item(toolName)
    
    Execute = tool.Execute(argsJson)
    Exit Function
    
ErrHandler:
    Execute = "{""error"":""Tool not found: " & toolName & """}"
End Function


