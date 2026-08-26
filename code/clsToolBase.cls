VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToolBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsToolBase - Abstract base class for all agent tools (Name, Description, Schema, Execute)
' === CLASS: clsToolBase ===
' Base class for all agent tools.
' Each specific tool inherits from it and overrides
' the Name, Description, Schema properties and the Execute method.
Option Explicit

' Tool name (what LLM sees in tools[])
Public Property Get name() As String
    Err.Raise 5, , "The Name method must be overridden in the derived class"
End Property

' Description for LLM
Public Property Get Description() As String
    Err.Raise 5, , "The Description method must be overridden in the derived class"
End Property

' JSON Schema of parameters
Public Property Get Schema() As String
    Err.Raise 5, , "The Schema method must be overridden in the derived class"
End Property

' Execute the tool
Public Function Execute(ByVal argsJson As String) As String
    Err.Raise 5, , "The Execute method must be overridden in the derived class"
End Function
