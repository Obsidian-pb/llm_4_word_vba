Attribute VB_Name = "m_toolbar"
' Module: m_toolbar - Toolbar creation and management (add/remove LLM toolbar)
Option Explicit

Public btns As c_buttons


Sub Autoexec()
    AddTB_LLM
End Sub




Sub AddTB_LLM()
'Procedure for adding the TOLLBAR_NAME toolbar-------------------------------

'---Declare variables and constants--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim i As Integer
    
'---Check if the TOLLBAR_NAME toolbar already exists------------------------------
    For i = 1 To Application.Vbe.CommandBars.Count
        If Application.Vbe.CommandBars(i).name = TOLLBAR_NAME Then Exit Sub
    Next i

'---Create the TOLLBAR_NAME toolbar--------------------------------------------
    Set Bar = Application.Vbe.CommandBars.Add(Position:=msoBarTop, Temporary:=True)
    With Bar
        .name = TOLLBAR_NAME
        .Visible = True
    End With

'---Add buttons
    AddButtons
    
End Sub

Sub RemoveTB_LLM()
'Procedure for removing the TOLLBAR_NAME toolbar -------------------------------
    On Error GoTo ex
    Application.Vbe.CommandBars(TOLLBAR_NAME).Delete
    Set btns = Nothing
ex:
End Sub



Sub AddButtons()
'Procedure for adding a new button to the TOLLBAR_NAME toolbar --------------
    On Error GoTo ex
'---Declare variables and constants--------------------------------------------------
    Dim Bar As CommandBar
    Dim docPath As String

    Set Bar = Application.Vbe.CommandBars(TOLLBAR_NAME)
    
'---Add buttons to the TOLLBAR_NAME toolbar --------------------------------
'---"Chat" button-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .caption = "Chat"
        .Tag = "LLM chat"
        .TooltipText = "Open chat"
        .FaceID = 201
    End With
'---"Settings" button -------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .caption = "Settings"
        .Tag = "LLM settings"
        .TooltipText = "LLM Settings"
        .FaceID = 642
    End With
    
    
'---Activate the button tracking class
    Set btns = New c_buttons
    
Set Bar = Nothing
Exit Sub
ex:
    MsgBox "An error occurred during program execution! If it repeats, contact the developer.", , ThisDocument.name
    Log Err & ": AddButtons LLM"
End Sub
