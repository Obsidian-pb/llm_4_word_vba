VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Module: ThisDocument - Document-level event handlers (commented out by default)
Option Explicit



Public Sub ShowChat(control As IRibbonControl)
    LLM_chat_html.Show
End Sub

Public Sub ShowSettings(control As IRibbonControl)
    LLM_config.Show
End Sub
