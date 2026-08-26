VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LLM_config 
   Caption         =   "Настройки ИИ"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5355
   OleObjectBlob   =   "LLM_config.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LLM_config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Form: LLM_config - User form for LLM settings configuration (URL, model, API key, system prompt)
Option Explicit


Private Sub cb_Save_Click()
    ' Save to current user registry
    SaveSettingToRegistry REG_LLM_API_URL, tb_base_url.text
    SaveSettingToRegistry REG_LLM_MODEL_ID, tb_model_main.text
'    SaveSettingToRegistry REG_LLM_MODEL_FAST_ID, tb_model_fast.text
    SaveSettingToRegistry REG_LLM_API_KEY, tb_api_key.text
    SaveSettingToRegistry REG_LLM_SYSTEM_PROMPT, tb_system_prompt.text

    MsgBox "LLM settings saved to registry" & vbCrLf, vbInformation
    
    Me.Hide
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    tb_base_url.text = CStr(GetSettingFromRegistry(REG_LLM_API_URL, DEF_LLM_API_URL))
    tb_model_main.text = CStr(GetSettingFromRegistry(REG_LLM_MODEL_ID, DEF_LLM_MODEL_ID))
'    tb_model_fast.text = CStr(GetSettingFromRegistry(REG_LLM_MODEL_FAST_ID, DEF_LLM_MODEL_FAST_ID))
    tb_api_key.text = CStr(GetSettingFromRegistry(REG_LLM_API_KEY, ""))
    tb_system_prompt.text = CStr(GetSettingFromRegistry(REG_LLM_SYSTEM_PROMPT, ""))
    On Error GoTo 0
End Sub
