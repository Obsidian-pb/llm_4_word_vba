VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LLM_config 
   Caption         =   "LLM Settings"
   ClientHeight    =   9405.001
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
'    SaveSettingToRegistry REG_LLM_API_URL, tb_base_url.text
'    SaveSettingToRegistry REG_LLM_MODEL_ID, tb_model_main.text
'    SaveSettingToRegistry REG_LLM_MODEL_FAST_ID, tb_model_fast.text
'    SaveSettingToRegistry REG_LLM_API_KEY, tb_api_key.text
'    SaveSettingToRegistry REG_LLM_SYSTEM_PROMPT, tb_system_prompt.text
'    MsgBox "LLM settings saved to registry" & vbCrLf, vbInformation

    SaveLLMProfile tb_profile_name.text, _
        Me.tb_base_url.text, _
        Me.tb_api_key.text, _
        Me.tb_model_main.text, _
        Me.tb_max_tokens.text, _
        Me.tb_temperature.text, _
        Me.tb_system_prompt.text


    MsgBox "LLM settings saved" & vbCrLf, vbInformation
    
    refresh_profiles_list
    
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    
    ' Profiles:
    refresh_profiles_list
End Sub

Private Sub refresh_profiles_list()
    ' Profiles:
    Dim names As Collection
    Dim i As Long

    Set names = GetProfileNames()
    Me.cbox_profile.Clear
    For i = 1 To names.Count
        Me.cbox_profile.AddItem names(i)
    Next i
End Sub

Private Sub cbox_profile_Change()
    Dim u As String, k As String, m As String, mt As String, t As String, s As String
    If GetLLMProfileParams(Me.cbox_profile.value, u, k, m, mt, t, s) Then
        Me.tb_profile_name.text = Me.cbox_profile.value
        Me.tb_base_url.text = u
        Me.tb_api_key.text = k
        Me.tb_model_main.text = m
        Me.tb_max_tokens.text = mt
        Me.tb_temperature.text = t
        Me.tb_system_prompt.text = s
    End If
End Sub


