Attribute VB_Name = "m_Constants"
' Module: m_Constants - Global constants (debug mode, registry keys, harness/toolbar names)
Public Const DEBUG_MODE = False
Public Const HARNESS_NAME = "HARNESS"


Public Const CHAT_LLM_FORM_CAPTION = "Word helper"
Public Const TOLLBAR_NAME = "LLM"

' Registry keys (HKCU)
Public Const REG_ROOT As String = "HKEY_CURRENT_USER\Software\LLMRedactorWord\"
Public Const REG_LLM_API_URL As String = REG_ROOT & "LLM_API_URL"
Public Const REG_LLM_MODEL_ID As String = REG_ROOT & "LLM_MODEL_ID"
Public Const REG_LLM_API_KEY As String = REG_ROOT & "LLM_API_KEY"
Public Const REG_LLM_SYSTEM_PROMPT As String = REG_ROOT & "LLM_SYSTEM_PROMPT"

' Default values
Public Const DEF_LLM_API_URL As String = "https://api.aitunnel.ru/v1/chat/completions"
Public Const DEF_LLM_MODEL_ID As String = "deepseek-v4-pro"

