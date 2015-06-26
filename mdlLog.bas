Attribute VB_Name = "mdlLog"
'---------------------------------------------------------------------------------------
' Class Module: mdlLog (Dmitry Gorelenkov, 23.05.2014)
'---------------------------------------------------------------------------------------
' Purpose: Logger
' Updated: 21.05.2014
' Remarks: TODO real log, with filter, cvs file, limits, settings, private properties
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/mdlLog.bas</file>
'  <license>VBALib/license.bas</license>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
'
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Enum: enm_logger_destination where to save messages
'---------------------------------------------------------------------------------------
' Item: enm_logger_dest_debug (0) debug window
'---------------------------------------------------------------------------------------
Public Enum enm_logger_destination
    enm_logger_dest_debug = 0&
    'enm_logger_dest_file = 1& 'TODO
    'enm_logger_dest_database = 2&
End Enum

'---------------------------------------------------------------------------------------
' Enum: enm_logger_format
'---------------------------------------------------------------------------------------
' Item: enm_logger_format_simple (0) no special format
'---------------------------------------------------------------------------------------
Public Enum enm_logger_format
    enm_logger_format_simple = 0&
    'enm_logger_format_own = 1& 'TODO
    'enm_logger_format_JSON = 2&
    'enm_logger_format_XML = 3&
End Enum

'---------------------------------------------------------------------------------------
' Enum: enm_logger_level
'---------------------------------------------------------------------------------------
' Item: enm_logger_error (10) error
' Item: enm_logger_warning (20) warning
' Item: enm_logger_info (30) info
'---------------------------------------------------------------------------------------
Public Enum enm_logger_level
    [_Default] = 0
    enm_logger_error = 10
    enm_logger_warning = 20
    enm_logger_info = 30
End Enum

Private m_log_format As String
Private m_log_csv_separator As String 'todo
Private m_log_current_level As String
Private m_templHandler As clsTemplateHandler
Private m_dictLoggerLevel As Scripting.Dictionary

Private Const m_DEFAULT_LOG_FORMAT As String = "{DateTime}|{LogLevel}|{Message}|{Category}"
Private Const m_DEFAULT_LOG_CSV_SEP As String = " " 'TODO
Private Const m_DEFAULT_LOG_LVL As Long = enm_logger_level.enm_logger_info

'---------------------------------------------------------------------------------------
' Function: log
'---------------------------------------------------------------------------------------
' Purpose: simple log some message
' Param  : String sMessage message to save
' Param  : enm_logger_level level level of message
' Param  : String Category what category has the message
' Returns: Boolean success
' Remarks: TODO category
'---------------------------------------------------------------------------------------
Public Function Log(Optional sMessage As String, Optional level As enm_logger_level = m_DEFAULT_LOG_LVL, Optional sCategory As String) As Boolean
    Dim sFormattedMsg As String
    On Error GoTo log_Error
    
    If enm_logger_level.[_Default] = level Then
        level = m_DEFAULT_LOG_LVL
    End If
    
    'dont log if higher level
    If level > m_log_current_level Then
        Log = False
        Exit Function
    End If
    'format message
    sFormattedMsg = getFormattedMsg(sMessage, level, sCategory)
    'send message
    sendMsg (sFormattedMsg)
    
    Log = True
    Exit Function

log_Error:

    Call hErr("log of Class Module clsLogger")
    Log = False
    
End Function


'---------------------------------------------------------------------------------------
' Function: getFormattedMsg
'---------------------------------------------------------------------------------------
' Purpose: formatting message
' Param  : String sMessage message to save
' Param  : enm_logger_level level level of message
' Param  : String Category what category has the message
' Returns: String
' Remarks: TODO. add formats.
'---------------------------------------------------------------------------------------
Private Function getFormattedMsg(Optional sMessage As String, Optional level As enm_logger_level = m_DEFAULT_LOG_LVL, Optional sCategory As String) As String
    Dim replaceDict As Scripting.Dictionary
    'create dict with placeholders names (key) and values
    Set replaceDict = createValuesDict(sMessage, level, sCategory)
    
    'replace format by the placeholdes
    getFormattedMsg = applyOnTemplate(m_log_format, replaceDict)
    
    Set replaceDict = Nothing
End Function

'---------------------------------------------------------------------------------------
' Function: applyOnTemplate
'---------------------------------------------------------------------------------------
' Purpose: apply parameters from dictVars to sTemplate
' Param  : String sTemplate template
' Param  : Scripting.Dictionary dictVars variables in dictionary ("replace me(key)", "with me(value)")
' Returns: Variant
' Remarks: TODO values type recogniiton
'---------------------------------------------------------------------------------------
Private Function applyOnTemplate(sTemplate As String, dictVars As Scripting.Dictionary) As String
    Dim sKey As Variant
    
    For Each sKey In dictVars.Keys
        sTemplate = replace$(sTemplate, "{" & sKey & "}", CStr(dictVars.Item(sKey)), , , vbTextCompare)
    Next sKey
    
    applyOnTemplate = sTemplate
End Function

'---------------------------------------------------------------------------------------
' Function: createValuesDict
'---------------------------------------------------------------------------------------
' Purpose: creates dictionary with values for now
' Param  :
' Returns: Scripting.Dictionary
' Remarks:
'---------------------------------------------------------------------------------------
Private Function createValuesDict(sMessage As String, level As enm_logger_level, sCategory As String) As Scripting.Dictionary
    Dim replaceDict As New Scripting.Dictionary
    
    Call replaceDict.add("DateTime", Format$(Now(), "dd.nn.yyyy - hh:mm"))
    Call replaceDict.add("LogLevel", getLevelName(level))
    Call replaceDict.add("Message", sMessage)
    Call replaceDict.add("Category", sCategory)
    
    Set createValuesDict = replaceDict
End Function

Private Function getLevelName(enmLevel As enm_logger_level) As String
    getLevelName = m_dictLoggerLevel.Item(enmLevel)
End Function

Private Function sendMsg(sMsg As String)
    'todo switch options
    Debug.Print sMsg
End Function

Private Sub Class_Initialize()
    m_log_format = m_DEFAULT_LOG_FORMAT
    m_log_csv_separator = m_DEFAULT_LOG_CSV_SEP
    m_log_current_level = m_DEFAULT_LOG_LVL
    
    Set m_dictLoggerLevel = New Scripting.Dictionary
    
    'fill dictionary with values
    Call m_dictLoggerLevel.add(enm_logger_level.enm_logger_error, "error")
    Call m_dictLoggerLevel.add(enm_logger_level.enm_logger_info, "info")
    Call m_dictLoggerLevel.add(enm_logger_level.enm_logger_warning, "warning")
End Sub

Private Sub Class_Terminate()
    Set m_dictLoggerLevel = Nothing
End Sub
