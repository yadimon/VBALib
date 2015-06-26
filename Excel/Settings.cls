VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Settings
' Author : Dmitry Gorelenkov
' Date   : 23.03.2015
' Purpose:
' Version: 0.9
' Remarks: singleton! Attribute VB_PredeclaredId = True
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/Excel/Settings.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Explicit

Private tExl As clsToolsExl
Public m_Data As Scripting.Dictionary
Private WithEvents m_wksSettings As Excel.Worksheet
Attribute m_wksSettings.VB_VarHelpID = -1
Private Const m_cStartErrorNr = vbObjectError + &H111111
Private m_bChanging As Boolean

Public Enum enm_Settings_Errors
    Settings_NO_SUCH_SETTING_ERRNR = m_cStartErrorNr + 1
    Settings_OBJECT_NOT_SUPPORTED_ERRNR = m_cStartErrorNr + 2
End Enum

Public Function SetSetting(SettingName As String, vValue As Variant)
    If IsObject(vValue) Or IsArray(vValue) Then
        Call ThrowError(Settings_OBJECT_NOT_SUPPORTED_ERRNR, "SetSetting", SettingName)
    End If
    
    m_Data(SettingName) = vValue
    
    Call SaveData
End Function

Public Function GetSetting(SettingName As String, Optional vDefault As Variant) As Variant
    'TODO ajax from server?
    If Not m_Data.Exists(SettingName) Then
        If IsMissing(vDefault) Then
            Call ThrowError(Settings_NO_SUCH_SETTING_ERRNR, "GetSetting", SettingName)
        End If
        GetSetting = vDefault
        Exit Function
    End If
    
    GetSetting = m_Data.Item(SettingName)
End Function

Private Sub Class_Initialize()
    Set tExl = New clsToolsExl
    Set m_Data = New Scripting.Dictionary
    
    Set m_wksSettings = GetOrCreateSettingsWorksheet
    
    LoadData
End Sub

Private Sub Class_Terminate()
    Call SaveData
    
    Set tExl = Nothing
    Set m_Data = Nothing
    Set m_wksSettings = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Function: GetOrCreateSettingsWorksheet
'---------------------------------------------------------------------------------------
' Purpose: get or create setting page
' Returns: Excel.Worksheet
' Remarks: TODO: by creation, create range, name it, hide the worksheet
'---------------------------------------------------------------------------------------
Private Function GetOrCreateSettingsWorksheet() As Excel.Worksheet
    Dim wksSettingWks As Worksheet
    Dim sWksName As String: sWksName = "_settings"
    
    If tExl.IsExistsWorksheet(sWksName, ThisWorkbook) Then
        Set wksSettingWks = ThisWorkbook.Worksheets(sWksName)
    Else
        Set wksSettingWks = tExl.AddWorksheet(sWksName, ThisWorkbook)
    End If
    
    Set GetOrCreateSettingsWorksheet = wksSettingWks
End Function


Private Function ThrowError(ErrNr As Long, sMethod As String, ParamArray aParams() As Variant)
    Dim aParamsVar As Variant
    aParamsVar = aParams
    Err.Raise ErrNr, GetErrSource(sMethod), GetErrDescription(ErrNr, aParamsVar)
End Function

Private Function GetErrSource(sMethodName As String) As String
    GetErrSource = TypeName(Me) & "." & sMethodName & "()"
End Function

Private Function GetErrDescription(ErrNumber As enm_Settings_Errors, aAdditional As Variant) As String
    Select Case ErrNumber
        Case Settings_NO_SUCH_SETTING_ERRNR:
            GetErrDescription = "Setting: " & aAdditional(0) & " not found"
        Case Settings_OBJECT_NOT_SUPPORTED_ERRNR:
            GetErrDescription = "Cant save an object/array to " & aAdditional(0) & ", object/array saving not supported"
        Case Else:
            GetErrDescription = ""
    End Select
End Function

'reload data on change
Private Sub m_wksSettings_Change(ByVal Target As Range)
    If Not m_bChanging Then 'dont load while self changing anything
        LoadData
    End If
End Sub

'loads data from worksheet to dictionary
Private Sub LoadData(Optional dictData As Scripting.Dictionary)
    Dim vDataArray As Variant
    
    Dim i As Long
    '2d array
    vDataArray = GetSettingsRange.Value2
    
    If dictData Is Nothing Then
        Set dictData = m_Data
    End If
    
    For i = LBound(vDataArray) To UBound(vDataArray)
        'key = value. Add or replace values
        dictData(vDataArray(i, 1)) = vDataArray(i, 2)
    Next i
    
End Sub

'saves data to worksheet
Private Sub SaveData(Optional dictData As Scripting.Dictionary, Optional rngData As Excel.Range)
    If dictData Is Nothing Then
        Set dictData = m_Data
    End If
    
    If rngData Is Nothing Then
        Set rngData = GetSettingsRange
    End If
    
    m_bChanging = True
    
    Call tExl.FillColumnByArray(rngData.Column + 1, dictData.Items, rngData.Row, , rngData.Parent, False)
    Call tExl.FillColumnByArray(rngData.Column, dictData.Keys, rngData.Row, , rngData.Parent, False)
    
    m_bChanging = False
End Sub

Private Function GetSettingsRange() As Excel.Range
    Dim lLastRow As Long
    lLastRow = tExl.GetLastRow(m_wksSettings, 1)
    Set GetSettingsRange = m_wksSettings.Range("A2:" & "B" & lLastRow)
End Function



