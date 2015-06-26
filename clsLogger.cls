VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsLogger
' Author : Dmitry Gorelenkov
' Date   : 11.09.2014
' Purpose: very simple temp logger. TODO create/use real one
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/clsLogger.cls</file>
'  <license>PruefDBApp/license.bas</license>
'  <use></use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Function Info(sMsg As String) As Boolean
    Info = Log("Info: " & sMsg)
End Function

Public Function Warn(sMsg As String) As Boolean
    Warn = Log("Warn: " & sMsg)
End Function

Public Function Error(sMsg As String) As Boolean
    Error = Log("Error: " & sMsg)
End Function

Public Function Trace(sMsg As String) As Boolean
    Trace = Log("Trace: " & sMsg)
End Function

Public Function Fatal(sMsg As String) As Boolean
    Fatal = Log("Fatal: " & sMsg)
End Function


Private Function Log(sMsg) As Boolean
    Debug.Print sMsg
End Function
