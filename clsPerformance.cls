VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsPerformance
' Author : Dmitry Gorelenkov
' Date   : 22.08.2014
' Purpose: mesure performance of different code pieces
' Version: 0.9
' Remarks: not adapted for very short code execution
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/clsPerformance.cls</file>
'  <license>VBALib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Explicit

Private colTimers As Collection

'---------------------------------------------------------------------------------------
' Function: Perf
'---------------------------------------------------------------------------------------
' Purpose: helper, if called first time, used START, if called second time, used END
' Param  : String sStringMark mark for performance measurement
' Returns: Double time now, or time passed
' Remarks: TODO output abilities
'---------------------------------------------------------------------------------------
Function Perf(Optional sStringMark As String) As Double
    On Error Resume Next
    Call colTimers.Item(sStringMark)
    'if already exists
    If Err.Number = 0 Then
        Perf = PerfEnd(sStringMark)
    'if not exists, first call, start
    Else
        Perf = PerfStart(sStringMark)
    End If
End Function

'---------------------------------------------------------------------------------------
' Function: PerfStart
'---------------------------------------------------------------------------------------
' Purpose: start performance measurement for the sStringMark
' Param  : String sStringMark mark for performance measurement
' Returns: Double time now
'---------------------------------------------------------------------------------------
Function PerfStart(Optional sStringMark As String) As Double
    On Error Resume Next
    Call colTimers.Remove(sStringMark)
    Call colTimers.Add(Timer, sStringMark)
    Err.Clear
    PerfStart = Timer
End Function

'---------------------------------------------------------------------------------------
' Function: PerfEnd
'---------------------------------------------------------------------------------------
' Purpose: end peformance measurement for the sStringMark
' Param  : String sStringMark mark for performance measurement
' Returns: Double time passed
'---------------------------------------------------------------------------------------
Function PerfEnd(Optional sStringMark As String) As Double
    Dim dblTime As Double
    dblTime = (Timer - CDbl(colTimers(sStringMark)))
    
    Debug.Print "Ausführen: " & sStringMark & " -  " & dblTime & " sekunden"
    
    Call colTimers.Remove(sStringMark)
    
    PerfEnd = dblTime
End Function


Private Sub Class_Initialize()
    Set colTimers = New Collection
End Sub

Private Sub Class_Terminate()
    If colTimers.Count <> 0 Then
        Debug.Print "There are " & colTimers.Count & " timers left. -> bad usage of performance class"
    End If
    
    Set colTimers = Nothing
End Sub
