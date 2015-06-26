VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMockCaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsMockCaller
' Author : Dmitry Gorelenkov
' Date   : 28.07.2014
' Purpose: Used in methods of mocked objects, to call Fns/Subs/Props
' Version: 0.9
' Remarks: TODO use params array, log it?
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/clsMockCaller.cls</file>
'  <license>PruefDBApp/license.bas</license>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'  <test>VBALib/tests/Tests_clsMockCaller.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_dictFnRetVal As Scripting.Dictionary
Private m_dictCallCounters As Scripting.Dictionary
Public ClassName As String

'---------------------------------------------------------------------------------------
' Function: CallFn
'---------------------------------------------------------------------------------------
' Purpose: call property/function/sub anything with return value
' Param  : String sFnName function/sub/property name
' Param  : Variant ptrReturnValue reference to return value
' Param  : ParamArray aParams() any other parameters 'TODO
'---------------------------------------------------------------------------------------
Public Function CallFn(sFnName As String, ByRef ptrReturnValue As Variant, ParamArray aParams() As Variant)
    Call IncrCallCount(sFnName)
    
    If m_dictFnRetVal.Exists(sFnName) Then
        If IsObject(m_dictFnRetVal.Item(sFnName)) Then
            Set ptrReturnValue = m_dictFnRetVal.Item(sFnName)
        Else
            ptrReturnValue = m_dictFnRetVal.Item(sFnName)
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Function: CallVoid
'---------------------------------------------------------------------------------------
' Purpose: call sub or set property
' Param  : String sFnName sub/property name
' Param  : ParamArray aParams() any other parameters
'---------------------------------------------------------------------------------------
Public Function CallVoid(sFnName As String, ParamArray aParams() As Variant)
    Call IncrCallCount(sFnName)
End Function

'---------------------------------------------------------------------------------------
' Function: SetRetValue
'---------------------------------------------------------------------------------------
' Purpose: set return value for some function/sub/property
' Param  : String sFnName function/sub/property name
' Param  : Variant vValue return value for the method
'---------------------------------------------------------------------------------------
Public Function SetRetValue(sFnName As String, vValue As Variant)
    If m_dictFnRetVal.Exists(sFnName) Then
        If IsObject(vValue) Then
            Set m_dictFnRetVal.Item(sFnName) = vValue
        Else
            m_dictFnRetVal.Item(sFnName) = vValue
        End If
    Else
        Call m_dictFnRetVal.Add(sFnName, vValue)
    End If
End Function

'---------------------------------------------------------------------------------------
' Function: CalledTimes
'---------------------------------------------------------------------------------------
' Purpose: Returns number of times, the sFnName was called
' Param  : String sFnName function/sub/property name
' Returns: Long number of call times
'---------------------------------------------------------------------------------------
Public Function CalledTimes(sFnName As String) As Long
    If m_dictCallCounters.Exists(sFnName) Then
        CalledTimes = m_dictCallCounters.Item(sFnName)
    Else
        CalledTimes = 0
    End If
End Function

'---------------------------------------------------------------------------------------
' Function: IncrCallCount
'---------------------------------------------------------------------------------------
' Purpose: increases calling times by one for the sFnName method
' Param  : String sFnName function/sub/property name
'---------------------------------------------------------------------------------------
Private Function IncrCallCount(sFnName As String)

    If m_dictCallCounters.Exists(sFnName) Then
        m_dictCallCounters.Item(sFnName) = m_dictCallCounters.Item(sFnName) + 1
    Else
        m_dictCallCounters.Item(sFnName) = 1
    End If
    
    'Log "Mock Called: " & ClassName & "." & sFnName & " | " & CalledTimes(sFnName) & " times"
End Function

Private Function Log(sText As String)
    Debug.Print Format(Now(), "hh:mm:ss ") & sText
End Function


Private Sub Class_Initialize()
    Set m_dictFnRetVal = New Scripting.Dictionary
    Set m_dictCallCounters = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set m_dictFnRetVal = Nothing
    Set m_dictCallCounters = Nothing
End Sub
