VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMockTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsMockTemplate
' Author : Dmitry Gorelenkov
' Date   : 28.07.2014
' Purpose: Template Class for mocked objects
' Version: 0.9
' Remarks: based on VBMockTests from Kelly Ethridge
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/clsMockTemplate.cls</file>
'  <license>PruefDBApp/license.bas</license>
'  <use>VBALib/clsMockCaller.cls</use>
'  <test>VBALib/tests/Tests_clsMockTemplate.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'Implements IMyInterface/MyClass

Private mMockCaller As clsMockCaller

'this function sets return values for any method of the mock object
Public Function MockSetReturnValue(sFnName As String, vValue As Variant)
    Call mMockCaller.SetRetValue(sFnName, vValue)
End Function

Public Function MockMethodCalledTimes(sFnName As String) As Long
    MockMethodCalledTimes = mMockCaller.CalledTimes(sFnName)
End Function

Private Sub Class_Initialize()
    Set mMockCaller = New clsMockCaller
    mMockCaller.ClassName = "IMyInterface/MyClass"
    'default values
    Call MockSetReturnValue("CallSub", "somevalue")
End Sub

Private Sub Class_Terminate()
    Set mMockCaller = Nothing
End Sub

' Samples on how to connect your interface/class calls to the mock object.
#If False Then
Private Sub IMyInterface_CallSub(ByVal Name As String)
    Call mMockCaller.CallVoid("CallSub", Name)
End Sub

Private Function IMyInterface_CallFunction(ByVal Name As String) As String
    Call mMockCaller.CallFn("CallFunction", IMyInterface_CallFunction, Name)
End Function

Private Property Get IMyInterface_GetValue() As Variant
    Call mMockCaller.CallFn("GetValue", IMyInterface_GetValue)
End Property

Private Property Let IMyInterface_LetValue(ByVal RHS As Variant)
    Call mMockCaller.CallVoid("LetValue", RHS)
End Property

Private Property Set IMyInterface_SetValue(ByVal RHS As Variant)
    Call mMockCaller.CallVoid("SetValue", RHS)
End Property
#End If
