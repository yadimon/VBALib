VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tests_clsMockHLP_CreatedObjFromTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Tests_clsMockHLP_CreatedObjFromTemplate
' Author : Dmitry Gorelenkov
' Date   : 28.07.2014
' Purpose: Object created from clsMockTemplate, for testing the template
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/tests/Tests_clsMockHLP_CreatedObjFromTemplate.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/clsMockCaller.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Implements Tests_clsMockHLP_SimpleObject 'here we make mock object for Tests_SimpleObject

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
    mMockCaller.ClassName = "Tests_clsMockHLP_SimpleObject"
End Sub

Private Sub Class_Terminate()
    Set mMockCaller = Nothing
End Sub

Private Function Tests_clsMockHLP_SimpleObject_GetThisValue(vThisValue As Variant) As Variant
    Call mMockCaller.CallFn("GetThisValue", Tests_clsMockHLP_SimpleObject_GetThisValue, vThisValue)
End Function

Private Sub Tests_clsMockHLP_SimpleObject_SimpleSubSave5(Save5Here As Variant)
    Call mMockCaller.CallVoid("SimpleSubSave5", Save5Here)
End Sub

Private Property Get Tests_clsMockHLP_SimpleObject_SomePropAsString() As String
    Call mMockCaller.CallFn("SomeProp", Tests_clsMockHLP_SimpleObject_SomePropAsString)
End Property
