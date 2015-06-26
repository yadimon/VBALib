VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tests_clsMockHLP_SimpleObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Tests_clsMockHLP_SimpleObject
' Author : Dmitry Gorelenkov
' Date   : 28.07.2014
' Purpose: simple Object to use in tests
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/tests/Tests_clsMockHLP_SimpleObject.cls</file>
'  <license>VBALib/license.bas</license>
'  <use></use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Property Get SomePropAsString() As String
    SomePropAsString = "property"
End Property

Public Function GetThisValue(vThisValue As Variant) As Variant
    If IsObject(GetThisValue) Then
        Set GetThisValue = vThisValue
    Else
        Set GetThisValue = vThisValue
    End If
End Function

Public Sub SimpleSubSave5(ByRef Save5Here)
    Save5Here = 5
End Sub
