VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tests_clsMockCaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Tests_clsMockCaller
' Author : Dmitry Gorelenkov
' Date   : 28.07.2014
' Purpose: Tests for clsMockCaller
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/tests/Tests_clsMockCaller.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/clsMockCaller.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private mMockCaller As clsMockCaller

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------
Public Sub Setup()
    Set mMockCaller = New clsMockCaller
    mMockCaller.ClassName = "MockCallerItSelf"
End Sub

Public Sub Teardown()
    Set mMockCaller = Nothing
End Sub


'AccUnit:Row(true)
'AccUnit:Row(false)
Public Sub CallFn_retBool_and_counter_and_setRetValue(bValue As Boolean)
    Call CallFn_something_and_counter_and_setRetValue(bValue)
End Sub

'AccUnit:Row(-55000)
'AccUnit:Row(0)
'AccUnit:Row(12345678)
Public Sub CallFn_retLong_and_counter_and_setRetValue(lValue As Long)
    Call CallFn_something_and_counter_and_setRetValue(lValue)
End Sub

Public Sub CallFn_retArray_and_counter_and_setRetValue()
    Dim aToTest As Variant
    Dim objCaller As Object
    Set objCaller = New clsMockCaller
    aToTest = Array("some", 5, "thing", objCaller)
    Call CallFn_something_and_counter_and_setRetValue(aToTest)
    Set objCaller = Nothing
End Sub

Public Sub CallFn_retObj_and_counter_and_setRetValue()
    Dim aToTest As Variant
    Set aToTest = New clsMockCaller
    Call CallFn_something_and_counter_and_setRetValue(aToTest)
End Sub

Private Sub CallFn_something_and_counter_and_setRetValue(ByRef vValue As Variant)
    Dim aTimesToTest As Variant
    Dim lLength As Variant 'long
    Dim i As Long
    Dim sTestName As String
    aTimesToTest = Array(0, 1, 10) 'how much times to test
    
    For Each lLength In aTimesToTest
        Call mMockCaller.SetRetValue(sTestName, vValue) 'set value to return
        sTestName = "testFn" & CStr(lLength)
        
        For i = 0 To (CLng(lLength) - 1) 'call lLength times
            Call mMockCaller.CallFn(sTestName, vValue, 1, 2, 3, 4, 5, "doesnt matter")
            Assert.That vValue, Iz.EqualTo(vValue)
        Next i
        
        Assert.That mMockCaller.CalledTimes(sTestName), Iz.EqualTo(CLng(lLength))
    Next lLength
End Sub


Public Sub CallVoid_test()
    Dim vDontChange As Variant
    Set vDontChange = Nothing
    Call mMockCaller.CallVoid("someFnVoid", vDontChange)
    Assert.That vDontChange, Iz.EqualTo(Nothing)
    Assert.That mMockCaller.CalledTimes("someFnVoid"), Iz.EqualTo(1)
End Sub
