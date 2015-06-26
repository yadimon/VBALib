VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tests_clsMockTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Tests_clsMockTemplate
' Author : Dmitry Gorelenkov
' Date   : 28.07.2014
' Purpose: test mock template class
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/tests/Tests_clsMockTemplate.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/tests/Tests_clsMockHLP_CreatedObjFromTemplate.cls</use>
'  <use>VBALib/tests/Tests_clsMockHLP_SimpleObject.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private objMock As Tests_clsMockHLP_CreatedObjFromTemplate
Private objSimpleObjectMocked As Tests_clsMockHLP_SimpleObject

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
    Set objMock = New Tests_clsMockHLP_CreatedObjFromTemplate
    Set objSimpleObjectMocked = objMock
End Sub

Public Sub Teardown()
    Set objMock = Nothing
    Set objSimpleObjectMocked = Nothing
End Sub

Public Sub Test_Mocked_Property()
    Const Expected = "some striiiiing"
    
    Call objMock.MockSetReturnValue("SomeProp", Expected)
    
    ' Assert
    Assert.That objSimpleObjectMocked.SomePropAsString, Iz.EqualTo(Expected)
    Assert.That objMock.MockMethodCalledTimes("SomeProp"), Iz.EqualTo(1)
End Sub

Public Sub Test_Mocked_Function()
    Const Expected = 1337
    Call objMock.MockSetReturnValue("GetThisValue", Expected)
    Assert.That objSimpleObjectMocked.GetThisValue(42), Iz.EqualTo(Expected)
    Assert.That objMock.MockMethodCalledTimes("GetThisValue"), Iz.EqualTo(1)
End Sub

Public Sub Test_Mocked_Sub()
    Dim lHereWillBeSaved5 As Long
    objSimpleObjectMocked.SimpleSubSave5 lHereWillBeSaved5
    Assert.That objMock.MockMethodCalledTimes("SimpleSubSave5"), Iz.EqualTo(1)
End Sub
