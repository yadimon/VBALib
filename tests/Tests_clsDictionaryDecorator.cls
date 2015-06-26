VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tests_clsDictionaryDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Tests_clsDictionaryDecorator
' Author : Dmitry Gorelenkov
' Date   : 14.08.2014
' Purpose: Tests for clsDictionaryDecorator.cls
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/tests/Tests_clsDictionaryDecorator.cls</file>
'  <license>PruefDBApp/license.bas</license>
'  <use>PruefDBApp/clsDictionaryDecorator.cls</use>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'AccUnit:TestClass
Private SomeDictionary As Scripting.Dictionary
Private ThisDictDecor As clsDictionaryDecorator
Private i As Long

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
    Set ThisDictDecor = New clsDictionaryDecorator
    Set SomeDictionary = New Scripting.Dictionary
    
    With SomeDictionary
        Call .Add("key1", "item1")
        Call .Add("key2", 2)
        Call .Add(3, "item3")
        Call .Add("key4Obj", New Scripting.Dictionary)
        Call .Add(New Scripting.Dictionary, "valie5FromObjKey")
    End With
End Sub

Public Sub Teardown()
    Set SomeDictionary = Nothing
    Set ThisDictDecor = Nothing
End Sub


Public Sub GetKeyByItem_SimpleTest()
    Set ThisDictDecor.UsedDictionary = SomeDictionary
    
    For i = 0 To UBound(SomeDictionary.Keys)
        Assert.That ThisDictDecor.GetKeyByItem(SomeDictionary.Items(i)), Iz.EqualTo(SomeDictionary.Keys(i))
    Next i
End Sub

Public Sub AddItemKeyExists_SimpleTest()

    For i = 0 To UBound(SomeDictionary.Keys)
        Call ThisDictDecor.Add(SomeDictionary.Keys(i), SomeDictionary.Items(i))
    Next i
    
    For i = 0 To UBound(SomeDictionary.Keys)
        Assert.IsTrue ThisDictDecor.Exists(SomeDictionary.Keys(i))
        Assert.IsFalse ThisDictDecor.Exists(SomeDictionary.Items(i))
        Assert.That ThisDictDecor.Item(SomeDictionary.Keys(i)), Iz.EqualTo(SomeDictionary.Items(i))
        Assert.That ThisDictDecor.Items(i), Iz.EqualTo(SomeDictionary.Items(i))
        Assert.That ThisDictDecor.Keys(i), Iz.EqualTo(SomeDictionary.Keys(i))
    Next i
End Sub

Public Sub Remove_SimpleTest()
    Dim vKey As Variant
    Set ThisDictDecor.UsedDictionary = SomeDictionary
    
    For Each vKey In SomeDictionary.Keys
        Assert.IsTrue ThisDictDecor.Exists(vKey)
        Call ThisDictDecor.Remove(vKey)
        Assert.IsFalse ThisDictDecor.Exists(vKey)
    Next vKey
End Sub

Public Sub RemoveAll_SimpleTest()
    Set ThisDictDecor.UsedDictionary = SomeDictionary
    ThisDictDecor.RemoveAll
    
    Assert.That ThisDictDecor.Count, Iz.EqualTo(0)
    Assert.That SomeDictionary.Count, Iz.EqualTo(0)
End Sub

Public Sub UsedDictionary_SimpleTest()
    Set ThisDictDecor.UsedDictionary = SomeDictionary
    
    Assert.That ThisDictDecor.UsedDictionary, Iz.EqualTo(SomeDictionary)
End Sub

Public Sub CompareMode_SimpleTest()
    'BAD TEST
    Dim aModes As Variant

    aModes = Array(Scripting.CompareMethod.BinaryCompare, _
                    Scripting.CompareMethod.DatabaseCompare, _
                    Scripting.CompareMethod.TextCompare)

    
    'can set compare mode only once
    For i = 0 To UBound(aModes)
        Set SomeDictionary = New Scripting.Dictionary
        Set ThisDictDecor.UsedDictionary = SomeDictionary
        ThisDictDecor.CompareMode = aModes(i)
        Assert.That ThisDictDecor.CompareMode, Iz.EqualTo(SomeDictionary.CompareMode)
    Next i
    
End Sub

Public Sub HashVal_SimpleTest()
    For i = 0 To UBound(SomeDictionary.Keys)
        Assert.That ThisDictDecor.HashVal(SomeDictionary.Items(i)), Iz.EqualTo(SomeDictionary.HashVal(SomeDictionary.Items(i)))
        Assert.That ThisDictDecor.HashVal(SomeDictionary.Keys(i)), Iz.EqualTo(SomeDictionary.HashVal(SomeDictionary.Keys(i)))
    Next i
End Sub
