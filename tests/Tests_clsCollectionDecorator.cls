VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tests_clsCollectionDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Tests_clsCollectionDecorator
' Author : Dmitry Gorelenkov
' Date   : 21.05.2014
' Purpose: test for clsCollectionDecorator.cls
' Version:
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/tests/Tests_clsCollectionDecorator.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/clsCollectionDecorator.cls</use>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'AccUnit:TestClass

'members
Private m_values As Variant 'array with values to test
Private m_not_values As Variant 'arra with values that are not in values array
Private m_ColH As clsCollectionDecorator
Private m_emptyColH As clsCollectionDecorator

'enum
Private Enum enm_Type
    eTypeArray = 0
    eTypeCollection = 1
    eTypeDictionary = 2
End Enum

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Methods
'--------------------------------------------------------------------

'helper, changes arrays
Private Sub ChangeArrays()
    Dim tmpArray As Variant
    tmpArray = m_values
    m_values = m_not_values
    m_not_values = tmpArray
End Sub

'before test starts
Private Sub Prepare(bReverse As Boolean, eType As Integer)
    Dim i As Long
    
    If (bReverse) Then ChangeArrays
    
    Select Case eType
        Case eTypeArray
            Call m_ColH.Init(m_values)
            
            
        Case eTypeCollection
            Dim tmpCol As New Collection
            For i = LBound(m_values) To UBound(m_values)
                Call tmpCol.Add(m_values(i))
            Next i
            
            Call m_ColH.Init(tmpCol)
            
            
        Case eTypeDictionary
            Dim tmpDic As New Scripting.Dictionary
            For i = LBound(m_values) To UBound(m_values)
                Call tmpDic.Add(i, m_values(i))
            Next i
            
            Call m_ColH.Init(tmpDic)
            
            
        Case Else
            Assert.Fail "unknown type?"
    End Select

End Sub

Public Sub Setup()
    Set m_emptyColH = New clsCollectionDecorator
    Set m_ColH = New clsCollectionDecorator
    
    m_values = Array("", "1", 0, m_ColH)
    m_not_values = Array(1, "0", " ", 6, "text", Null, m_emptyColH)
End Sub

'after test ends
Public Sub Teardown()
    m_values = Null
    m_not_values = Null

    Set m_ColH = Nothing
    Set m_emptyColH = Nothing
End Sub



'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row(false, 0)
'AccUnit:Row(false, 1)
'AccUnit:Row(false, 2)
'AccUnit:Row(true, 0)
'AccUnit:Row(true, 1)
'AccUnit:Row(true, 2)
Public Sub InitTest(bReverse As Boolean, iType As Integer)
    Call Prepare(bReverse, iType)
    
    Assert.IsFalse m_emptyColH.IsInitialized
    Assert.IsTrue m_ColH.IsInitialized
End Sub

'AccUnit:Row(false, 0)
'AccUnit:Row(false, 1)
'AccUnit:Row(false, 2)
'AccUnit:Row(true, 0)
'AccUnit:Row(true, 1)
'AccUnit:Row(true, 2)
Public Sub Contains(bReverse As Boolean, iType As Integer)
    Call Prepare(bReverse, iType)
    
    
    Dim i As Long
    
    For i = LBound(m_values) To UBound(m_values)
        Assert.IsTrue m_ColH.Contains(m_values(i))
    Next i
    
    For i = LBound(m_not_values) To UBound(m_not_values)
        Assert.IsFalse m_ColH.Contains(m_not_values(i))
    Next i

End Sub

'AccUnit:Row(false, 0)
'AccUnit:Row(false, 1)
'AccUnit:Row(false, 2)
'AccUnit:Row(true, 0)
'AccUnit:Row(true, 1)
'AccUnit:Row(true, 2)
Public Sub Length_test(bReverse As Boolean, iType As Integer)
    Call Prepare(bReverse, iType)
    
    Assert.IsTrue m_ColH.Length = (UBound(m_values) - LBound(m_values) + 1)
End Sub
