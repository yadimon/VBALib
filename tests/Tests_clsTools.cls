VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tests_clsTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Tests_clsTools
' Author : Dmitry Gorelenkov
' Date   : 20.06.2014
' Purpose: tests for clsTools.cls
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/tests/Tests_clsTools.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/clsTools.cls</use>
'  <ref><name>ADODB</name><major>6</major><minor>1</minor><guid>{B691E011-1797-432E-907A-4D8C69339129}</guid></ref>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'AccUnit:TestClass
Private t As clsTools
Private aTestArrayFrom0 As Variant
Private aTestArrayFrom4to7 As Variant
Private aTestArrays As Variant
Private aTestArrayOtherValues As Variant
Private sPathToTestTempFolder As String
Private FSO As Scripting.FileSystemObject

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
    Set t = New clsTools
    Set FSO = New Scripting.FileSystemObject
    
    sPathToTestTempFolder = t.TempPath & "tempFolderForAccesUnitTests" ' TempPath should work..
    RemoveTempDir
    
    
    aTestArrayFrom0 = Array(1, "2", "abBT_ abt &/)GoreLenKoV%(""67%H", 1241435, t)
    
    ReDim aTestArrayFrom4to7(4 To 7)
    aTestArrayFrom4to7(4) = "4"
    aTestArrayFrom4to7(5) = 5
    Set aTestArrayFrom4to7(6) = t
    aTestArrayFrom4to7(7) = "böl Dmitry a aa ""%&!%$/(§"
    
    aTestArrayOtherValues = Array("xxxxxBadValuexxxxx", 1000090, 46515616516156#, aTestArrayFrom4to7)
    
    aTestArrays = Array(aTestArrayFrom0, aTestArrayFrom4to7)
    
End Sub

'after test ends
Public Sub Teardown()
    On Error Resume Next
    RemoveTempDir
    Set FSO = Nothing
    Set t = Nothing
End Sub

Private Sub RemoveTempDir()
    If FSO.FolderExists(sPathToTestTempFolder) Then
        Call FSO.DeleteFolder(sPathToTestTempFolder, True)
    End If
End Sub


'AccUnit:Row("c:\somefolder\this.file.does.not.exists")
'AccUnit:Row("c:\somefolder\")
'AccUnit:Row("c:\")
'AccUnit:Row("C:\")
'AccUnit:Row("C:\Windows")
'AccUnit:Row("C:\Windows\")
Public Sub FileExists_false(sPath As String)
    Call Assert.IsFalse(t.FileExists(sPath))
End Sub

Public Sub FileExists_true()
    Dim sPath As String

    sPath = Application.CurrentProject.FullName
    
    Call Assert.IsTrue(t.FileExists(sPath))
End Sub

Public Sub FilesList_noFilesOrBadPath_ReturnEmptyArray()
    'create empty folder
    On Error GoTo Err
    Dim sEmptyDirPath As String
    Dim aFilesList As Variant
    sEmptyDirPath = CurrentProject.Path & "\NotR3al3mpTyF0lder\"

    Call MkDir(sEmptyDirPath)
    
    aFilesList = t.FilesList(sEmptyDirPath)
    Assert.That aFilesList, Iz.Empty
    
    aFilesList = t.FilesList(sEmptyDirPath & "notReal")
    Assert.That aFilesList, Iz.Empty
    
    Call RmDir(sEmptyDirPath)
    Exit Sub

Err:
    Call RmDir(sEmptyDirPath)
    Call Assert.Fail("error: " & Err.Description)
    
End Sub

Public Sub FileList_realFiles_shouldReturnArrayWithIt()
    On Error GoTo exit_test
    Dim sEmptyDirPath As String
    Dim aFiles As Variant
    Dim sFile As Variant 'String
    Dim lFNum As Long
    aFiles = Array("someF1le", "some file", "some.txt", "o m g _ f i l e. x. . t x t")
    sEmptyDirPath = CurrentProject.Path & "\NotR3al3mpTyF0lder\"
    
    Call MkDir(sEmptyDirPath)
    
    'create files
    For Each sFile In aFiles
        lFNum = FreeFile
        Open sEmptyDirPath & CStr(sFile) For Output As lFNum
        Close lFNum
    Next sFile
    
    
    Assert.That t.FilesList(sEmptyDirPath), Iz.EqualTo(aFiles)
    
exit_test:
    'delete files
    For Each sFile In aFiles
        Kill sEmptyDirPath & CStr(sFile)
    Next sFile
    'delete dir
    Call RmDir(sEmptyDirPath)
    
    If Err.Number <> 0 Then Call Assert.Fail("error: " & Err.Description)
End Sub



'AccUnit:Row("C:\Windows\", "C:\Windows\")
'AccUnit:Row("C:\Windows", "C:\")
'AccUnit:Row("C:\", "C:\")
'AccUnit:Row("C:", "")
'AccUnit:Row("C:\folder\file.txt", "C:\folder\")
'AccUnit:Row("C:\fo ld er\.txt", "C:\fo ld er\")
'AccUnit:Row("any\relative\", "any\relative\")
'AccUnit:Row("any\relative\file.", "any\relative\")
Public Sub PathGetDir_test(sInput As String, sOutput As String)
    Call Assert.AreEqual(sOutput, t.PathGetDir(sInput))
End Sub

'AccUnit:Row("C:\Windows\", "")
'AccUnit:Row("C:\Windows", "Windows")
'AccUnit:Row("C:\", "")
'AccUnit:Row("C:", "C:")
'AccUnit:Row("C:\folder\file.txt", "file.txt")
'AccUnit:Row("C:\fo ld er\.txt", ".txt")
'AccUnit:Row("any\relative\", "")
'AccUnit:Row("any\relative\file.", "file.")
'AccUnit:Row("any\relative\f il e...", "f il e...")
Public Sub PathGetFileName_test(sInput As String, sOutput As String)
    Call Assert.AreEqual(sOutput, t.PathGetFileName(sInput))
End Sub

Public Sub TempPath_test()
    Assert.IsTrue t.DirExists(t.TempPath)
End Sub

Public Sub FileGetContent_test()
    Const sTestStringUTF As String = "testÜÜÜÄßß?+§%'//&%§("
    Dim TempFile As String
    Dim Actual As Variant
    TempFile = t.TempFileName
    
    'http://stackoverflow.com/a/2525196/2519073
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2 'Specify stream type - we want To save text/string data.
    fsT.Charset = "utf-8" 'Specify charset For the source text data.
    fsT.Open 'Open the stream And write binary data To the object
    fsT.WriteText sTestStringUTF
    fsT.SaveToFile TempFile, 2 'Save binary data To disk

    
    ' Act
    Actual = t.FileGetContent(TempFile)
    
    'delete created file
    Kill TempFile
    
    ' Assert
    Assert.That Actual, Iz.EqualTo(sTestStringUTF)
    
    
End Sub

Public Sub ArrayToCollection_test()
    Dim resultColection As Collection
    Dim i As Long
    Dim arrayToTest As Variant
    
    For Each arrayToTest In aTestArrays
        Set resultColection = t.ArrayToCollection(arrayToTest)
        
        For i = LBound(arrayToTest) To UBound(arrayToTest)
            Assert.That resultColection.Item(CStr(i)), Iz.EqualTo(arrayToTest(i))
        Next i
        
    Next arrayToTest
End Sub

Public Sub ArrayToDictionary_test()
    Dim resultDict As Scripting.Dictionary
    Dim i As Long
    Dim arrayToTest As Variant
    
    For Each arrayToTest In aTestArrays
        Set resultDict = t.ArrayToDictionary(arrayToTest)
        
        For i = LBound(arrayToTest) To UBound(arrayToTest)
            Assert.That resultDict.Item(i), Iz.EqualTo(arrayToTest(i))
        Next i
        
    Next arrayToTest
End Sub

'AccUnit:Row(" Whaaat!!1 Dmitry GorelenkoV \n just made this code!!", "\d (DmIt.*?y Gor.+?ov)", "Dmitry GorelenkoV")
'AccUnit:Row("Some Match", "(Match)", "Match", 0)
'AccUnit:Row("Some Match §% found", "(o|m)e Match (.{2})\s(.*?)$", "found", 2)
Public Sub RegEx_getStringMatched_test(sText As String, sPattern As String, sExpected As String, Optional lSubMatchNr As Long)
    Dim Actual As Variant
    ' Act
    Actual = t.RegEx_getStringMatched(sPattern, sText, lSubMatchNr)
    ' Assert
    Assert.That Actual, Iz.EqualTo(sExpected)
End Sub

Public Sub arrayXDimTo1Dim_usualArraysConvertation()
    Dim aMultiArray2dim(2, 1) As Variant
    Dim aMultiArray3dim(1, 1, 1) As Variant
    Dim aMultiArray, aMultiArrayHardOne 'variants
    Dim aToTest As Variant 'now tested array
    Dim aResult As Variant '1dim modified array
    Dim i As Long
    
    aMultiArray2dim(0, 0) = 0
    aMultiArray2dim(1, 0) = 1
    aMultiArray2dim(2, 0) = 2
    aMultiArray2dim(0, 1) = 3
    aMultiArray2dim(1, 1) = 4
    aMultiArray2dim(2, 1) = 5
    
    aMultiArray3dim(0, 0, 0) = 0
    aMultiArray3dim(1, 0, 0) = 1
    aMultiArray3dim(0, 1, 0) = 2
    aMultiArray3dim(1, 1, 0) = 3
    aMultiArray3dim(0, 0, 1) = 4
    aMultiArray3dim(1, 0, 1) = 5
    aMultiArray3dim(0, 1, 1) = 6
    aMultiArray3dim(1, 1, 1) = 7

    
    aMultiArray = Array(Array(0, 1), Array(2, 3), Array(4, 5, 6, 7, 8))
    
    aMultiArrayHardOne = _
                Array( _
                    Array(), _
                    Array( _
                        aMultiArray, _
                        Array(9), _
                        Array( _
                            Array(10, 11, 12, 13), _
                            Array(14) _
                            ) _
                        ), _
                    Array(15, 16, Array(17, 18, Array(19, 20), 21, 22), 23, 24, 25) _
                )
    
    
    For Each aToTest In Array(aMultiArray2dim, aMultiArray3dim, aMultiArray, aMultiArrayHardOne)
        aResult = t.ArrayXDimTo1Dim(aToTest)
        For i = 0 To UBound(aResult)
            Assert.AreEqual i, aResult(i)
        Next i
    Next aToTest
    
    Assert.That UBound(t.ArrayXDimTo1Dim(aMultiArrayHardOne)), Iz.EqualTo(25)
End Sub

Public Sub arrayXDimTo1Dim_emptyArray_returnEmpty()
    Dim aEmpty As Variant
    aEmpty = Array()
    
    Assert.That t.ArrayXDimTo1Dim(aEmpty), Iz.EqualTo(Array())
End Sub

Public Sub arrayXDimTo1Dim_wrongTypeOrNotArray_raiseError()
    Dim vVariant As Variant
    Dim objObject As Object
    
    Assert.Throws vbObjectError
    Call t.ArrayXDimTo1Dim(vVariant)
    Call t.ArrayXDimTo1Dim(objObject)
End Sub

Public Sub ArrayPush_SimplePush()
    Dim aEmpty, aOne, aTwo
    aEmpty = Array()
    aOne = Array("a")
    aTwo = Array("a", 2)
    
    
    Assert.That t.ArrayPush(aEmpty, "a"), Iz.EqualTo(aOne)
    Assert.That t.ArrayPush(aOne, 2), Iz.EqualTo(aTwo)
End Sub

Public Sub ArrayPush_AssignValue()
    Dim aTwo, aThree, aFour, aAssigned
    aTwo = Array("a", 2)
    aThree = Array("a", 2, "3")
    aFour = Array("a", 2, "3", 4)
    
    aTwo = t.ArrayPush(aTwo, "3")
    Assert.That aTwo, Iz.EqualTo(aThree)
    
    aAssigned = t.ArrayPush(aThree, 4)
    Assert.That aAssigned, Iz.EqualTo(aFour)
End Sub

Public Sub ArrayPop_tests()
    Dim lLengthb4 As Long
    Dim vLastElement As Variant
    
    Dim aArrayVariant As Variant
    Dim aArray() As Variant
    For Each aArrayVariant In aTestArrays
        aArray = aArrayVariant
        While t.ArrayLength(aArray) > 1
            lLengthb4 = t.ArrayLength(aArray)
            Call t.AssignValue(vLastElement, aArray(UBound(aArray)))
            Assert.That t.ArrayPop(aArray), Iz.EqualTo(vLastElement)
            Assert.That (lLengthb4 - 1), Iz.EqualTo(t.ArrayLength(aArray)), "length must be 1 fewer"
        Wend
    Next aArrayVariant
End Sub

Public Sub ArrayPop_errorCases()
    Dim vEmptyArray() As Variant
    vEmptyArray = Array()
    Assert.IsNull t.ArrayPop(vEmptyArray)
End Sub

Public Sub ArrayContains_realValues_mustContain()
    Dim aTestArray As Variant
    Dim Ele As Variant
    
    For Each aTestArray In aTestArrays
        For Each Ele In aTestArray
            Assert.IsTrue t.ArrayContains(Ele, aTestArray)
        Next Ele
    Next aTestArray
    
End Sub

Public Sub ArrayContains_notRealValues_shouldNotContain()
    Dim aTestArray As Variant
    Dim Ele As Variant
   
    For Each aTestArray In aTestArrays
        For Each Ele In aTestArrayOtherValues
            Assert.IsFalse t.ArrayContains(Ele, aTestArray)
        Next Ele
    Next aTestArray
End Sub

Public Sub ArrayIndexByValue_RealIndexes_MustReturnTheIndexes()
    Dim aTestArray As Variant
    Dim i As Long
    
    For Each aTestArray In aTestArrays
        For i = LBound(aTestArray) To UBound(aTestArray)
            Assert.AreEqual i, t.ArrayIndexByValue(aTestArray(i), aTestArray)
        Next i
    Next aTestArray
End Sub

Public Sub ArrayIndexByValue_EmptyArray_ReturnedMinusOne()
    Dim aTestArray As Variant
    Dim Ele As Variant
   
    For Each aTestArray In aTestArrays
        For Each Ele In aTestArray
            Assert.AreEqual -1, t.ArrayIndexByValue(Ele, Array())
        Next Ele
    Next aTestArray
End Sub

Public Sub ArrayIndexByValue_NotExistsValuesInArray_ReturnedMinusOne()
    Dim aTestArray As Variant
    Dim Ele As Variant
   
    For Each aTestArray In aTestArrays
        For Each Ele In aTestArray
            Assert.AreEqual -1, t.ArrayIndexByValue(Ele, aTestArrayOtherValues)
        Next Ele
        
        'reverse
        For Each Ele In aTestArrayOtherValues
            Assert.AreEqual -1, t.ArrayIndexByValue(Ele, aTestArray)
        Next Ele
    Next aTestArray
End Sub

Public Sub ArrayIndexExists_tests()
    Dim aTestArray As Variant
    Dim i As Long
   
    For Each aTestArray In aTestArrays
        For i = LBound(aTestArray) To UBound(aTestArray)
            Assert.IsTrue t.ArrayIndexExists(i, aTestArray)
        Next i
        
        'some values to test
        Assert.IsFalse t.ArrayIndexExists(LBound(aTestArray) - 1, aTestArray)
        Assert.IsFalse t.ArrayIndexExists(UBound(aTestArray) + 1, aTestArray)
        Assert.IsFalse t.ArrayIndexExists(10000000, aTestArray)
        Assert.IsFalse t.ArrayIndexExists(-1, aTestArray)
    Next aTestArray
    
    
End Sub

Public Sub ArrayAllIndexesByValue_tests()
    Dim aTestArray As Variant
    aTestArray = Array(1, "5", 5, 5, "5", "some", "some ", "Some")

    Assert.That t.ArrayAllIndexesByValue(1, aTestArray), Iz.EqualTo(Array(0))
    Assert.That t.ArrayAllIndexesByValue(5, aTestArray), Iz.EqualTo(Array(2, 3))
    Assert.That t.ArrayAllIndexesByValue("5", aTestArray), Iz.EqualTo(Array(1, 4))
    Assert.That t.ArrayAllIndexesByValue("some", aTestArray), Iz.EqualTo(Array(5))
    Assert.That t.ArrayAllIndexesByValue("5", aTestArray, True), Iz.EqualTo(Array(1, 2, 3, 4))
    Assert.That t.ArrayAllIndexesByValue("some", aTestArray, False, True), Iz.EqualTo(Array(5, 6))
    Assert.That t.ArrayAllIndexesByValue(" sOmE  ", aTestArray, True, True, True), Iz.EqualTo(Array(5, 6, 7))
    Assert.That t.ArrayAllIndexesByValue("sOmE", aTestArray, False, False, True), Iz.EqualTo(Array(5, 7))
    
End Sub

Public Sub CompareValuesIfPossible_tests()
    Dim aValuesNotComparable As Variant
    Dim i As Long
    aValuesNotComparable = Array("string", 10, Array(1, 2, 3), t)
    
    For i = LBound(aValuesNotComparable) To UBound(aValuesNotComparable) - 1
        Assert.AreEqual -1, t.CompareValuesIfPossible(aValuesNotComparable(i), aValuesNotComparable(i + 1))
    Next i
    
    For i = LBound(aValuesNotComparable) To UBound(aValuesNotComparable)
        If i <> 2 Then 'array not comparable
            Assert.AreEqual 0, t.CompareValuesIfPossible(aValuesNotComparable(i), aValuesNotComparable(i))
        End If
        
        'specials
        If i <> 0 Then 'not the string
            Assert.AreEqual -1, t.CompareValuesIfPossible("", aValuesNotComparable(i))
        End If

        If i <> 1 Then 'not the number
            Assert.AreEqual -1, t.CompareValuesIfPossible(0, aValuesNotComparable(i))
        End If
    Next i
    
    'more cases
    Assert.AreEqual -1, t.CompareValuesIfPossible(0, "0")
    Assert.AreEqual -1, t.CompareValuesIfPossible("", 0)
    Assert.AreEqual -1, t.CompareValuesIfPossible(0, "")
    Assert.AreEqual 0, t.CompareValuesIfPossible("", "")
    Assert.AreEqual 0, t.CompareValuesIfPossible(0, 0)
    Assert.AreEqual 1, t.CompareValuesIfPossible(1, 1000043354)
    Assert.AreEqual 1, t.CompareValuesIfPossible(1, 10000433540000#)
    
    Dim t1 As clsTools:: Set t1 = New clsTools
    Dim t2 As clsTools:: Set t2 = New clsTools
    Assert.AreEqual 1, t.CompareValuesIfPossible(t1, t2)
    Set t1 = Nothing
    Set t2 = Nothing
End Sub

'AccUnit:Row("string", 0, 0, 0, "string", 1)
'AccUnit:Row("  string  ", 0, 1, 0, "string", 1)
'AccUnit:Row("  strInG  ", 0, 1, 1, "string", 1)
'AccUnit:Row(0, 1, 0, 0, "0", 1)
'AccUnit:Row(0, 1, 1, 0, "0", 1)
'AccUnit:Row(0, 1, 1, 1, "0", 1)
Public Sub UnifyValueIfPossible_tests(vStartEle As Variant, _
                                      bToString As Boolean, _
                                      bTrim As Boolean, _
                                      bCaseInsensive As Boolean, _
                                      vResultMustBe As Variant, _
                                      bCanModifyExpected As Boolean)
    Dim vResult As Variant
    Dim bCanModify As Boolean
    
    bCanModify = t.UnifyValueIfPossible(vStartEle, vResult, bToString, bTrim, bCaseInsensive)
    
    Assert.AreEqual vResultMustBe, vResult
    Assert.AreEqual bCanModifyExpected, bCanModify
End Sub

Public Sub UnifyValueIfPossible_objects()
    Dim vResult As Variant
    Dim bCanModify As Boolean
    Dim t2 As clsTools:: Set t2 = New clsTools
    bCanModify = t.UnifyValueIfPossible(t, vResult, 1, 1, 1)
    
    Assert.That t2, Iz.Not.EqualTo(vResult)
    Assert.That t, Iz.EqualTo(vResult)
    Assert.IsFalse bCanModify
    
    Set t2 = Nothing
End Sub


Public Sub CSVStringToArrayOfArrays_someTests()
    Dim sCSVString As String
    Dim arrExpected As Variant
    Dim arrExpectedNoHeader As Variant
    Dim aRow As Variant
    Dim aPossibleRowDelimiters  As Variant
    Dim aPossibleFieldDelimiters  As Variant
    Dim FieldDelim As Variant
    Dim RowDelim As Variant
    Dim i As Long
    
    arrExpected = Array(Array("field1", "field2"), _
                        Array("value1", "Value2"), _
                        Array("1", "2"))
                        
'    arrExpectedNoHeader = Array(Array("value1", "Value2"), _
'                                Array("1", "2"))
    arrExpectedNoHeader = t.ArrayGetSubArray(arrExpected, 1, UBound(arrExpected), 0)
    
    aPossibleRowDelimiters = Array(vbCr, vbCrLf, vbNewLine, "ENDLINE")
    aPossibleFieldDelimiters = Array(vbTab, " ", "delim", ";")
    
    For Each RowDelim In aPossibleRowDelimiters
        For Each FieldDelim In aPossibleFieldDelimiters

            sCSVString = ""
            'create csv string
            For i = LBound(arrExpected) To UBound(arrExpected)
                aRow = arrExpected(i)
                sCSVString = sCSVString & Join(aRow, FieldDelim) & RowDelim
            Next i
            
            'assertion full, with extra row delimiter at the end
            Assert.That t.CSVStringToArrayOfArrays(sCSVString, CStr(RowDelim), CStr(FieldDelim)), Iz.EqualTo(arrExpected)
            'assertion without last row delimiter, right text
            Assert.That t.CSVStringToArrayOfArrays(Left$(sCSVString, Len(sCSVString) - Len(RowDelim)), CStr(RowDelim), CStr(FieldDelim)), Iz.EqualTo(arrExpected)
            'without header
            Assert.That t.CSVStringToArrayOfArrays(sCSVString, CStr(RowDelim), CStr(FieldDelim), True), Iz.EqualTo(arrExpectedNoHeader)
        Next FieldDelim
    Next RowDelim
End Sub

Public Sub CSVStringToArrayOfArrays_emptyInput()
    Assert.That t.CSVStringToArrayOfArrays("", vbCr, ";"), Iz.EqualTo(Array())
    Assert.That t.CSVStringToArrayOfArrays(";;", vbCr, ";"), Iz.EqualTo(Array(Array("", "", "")))
    Assert.That t.CSVStringToArrayOfArrays("header1;header2", vbCr, ";", True), Iz.EqualTo(Array())
End Sub

Public Sub ArrayRemoveByElement_tests()
    Dim vEle As Variant
    Dim vTestArray As Variant
    Dim copyArray() As Variant
    Dim lLengthb4 As Long
    Dim aArray() As Variant
    
    For Each vTestArray In aTestArrays
        aArray = vTestArray
        'array to modify, with random order
        copyArray = t.ArrayShuffle(aArray)
        'test remove each element
        For Each vEle In aArray
            lLengthb4 = t.ArrayLength(copyArray)
            Assert.IsTrue t.ArrayRemoveByElement(copyArray, vEle)
            Assert.AreEqual t.ArrayLength(copyArray), (lLengthb4 - 1)
        Next vEle
        'no elements left
        Assert.IsTrue (t.ArrayLength(copyArray) = 0)
    Next vTestArray

End Sub

Public Sub ArrayRemoveByElement_badCases()
    Dim testArray() As Variant
    testArray = Array(1, 2)
    Assert.IsFalse t.ArrayRemoveByElement(testArray, "no value")
    testArray = Array()
    Assert.IsFalse t.ArrayRemoveByElement(testArray, "")
End Sub

Public Sub ArrayRemoveByIndex_tests()
    Dim i As Long
    Dim aArray() As Variant
    'ArrayRemoveByElement tests already covers byIndex tests
    'just some tests more
    aArray = t.ArrayCopy(aTestArrayFrom4to7)
    For i = LBound(aTestArrayFrom4to7) To UBound(aTestArrayFrom4to7)
        Assert.IsTrue t.ArrayRemoveByIndex(aArray, 4)
    Next i
    'no elements left
    Assert.IsTrue (t.ArrayLength(aArray) = 0)
End Sub

Public Sub ArrayRemoveByIndex_badCases()
    Dim aArray() As Variant
    aArray = t.ArrayCopy(aTestArrayFrom4to7)
    Assert.IsFalse t.ArrayRemoveByIndex(aArray, 2)
    Assert.IsFalse t.ArrayRemoveByIndex(aArray, -2)
End Sub


Public Sub AssignValue_someFastDirtyTests()
    Dim Obj As clsTools
    Dim sString As String
    
    Assert.IsTrue t.AssignValue(Obj, t)
    Assert.AreEqual Obj, t
    
    Assert.IsTrue t.AssignValue(sString, "string")
    Assert.AreEqual sString, "string"
    Assert.IsTrue t.AssignValue(sString, "")
    Assert.AreEqual sString, ""
    Assert.IsTrue t.AssignValue(sString, 42)
    Assert.AreEqual sString, "42"
End Sub

Public Sub AssignValue_textToLong_Error()
    Dim lLong As Long
    
    Assert.Throws 13
    Call t.AssignValue(lLong, "text")
End Sub

Public Sub AssignValue_objectToString_Error()
    Dim sString As String
    
    Assert.Throws 13
    Call t.AssignValue(sString, t)
End Sub

Public Sub AssignValue_longToObject_Error()
    Dim Obj As New clsTools
    
    Assert.Throws 438
    Call t.AssignValue(Obj, 42)
End Sub

Public Sub AssignValue_stringToObject_Error()
    Dim Obj As New clsTools
    
    Assert.Throws 438
    Call t.AssignValue(Obj, "42")
End Sub

Public Sub ArrayShuffle_tests()
    'uses ShuffleInPlace
    Const lTimes = 10000
    Const lExpectedProbabilityMaxDiff = 0.05 '5 percent
    Dim i As Long
    Dim lIndexOfOne As Long
    Dim aArrayToShuffle() As Variant
    Dim aResult() As Variant
    Dim aStatisticForOne() As Long
    Dim aArray As Variant
    
    'add one array more to deafults
    Call t.ArrayPush(aTestArrays, Array(1, "something", t))
    
    For Each aArray In aTestArrays
        aArrayToShuffle = aArray
        
        'setup statistic array
        ReDim aStatisticForOne(UBound(aArrayToShuffle))
        
        'count new places of 1
        For i = 1 To lTimes
            aResult = t.ArrayShuffle(aArrayToShuffle)
            lIndexOfOne = t.ArrayIndexByValue(aArrayToShuffle(UBound(aArrayToShuffle)), aResult)
            aStatisticForOne(lIndexOfOne) = aStatisticForOne(lIndexOfOne) + 1
        Next i
        
        'probability test
        For i = LBound(aStatisticForOne) To UBound(aStatisticForOne) - 1
            'difference between values must be less then 5%
            Assert.That (aStatisticForOne(i) - aStatisticForOne(i + 1)) / lTimes, Iz.AtMost(lExpectedProbabilityMaxDiff)
        Next i
    Next aArray
End Sub

'AccUnit:Row("123456789", "1,2,3", "", "456789")
'AccUnit:Row("12345", "12,15,13", "", "345")
'AccUnit:Row("abcd", "a", "b", "bbcd")
'AccUnit:Row("AaBbcD", "A,B", "_", "_a_bcD")
'AccUnit:Row("  A0", " ", "", "A0")
'AccUnit:Row("???__0 ", ".,..,...,a,1,00, 0", "-", "???__0 ")
Public Sub StrRemoveChars_RemovementByOneChar_CharsMustBeRemoved(sSource As String, sToRemove As String, sReplacement As String, sExpected As String)
    Dim sResult As String
    Dim aToRemove As Variant
    aToRemove = Split(sToRemove, ",")
    sResult = t.StrReplaceAll(sSource, aToRemove, sReplacement)

    ' Assert
    Assert.That sResult, Iz.EqualTo(sExpected)
End Sub

'AccUnit:Row("???__0 ", "1,2,3,4", "-,-,-,-", "???__0 ")
'AccUnit:Row("abcd", "d,f,g", "a,a,a", "abca")
'AccUnit:Row(" __ z", "_, ,?", "a,b,c", "baabz")
'AccUnit:Row("", "d,f,g", "a,a,a", "")
'AccUnit:Row("abcd", "b", "a", "aacd")
'AccUnit:Row("ZzZz", "z,Z", "_,Z", "Z_Z_")
'AccUnit:Row("0000", "0,00,000", "1,2,3", "1111")
'AccUnit:Row(" 0", "", "", " 0")
Public Sub StrRemoveChars_RemovementByArray_CharsMustBeRemoved(sSource As String, sToRemove As String, sReplacement As String, sExpected As String)
    Dim sResult As String
    Dim aToRemove As Variant
    Dim aToReplace As Variant
    aToRemove = Split(sToRemove, ",")
    aToReplace = Split(sReplacement, ",")
    sResult = t.StrReplaceAll(sSource, aToRemove, aToReplace)

    ' Assert
    Assert.That sResult, Iz.EqualTo(sExpected)
End Sub

Public Sub StrRemoveChars_NotEqualArrayLegth_ThrowsError()
    Assert.Throws vbObjectError
    
    Call t.StrReplaceAll("abc", Array("a", "b"), Array("z"))
End Sub

Public Sub StrRemoveChars_WrongReplaceArg_ThrowsError()
    Assert.Throws vbObjectError
    
    Call t.StrReplaceAll("abc", Array("a", "b"), t)
End Sub

Public Sub CreateDirectory_CreateByPathWithSubDirs_allSubDirsCreated()
    ' Arrange

    Dim sBigPath As String: sBigPath = sPathToTestTempFolder & "\sub1\_2sub\3 sub folder\noo\"
    
    Assert.IsTrue t.CreateDirectory(sBigPath)

    Assert.IsTrue FSO.FolderExists(sPathToTestTempFolder & "\sub1")
    Assert.IsTrue FSO.FolderExists(sPathToTestTempFolder & "\sub1\_2sub\")
    Assert.IsTrue FSO.FolderExists(sPathToTestTempFolder & "\sub1\_2sub\3 sub folder\")
    Assert.IsTrue FSO.FolderExists(sPathToTestTempFolder & "\sub1\_2sub\3 sub folder\noo\")

    Assert.IsFalse t.CreateDirectory(sBigPath) ' second create dir is buggy? blocks dir?
    
    Assert.IsTrue t.DirExists(sPathToTestTempFolder & "\sub1\_2sub\3 sub folder\noo\")
End Sub
