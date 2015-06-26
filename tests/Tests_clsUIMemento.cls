VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tests_clsUIMemento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Tests_clsUIMemento
' Author : Dmitry Gorelenkov
' Date   : 23.01.2015
' Purpose: Test for clsUIMemento
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/tests/Tests_clsUIMemento.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/clsUIMemento.cls</use>
'  <use>VBALib/tests/frmTestHelpForm_UIMemento.frm</use>
'  <use>VBALib/clsTools.cls</use>
'  <use>VBALib/Access/clsToolsAccs.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'


Option Compare Text
Option Explicit

Private m_frmTestForm As Access.Form
Private m_UIMemento As clsUIMemento
Private t As clsTools
Private tAccs As clsToolsAccs
Private Const m_TEMP_FORM_NAME = "frmTempFormForTestingUIMemento"
'delimiters aliases
Private O As String
Private C As String
Private D As String

'AccUnit:TestClass

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
    Set m_UIMemento = New clsUIMemento
    Set t = New clsTools
    Set tAccs = New clsToolsAccs
    O = m_UIMemento.OPEN_ELEM
    C = m_UIMemento.CLOSE_ELEM
    D = m_UIMemento.DELIM
    Set m_frmTestForm = GetForm()
End Sub

Public Sub Teardown()
    
    Call RemoveForm
    Set m_frmTestForm = Nothing
    Set m_UIMemento = Nothing
    Set t = Nothing
    Set tAccs = Nothing
End Sub

'fill opened form with some default values
Private Sub LoadDefaultValues()
    m_frmTestForm.Controls("txtTestTxt1").Value = "text value 1"
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem1")
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem2")
End Sub

Public Sub ReadFromForm_UsualForm_NoErrors()
    LoadDefaultValues
    ' Assert
    Assert.AreSame m_UIMemento, m_UIMemento.ReadFromForm(m_frmTestForm)
End Sub

Public Sub ReadFromForm_SupportedOnlyDisabled_ThrowError()
    LoadDefaultValues
    
    ' Assert
    Assert.Throws enm_UIMemento_Errors.UIMemento_CANNOT_RECOGNIZE_TYPE_ERRNR
    Call m_UIMemento.ReadFromForm(m_frmTestForm, bSupportedOnly:=False)
End Sub

'todo make more
Public Sub ReadFromFormAndGetSavedControlNames_ReadSpecialFields_ContainsOnlySpecialFields()
    Dim aMemntoItems As Variant
    LoadDefaultValues

    ' Assert
    Call m_UIMemento.ReadFromForm(m_frmTestForm, aFieldNames:=Array("lstTestLst1", "txtTestTxt1"))
    aMemntoItems = m_UIMemento.GetSavedControlNames()
    
    Assert.IsTrue t.ArrayContains("lstTestLst1", aMemntoItems)
    Assert.IsTrue t.ArrayContains("txtTestTxt1", aMemntoItems)
End Sub

Public Sub ReadFromFormAndGetSavedValue_ReadSpecialFields_ContainsSpecialFieldValues()
    LoadDefaultValues
    m_frmTestForm.Controls("lstTestLst1").Selected(1) = True 'select only second

    
    ' Assert
    Call m_UIMemento.ReadFromForm(m_frmTestForm, aFieldNames:=Array("lstTestLst1", "txtTestTxt1"))
    
    Assert.AreEqual "text value 1", m_UIMemento.GetSavedValue("txtTestTxt1")
    Assert.AreEqual Array("someItem2"), m_UIMemento.GetSavedValue("lstTestLst1")
End Sub

Public Sub GetSavedValue_ReadWrongControl_ThrowsErroNoSuchControl()
    Call LoadDefaultValues
    Call m_UIMemento.ReadFromForm(m_frmTestForm, aFieldNames:=Array("lstTestLst1", "txtTestTxt1"))
    
    ' Assert
    Assert.Throws enm_UIMemento_Errors.UIMemento_NO_SUCH_CONTROL_SAVED_ERRNR
    Call m_UIMemento.GetSavedValue("noSuchContorl")
End Sub

Public Sub SetFromControl_AddElementsAndSetNew_OnlyNewIsSavedInMemento()
    ' Arrange
    LoadDefaultValues
    ' Act
    Call m_UIMemento.ReadFromForm(m_frmTestForm)
    Call m_UIMemento.SetFromControl(m_frmTestForm.Controls("txtTestTxt1"))
    ' Assert
    Assert.That m_UIMemento.GetSavedControlNames, Iz.EqualTo(Array("txtTestTxt1"))
End Sub

Public Sub AddFromControl_AddNewControl_ContainsItsControlValue()
    ' Arrange
    LoadDefaultValues
    ' Act
    Call m_UIMemento.AddFromControl(m_frmTestForm.Controls("txtTestTxt1"))
    ' Assert
    Assert.That m_UIMemento.GetSavedControlNames, Iz.EqualTo(Array("txtTestTxt1"))
    Assert.That m_UIMemento.GetSavedValue("txtTestTxt1"), Iz.EqualTo("text value 1")
End Sub

Public Sub SetFromString_AddValuesSetValueFromString_OnlySetValueSaved()
    ' Arrange
    LoadDefaultValues
    ' Act
    Call m_UIMemento.ReadFromForm(m_frmTestForm)
    Call m_UIMemento.SetFromString(O & "txtTestTxt1" & D & "TextBox" & D & "newValue" & C)
    ' Assert
    Assert.That m_UIMemento.GetSavedControlNames, Iz.EqualTo(Array("txtTestTxt1"))
    Assert.That m_UIMemento.GetSavedValue("txtTestTxt1"), Iz.EqualTo("newValue")
End Sub

Public Sub AddFromString_AddedFromStrings_MementoContainsAllValuesOfObjects()
    ' Arrange
    ' Act
    Call m_UIMemento.AddFromString(O & "txt1" & D & "TextBox" & D & "value1" & C)
    Call m_UIMemento.AddFromString(O & "lst2" & D & "ListBox" & D _
                                        & "value1" & D _
                                        & "val u""§ e2" & D _
                                        & "v%_al  /ue3" _
                                        & C)
                                        
    Call m_UIMemento.AddFromString(O & "chk1" & D & "CheckBox" & D & "1" & C)
    Call m_UIMemento.AddFromString(O & "optGrp1" & D & "OptionGroup" & D & "3" & C)
    ' Assert
    Assert.That m_UIMemento.GetSavedControlNames, Iz.EqualTo(Array("txt1", "lst2", "chk1", "optGrp1"))
    Assert.That m_UIMemento.GetSavedValue("txt1"), Iz.EqualTo("value1")
    Assert.That m_UIMemento.GetSavedValue("chk1"), Iz.EqualTo(1)
    Assert.That m_UIMemento.GetSavedValue("optGrp1"), Iz.EqualTo(3)
    Assert.That m_UIMemento.GetSavedValue("lst2"), Iz.EqualTo(Array("value1", "val u""§ e2", "v%_al  /ue3"))
End Sub

Public Sub AsString_getStringOfSavedForm_returnRightString()
    ' Arrange
    m_frmTestForm.Controls("txtTestTxt1").Value = "some text by 1"
    m_frmTestForm.Controls("txtTestTxt2").Value = "  some text by 2"
    m_frmTestForm.Controls("chkTestChk1").Value = 1
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem1")
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem2")
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem3")
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem3")
    m_frmTestForm.Controls("lstTestLst1").Selected(0) = False
    m_frmTestForm.Controls("lstTestLst1").Selected(1) = True
    m_frmTestForm.Controls("lstTestLst1").Selected(2) = True
    m_frmTestForm.Controls("lstTestLst1").Selected(3) = True
    
    Call m_UIMemento.ReadFromForm(m_frmTestForm, , Array("txtTestTxt1", "txtTestTxt2", "chkTestChk1", "lstTestLst1"))
    
    
    Dim Actual As String
    Dim Expected: Expected = O & "txtTestTxt1" & D & "TextBox" & D & "some text by 1" & C & _
                    O & "txtTestTxt2" & D & "TextBox" & D & "  some text by 2" & C & _
                    O & "lstTestLst1" & D & "ListBox" & D & "someItem2" & D & "someItem3" & D & "someItem3" & C & _
                    O & "chkTestChk1" & D & "CheckBox" & D & "1" & C
    ' Act
    Actual = m_UIMemento.AsString()
    
    ' Assert
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'bad test...
Public Sub ApplyOnFormAndMore_applyOnForm_ManyFunctionsExpected()
    ' Arrange
    Dim FormAsString As String
    
    Call AddValuesToListBox(Split("someItem1,someItem2,someItem3,someItem3", ","))
    
    Call AddToMemento("some text by 1", "TextBox", "txtTestTxt1")
    Call AddToMemento("  some text by 2", "TextBox", "txtTestTxt2")
    Call AddToMemento(1, "OptionGroup", "grpTestGrp1")
    Call AddToMemento(1, "CheckBox", "chkTestChk1")
    Call AddToMementoAsListBox(Split("someItem2,someItem3,someItem3", ","))
    
    ' Act
    Call m_UIMemento.ApplyOnForm(m_frmTestForm)
    ' Assert
    'Assert.AreEqual FormAsString, m_UIMemento.AsString
    Assert.AreEqual Array("txtTestTxt1", "txtTestTxt2", "grpTestGrp1", "chkTestChk1", "lstTestLst1"), m_UIMemento.GetSavedControlNames
    Assert.AreEqual "some text by 1", m_UIMemento.GetSavedValue("txtTestTxt1")
    Assert.AreEqual "  some text by 2", m_UIMemento.GetSavedValue("txtTestTxt2")
    Assert.AreEqual 1, m_UIMemento.GetSavedValue("grpTestGrp1")
    Assert.AreEqual 1, m_UIMemento.GetSavedValue("chkTestChk1")
    Assert.AreEqual Array("someItem2", "someItem3", "someItem3"), m_UIMemento.GetSavedValue("lstTestLst1")
    
    'test form content
    'list selected
    Call TestIfAllSelected(Array(0, 1, 1, 1), "lstTestLst1")
    
    'text fields filled
    Assert.AreEqual "some text by 1", m_frmTestForm.Controls("txtTestTxt1").Value
    Assert.AreEqual "  some text by 2", m_frmTestForm.Controls("txtTestTxt2").Value
    Assert.AreEqual 1, m_frmTestForm.Controls("grpTestGrp1").Value
    Assert.AreEqual 1, m_frmTestForm.Controls("chkTestChk1").Value
End Sub

Public Sub ReadFromForm_SaveAsString_ReadTheStringToAnotherMemento_CompareMomentos_MustBeEqual()
    Dim Memento2 As New clsUIMemento

    'fill the form somehow, todo Random!
    m_frmTestForm.Controls("txtTestTxt1").Value = "some text by 1"
    m_frmTestForm.Controls("txtTestTxt2").Value = "  some text by 2"
    m_frmTestForm.Controls("chkTestChk1").Value = False
    m_frmTestForm.Controls("grpTestGrp1").Value = 2
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem1")
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem2")
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem3")
    
    Call m_frmTestForm.Controls("lstTestLst1").AddItem("someItem3")
    m_frmTestForm.Controls("lstTestLst1").Selected(1) = True 'second
    m_frmTestForm.Controls("lstTestLst1").Selected(2) = True
    m_frmTestForm.Controls("lstTestLst1").Selected(3) = True
    
    Call m_UIMemento.ReadFromForm(m_frmTestForm)
    Call Memento2.AddFromString(m_UIMemento.AsString)
    
    'must be same strings
    Assert.AreEqual m_UIMemento.AsString, Memento2.AsString, "memento objects are not euqal"
    
    'test if possible to apply
    Call Memento2.ApplyOnForm(m_frmTestForm)
    
    'check after apply again
    Call m_UIMemento.ReadFromForm(m_frmTestForm) ' TODO
    Assert.AreEqual m_UIMemento.AsString, Memento2.AsString, "memento objects are not euqal after re apply on form"
End Sub

'good random values test, but hard to trace!
Public Sub GenerateRandomForm_TestMementoConsistence()
    Dim i As Long
    Dim Memento2 As clsUIMemento
    
    Set m_UIMemento = Nothing
    Set m_frmTestForm = Nothing
    
    Dim frmTempForm As Access.Form
    
    For i = 0 To 200
        Set Memento2 = New clsUIMemento
        Set m_UIMemento = New clsUIMemento
        Set frmTempForm = GetForm()
        Call FillFormRandomly(frmTempForm)
        
        Call m_UIMemento.ReadFromForm(frmTempForm)
        Call Memento2.AddFromString(m_UIMemento.AsString)
        
        Assert.AreEqual m_UIMemento.AsString, Memento2.AsString, "memento objects are not euqal"

        Call Memento2.ApplyOnForm(frmTempForm)
        Call m_UIMemento.ReadFromForm(frmTempForm)
  
        
        Assert.AreEqual m_UIMemento.AsString, Memento2.AsString, "memento objects are not euqal after re apply on form"
        
        Set Memento2 = Nothing
        Set m_UIMemento = Nothing
        DoCmd.Close acForm, frmTempForm.Name
        Set frmTempForm = Nothing
    Next i
End Sub

'AccUnit:Row("0,0,0,0", "1,0,0,1", "1,0,0,1")
'AccUnit:Row("1,0,1,0", "1,0,1,0", "1,0,1,0")
'AccUnit:Row("0", "1", "1")
'AccUnit:Row("", "1,1,1,1", "")
'AccUnit:Row("1,1,0,1", "", "0,0,0,0")
'AccUnit:Row("", "", "")
'AccUnit:Row("", "0", "")
'AccUnit:Row("0,0", "1,1,1,1", "1,1")
'AccUnit:Row("1,1,0", "1", "1,0,0")
'AccUnit:Row("1,1,0", "1,0", "1,0,0")
'AccUnit:Row("1,1,0", "0", "0,0,0")
'AccUnit:Row("1,1,1,1", "0,0,0,0,0,0,0", "0,0,0,0")
'AccUnit:Row("0,0,0,0", "1,1,1,1,1,1,1", "1,1,1,1")
'AccUnit:Row("0,1,0,1,0,1,0,1,0,1", "1,0,1,0,1,0,1,0,1,0", "1,0,1,0,1,0,1,0,1,0")
Public Sub ApplyOnFormList_selectedItems_reselectItems(sInitialState As String, sToChange As String, sExpectedState As String)
    Dim i As Long
    Dim lMaxNeededItems As Long
    Dim aItems As Variant
    aItems = Array()
    Dim aInitialState As Variant
    Dim aToChange As Variant
    Dim aExpectedState As Variant
    aInitialState = Split(sInitialState, ",")
    aToChange = Split(sToChange, ",")
    aExpectedState = Split(sExpectedState, ",")
    
    lMaxNeededItems = t.Max(UBound(aInitialState), UBound(aToChange))
    
    'create array with values
    For i = 0 To lMaxNeededItems
        Call t.ArrayPush(aItems, "ItemValue" & CStr(i))
    Next i
    
    'add all items to form
    Call AddValuesToListBox(aItems)
    
    'select by initial state
    Call SelectValuesByStringMask(aInitialState)
    
    'setup momento
    Dim aItemsForMomento As Variant
    aItemsForMomento = Array()
    For i = LBound(aToChange) To UBound(aToChange)
        If CBool(aToChange(i)) Then
            Call t.ArrayPush(aItemsForMomento, aItems(i))
        End If
    Next i
    Call AddToMementoAsListBox(aItemsForMomento)
    
    'apply on form
    Call m_UIMemento.ApplyOnForm(m_frmTestForm)
    
    'check if selection is as expected after apply
    Call TestIfAllSelected(aExpectedState)
End Sub


Public Sub ApplyOnFormList_dublicateItemsInList_allSelectedIfOneInMomento()
    Call AddValuesToListBox(Split("val1,val2,dub,dub,val3", ","))
    Call AddToMementoAsListBox(Array("dub"))
    Call m_UIMemento.ApplyOnForm(m_frmTestForm)
    Call TestIfAllSelected(Array(0, 0, 1, 1, 0))
End Sub

Public Sub GetSavedControlNames_setSomeControls_mustContainsThem()
    ' Act
    Call AddToMementoAsListBox(Array("oneValue", "TwoValue"), "testList1Name")
    Call AddToMementoAsListBox(Array("oneValue", "TwoValue"), "testList2Name")
    Call AddToMementoAsListBox(Array("oneValue", "TwoValue"), "another list")
    
    Call AddToMemento("some value", "TextBox", "txtbox1")
    Call AddToMemento("some value", "TextBox", "txtbox2")
    Call AddToMemento("1", "OptionGroup", "grpTestGrp1")
    ' Assert
    Assert.AreEqual Array("testList1Name", "testList2Name", "another list", "txtbox1", "txtbox2", "grpTestGrp1"), m_UIMemento.GetSavedControlNames
End Sub




'|-----------------------------------------------------------------------------------------------------------------------------------------------------|
'|
'|                           Private Methods
'|
'|-----------------------------------------------------------------------------------------------------------------------------------------------------|

'helper
Private Function AddValuesToListBox(aValues As Variant, Optional sListBoxName As String = "lstTestLst1")
    Dim i As Variant
    For i = LBound(aValues) To UBound(aValues)
        Call m_frmTestForm.Controls(sListBoxName).AddItem(aValues(i))
    Next i
End Function

'helper, vMask is like Array(1,0,0,1)
Private Function SelectValuesByStringMask(vMask As Variant, Optional sListBoxName As String = "lstTestLst1")
    Dim i As Variant
    For i = LBound(vMask) To UBound(vMask)
        m_frmTestForm.Controls(sListBoxName).Selected(i) = CBool(vMask(i))
    Next i
End Function

Private Function AddToMementoAsListBox(vValues As Variant, Optional sListBoxName As String = "lstTestLst1")
    Call m_UIMemento.AddControl(sListBoxName, "ListBox", vValues)
End Function

Private Function AddToMemento(sValue As Variant, Optional sType As String = "TextBox", Optional sTextBoxName As String = "txtTestLst1")
    Call m_UIMemento.AddControl(sTextBoxName, sType, sValue)
End Function


'---------------------------------------------------------------------------------------
' Function: TestIfAllSelected
'---------------------------------------------------------------------------------------
' Purpose: test if vMask marked items are selected in listbox
' Param  : Variant vMask array like Array(1,0,0,1) 1 - selected, 0 - not selected
'---------------------------------------------------------------------------------------
Private Function TestIfAllSelected(vMask As Variant, Optional sListBoxName As String = "lstTestLst1")
    Dim i As Long
    'check if selection is right now
    For i = LBound(vMask) To UBound(vMask)
        Assert.AreEqual CBool(vMask(i)), m_frmTestForm.Controls(sListBoxName).Selected(i), "wrong selection after momento applied"
    Next i
End Function


Private Function GetForm() As Access.Form
    If Not tAccs.FormExists(m_TEMP_FORM_NAME) Then
        Call CreateTempForm
    End If
    
    If tAccs.IsFormLoaded(m_TEMP_FORM_NAME) Then
        DoCmd.Close acForm, m_TEMP_FORM_NAME, acSaveNo
    End If
    
    DoCmd.OpenForm m_TEMP_FORM_NAME, , , , , acHidden
    Set GetForm = Access.Forms(m_TEMP_FORM_NAME)
End Function

Private Function CreateTempForm()
    Dim frmNewForm As Access.Form
    Dim lstListBox As Access.ListBox
    Dim cboComboBox As Access.ComboBox
    Dim grpOptGrp As Access.OptionGroup
    Dim i As Long
    Dim sOldName
    Set frmNewForm = Access.CreateForm
    sOldName = frmNewForm.Name

    CreateControl(frmNewForm.Name, acTextBox).Name = "txtTestTxt1"
    CreateControl(frmNewForm.Name, acTextBox).Name = "txtTestTxt2"
    
    For i = 1 To 2
        Set lstListBox = CreateControl(frmNewForm.Name, acListBox)
        lstListBox.Name = "lstTestLst" & i
        lstListBox.MultiSelect = True
        lstListBox.RowSourceType = "Value List"
    Next i
    
    CreateControl(frmNewForm.Name, acCheckBox).Name = "chkTestChk1"
    
    Set cboComboBox = CreateControl(frmNewForm.Name, acComboBox)
    cboComboBox.Name = "cboTestCbo1"
    cboComboBox.RowSourceType = "Value List"
    
    'options group
    Set grpOptGrp = CreateControl(frmNewForm.Name, acOptionGroup)
    grpOptGrp.Name = "grpTestGrp1"

    'add 3 options
    Dim optbtnTmp As Access.OptionButton
    For i = 1 To 3
        CreateControl(frmNewForm.Name, acOptionButton, , grpOptGrp.Name).OptionValue = i
    Next i
    
    'frmNewForm.CurrentView = 1
    DoCmd.Close acForm, sOldName, acSaveYes
    DoCmd.Rename m_TEMP_FORM_NAME, acForm, sOldName
End Function

Private Function RemoveForm()
    
    If tAccs.FormExists(m_TEMP_FORM_NAME) Then
        If tAccs.IsFormLoaded(m_TEMP_FORM_NAME) Then
            DoCmd.Close acForm, m_TEMP_FORM_NAME, acSaveNo
        End If
        
        DoCmd.DeleteObject acForm, m_TEMP_FORM_NAME
    End If
End Function

Private Sub FillFormRandomly(frmForm As Access.Form)
    Dim vValues As Variant
    Dim vOptGrpValues As Variant
    Dim ctrlControl As Access.Control
    Const iMaxListBoxElems = 20
    Const iMaxStringLength = 50
    Dim i As Long
    Dim sStringToAdd As String
    
    vValues = Array(Null, 0, -1, 1, False, True)
    

    Dim dblFillRate As Double
    Dim aUsedStrings As Variant
    dblFillRate = Rnd()

    For Each ctrlControl In frmForm.Controls
        If Rnd() < dblFillRate Then
            If TypeName(ctrlControl) = "ListBox" Or TypeName(ctrlControl) = "ComboBox" Then
                aUsedStrings = Array()
                For i = 0 To Rnd() * iMaxListBoxElems
                    sStringToAdd = GetRandomString(1 + (Rnd() * iMaxStringLength - 1)) 'for debug
                    'aUsedStrings needed to add only unique items
                    If Not t.ArrayContains(sStringToAdd, aUsedStrings) Then
                        Call ctrlControl.AddItem("'" & Replace(sStringToAdd, "'", """") & "'") 'quotation to add with semicolon and quotes
                        Call t.ArrayPush(aUsedStrings, sStringToAdd)
                    End If
                Next i
            End If
            
            'make 30% of cases value = null
            Select Case TypeName(ctrlControl)
                Case "ComboBox", "CheckBox", "OptionGroup":
                     If Rnd() <= 0.3 Then
                        ctrlControl.Value = Null
                        GoTo NextControl
                    End If
            End Select
            
            If TypeName(ctrlControl) = "ListBox" Then
                'select random
                For i = 0 To ctrlControl.ListCount - 1
                    ctrlControl.Selected(i) = CBool(Rnd() > 0.5)
                Next i
            End If
            
            If TypeName(ctrlControl) = "ComboBox" Then
                'TODO select 1 value
                GoTo NextControl
            End If

            If TypeName(ctrlControl) = "CheckBox" Then
                ctrlControl.Value = PickRandom(vValues)
                GoTo NextControl
            End If
            
            If TypeName(ctrlControl) = "OptionGroup" Then
                'ctrlControl.Value = -10000 + (Rnd() * 20000)
                ctrlControl.Value = CInt(1 + Rnd() * 3) '1,2 or 3, for current form
                
                GoTo NextControl
            End If
            

            If TypeName(ctrlControl) = "TextBox" Then
                ctrlControl.Value = GetRandomString(Rnd() * iMaxStringLength)
                GoTo NextControl
            End If

        End If
NextControl:
    Next
End Sub


'http://stackoverflow.com/questions/22630264/ms-access-visual-basic-generate-random-string-in-text-field
Private Function GetRandomString(lLength As Long, Optional bSpecials As Boolean = True) As String
    Dim sString As String
    Dim n As Integer
    Dim ch As Integer 'the character
    For n = 1 To lLength
'        Do
'            ch = Rnd() * 127 'This could be more efficient.
'            '48 is '0', 57 is '9', 65 is 'A', 90 is 'Z', 97 is 'a', 122 is 'z'.
'        Loop While ch < 48 Or ch > 57 And ch < 65 Or ch > 90 And ch < 97 Or ch > 122
'        sString = sString & Chr(ch)
        sString = sString & Chr(32 + Rnd() * 127)
    Next
    
    GetRandomString = sString
End Function

Private Function PickRandom(vArray As Variant) As Variant
    PickRandom = vArray((Rnd() * UBound(vArray)))
End Function
