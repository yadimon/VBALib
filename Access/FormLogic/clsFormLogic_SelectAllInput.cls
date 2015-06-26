VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormLogic_SelectAllInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsFormLogic_SelectAllInput
' Author : Dmitry Gorelenkov
' Date   : 14.01.2015
' Purpose: class that implements "select all by input" logic on checkbox and input text fields
' Version: 0.9
' Remarks: max 3 fields. TODO class holder? for N fields?
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/Access/FormLogic/clsFormLogic_SelectAllInput.cls</file>
'  <license>PruefDBApp/license.bas</license>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private WithEvents m_chkBox As Access.CheckBox
Attribute m_chkBox.VB_VarHelpID = -1
Private WithEvents m_txtTextBox1 As Access.TextBox
Attribute m_txtTextBox1.VB_VarHelpID = -1
Private WithEvents m_txtTextBox2 As Access.TextBox
Attribute m_txtTextBox2.VB_VarHelpID = -1
Private WithEvents m_txtTextBox3 As Access.TextBox
Attribute m_txtTextBox3.VB_VarHelpID = -1
Private m_DictTxts As Scripting.Dictionary

Private m_DoingEvent As Boolean

Public Function Init(ByRef chkCheckBox As Access.CheckBox, _
                    Optional ByRef txtTextBox1 As Access.TextBox, _
                    Optional ByRef txtTextBox2 As Access.TextBox, _
                    Optional ByRef txtTextBox3 As Access.TextBox)
                    
    If chkCheckBox Is Nothing Then Err.Raise vbObjectError, "clsFormLogic_SelectAllInput.Init", "checkbox is not defined"
    
    Set CheckBox = chkCheckBox
    Set m_DictTxts = New Scripting.Dictionary

    
    Set TextBox1 = txtTextBox1
    Set TextBox2 = txtTextBox2
    Set TextBox3 = txtTextBox3
    
    m_DoingEvent = False
End Function

Public Function TriggerCheckBoxUpdate()
    Call m_chkBox_AfterUpdate
End Function

Public Function TriggerTextBoxsUpdate()
    Call UpdateCheckBox
End Function


'------- CheckBox Property ---------------
Public Property Get CheckBox() As Access.CheckBox
    Set CheckBox = m_chkBox
End Property

Public Property Set CheckBox(ByRef chkCheckBox As Access.CheckBox)
    Set m_chkBox = chkCheckBox
    m_chkBox.AfterUpdate = "[Event Procedure]"
End Property

'------- TextBox 1 Property ---------------
Public Property Get TextBox1() As Access.TextBox
    Set TextBox1 = m_txtTextBox1
End Property

Public Property Set TextBox1(ByRef txtTextBox As Access.TextBox)
     Set m_txtTextBox1 = AddOrReplaceTxt(m_txtTextBox1, txtTextBox)
End Property

'------- TextBox 2 Property ---------------
Public Property Get TextBox2() As Access.TextBox
    Set TextBox2 = m_txtTextBox2
End Property

Public Property Set TextBox2(ByRef txtTextBox As Access.TextBox)
    Set m_txtTextBox2 = AddOrReplaceTxt(m_txtTextBox2, txtTextBox)
End Property

'------- TextBox 3 Property ---------------
Public Property Get TextBox3() As Access.TextBox
    Set TextBox3 = m_txtTextBox3
End Property

Public Property Set TextBox3(ByRef txtTextBox As Access.TextBox)
    Set m_txtTextBox2 = AddOrReplaceTxt(m_txtTextBox3, txtTextBox)
End Property

'if new txtbox, will be added in dict. if replace old, old will be replaced in dict.
Private Function AddOrReplaceTxt(ByRef m_txt As Access.TextBox, ByRef newTxt As Access.TextBox) As Access.TextBox
    If m_DictTxts.Exists(m_txt) Then
        m_DictTxts.Remove (m_txt)
    End If
    
    'if some new value exists, add to dict.
    If Not newTxt Is Nothing Then
        newTxt.AfterUpdate = "[Event Procedure]"
        Call m_DictTxts.Add(newTxt, newTxt)
    End If
    
    Set AddOrReplaceTxt = newTxt
End Function


'------- Events ---------------------
Private Sub m_chkBox_AfterUpdate()
    If m_DoingEvent Then
        Exit Sub
    End If
    m_DoingEvent = True
    Dim txtBox As Variant
    'if checked, empty other fields
    If m_chkBox.Value Then
        'only existing textboxes
        For Each txtBox In m_DictTxts
            txtBox.Value = vbNullString
        Next txtBox
    End If
    
    
    m_DoingEvent = False
End Sub

Private Sub m_txtTextBox1_AfterUpdate()
    If m_DoingEvent Then
        Exit Sub
    End If
    m_DoingEvent = True
    
    Call UpdateCheckBox

    m_DoingEvent = False
End Sub

Private Sub m_txtTextBox2_AfterUpdate()
    If m_DoingEvent Then
        Exit Sub
    End If
    m_DoingEvent = True
    
    Call UpdateCheckBox

    m_DoingEvent = False
End Sub

Private Sub m_txtTextBox3_AfterUpdate()
    If m_DoingEvent Then
        Exit Sub
    End If
    m_DoingEvent = True
    
    Call UpdateCheckBox

    m_DoingEvent = False
End Sub

Private Sub UpdateCheckBox()
    Dim bAnyValueSet As Boolean
    Dim txtBox As Variant
    bAnyValueSet = False
    
    For Each txtBox In m_DictTxts
        If Len(Nz(txtBox.Value, vbNullString)) > 0 Then
            bAnyValueSet = True
            Exit For
        End If
    Next txtBox
    
    'if anything is set, checkbox = false, otherwiese true
    m_chkBox.Value = Not bAnyValueSet
    
End Sub
