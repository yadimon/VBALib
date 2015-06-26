VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormLogic_SelectAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsFormLogic_SelectAll
' Author : Dmitry Gorelenkov
' Date   : 14.01.2015
' Purpose: class that implements "select all" logic on checkbox and list control
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/Access/FormLogic/clsFormLogic_SelectAll.cls</file>
'  <license>PruefDBApp/license.bas</license>
'  <use></use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private WithEvents m_chkBox As Access.CheckBox
Attribute m_chkBox.VB_VarHelpID = -1
Private WithEvents m_lstList As Access.ListBox
Attribute m_lstList.VB_VarHelpID = -1
Private m_DoingEvent As Boolean

Public Function Init(ByRef chkCheckBox As Access.CheckBox, ByRef lstList As Access.ListBox)
    Set CheckBox = chkCheckBox
    Set ListBox = lstList
    
    m_DoingEvent = False
End Function

Public Function TriggerListUpdate()
    Call m_lstList_AfterUpdate
End Function

Public Function TriggerCheckBoxUpdate()
    Call m_chkBox_AfterUpdate
End Function

Public Property Set CheckBox(ByRef chkCheckBox As Access.CheckBox)
    Set m_chkBox = chkCheckBox
    m_chkBox.AfterUpdate = "[Event Procedure]"
End Property

Public Property Set ListBox(ByRef lstListBox As Access.ListBox)
    Set m_lstList = lstListBox
    m_lstList.AfterUpdate = "[Event Procedure]"
End Property

Public Property Get ListBox()
    Set ListBox = m_lstList
End Property

Public Property Get CheckBox()
    Set CheckBox = m_chkBox
End Property

Private Sub m_chkBox_AfterUpdate()
    If m_DoingEvent Then
        Exit Sub
    End If
    m_DoingEvent = True
    
    Dim bSelectAll As Boolean
    Dim i As Long
    
    bSelectAll = Nz(m_chkBox.Value, False)
    
    'select or deselect all items
    For i = 0 To m_lstList.ListCount - 1
        m_lstList.Selected(i) = bSelectAll
    Next

    m_DoingEvent = False
End Sub


Private Sub m_lstList_AfterUpdate()
    If m_DoingEvent Then
        Exit Sub
    End If
    m_DoingEvent = True
    
    'if all selected, or nothing selected, set checkbox true
    If m_lstList.ItemsSelected.Count = 0 Or _
       m_lstList.ItemsSelected.Count = m_lstList.ListCount Then
        m_chkBox.Value = -1
    'otherwiese set checkbox false
    Else
        m_chkBox.Value = 0
    End If
    
    
    m_DoingEvent = False
End Sub
