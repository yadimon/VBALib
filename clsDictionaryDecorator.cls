VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDictionaryDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsDictionaryDecorator
' Author : Dmitry Gorelenkov
' Date   : 14.08.2014
' Purpose: Make Scripting.Dictionary more powerfull
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/clsDictionaryDecorator.cls</file>
'  <test>VBALib/tests/Tests_clsDictionaryDecorator.cls</test>
'  <license>PruefDBApp/license.bas</license>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Implements Scripting.Dictionary

Private m_ThisDictionary As Scripting.Dictionary

Public Property Set UsedDictionary(dictDictionary As Scripting.Dictionary)
    If Not m_ThisDictionary Is Nothing Then Set m_ThisDictionary = Nothing
    Set m_ThisDictionary = dictDictionary
End Property

Public Property Get UsedDictionary()
    Set UsedDictionary = m_ThisDictionary
End Property


'---------------------------------------------------------------------------------------
' Function: GetKeyByItem
'---------------------------------------------------------------------------------------
' Purpose: Get Key by Item
' Param  : Variant vItem item that bound to searched key
' Returns: Variant found key if found
' Remarks:
'---------------------------------------------------------------------------------------
Public Function GetKeyByItem(vItem As Variant) As Variant
    Dim vKey As Variant
    Dim bFound As Boolean
    
    For Each vKey In m_ThisDictionary.Keys
        'if objects
        If IsObject(vItem) And IsObject(m_ThisDictionary.Item(vKey)) Then
            If vItem Is m_ThisDictionary.Item(vKey) Then
                bFound = True
                Exit For
            End If
        'if simple values
        ElseIf Not IsObject(vItem) And Not IsObject(m_ThisDictionary.Item(vKey)) Then
            If vItem = m_ThisDictionary.Item(vKey) Then
                bFound = True
                Exit For
            End If
        End If
    Next vKey
    
    'nothing found? exit
    If Not bFound Then Exit Function
    
    'something found, return depends on key type.
    If IsObject(vKey) Then
        Set GetKeyByItem = vKey
    Else
        GetKeyByItem = vKey
    End If
End Function


Private Sub Class_Initialize()
    Set m_ThisDictionary = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set m_ThisDictionary = Nothing
End Sub


'**********************************************************************************
'REAL METHODS "Inherited"
'**********************************************************************************
Public Sub Add(Key As Variant, Item As Variant)
    Call m_ThisDictionary.Add(Key, Item)
End Sub

Public Property Let CompareMode(ByVal RHS As Scripting.CompareMethod)
    m_ThisDictionary.CompareMode = RHS
End Property

Public Property Get CompareMode() As Scripting.CompareMethod
    CompareMode = m_ThisDictionary.CompareMode
End Property

Public Property Get Count() As Long
    Count = m_ThisDictionary.Count
End Property

Public Function Exists(Key As Variant) As Boolean
    Exists = m_ThisDictionary.Exists(Key)
End Function

Public Property Get HashVal(Key As Variant) As Variant
    HashVal = m_ThisDictionary.HashVal(Key)
End Property

Public Property Let Item(Key As Variant, RHS As Variant)
    m_ThisDictionary.Item(Key) = RHS
End Property

Public Property Get Item(Key As Variant) As Variant
    If IsObject(m_ThisDictionary.Item(Key)) Then
        Set Item = m_ThisDictionary.Item(Key)
    Else
        Item = m_ThisDictionary.Item(Key)
    End If
End Property

Public Property Set Item(Key As Variant, RHS As Variant)
    Set m_ThisDictionary.Item(Key) = RHS
End Property

Public Function Items() As Variant
    Items = m_ThisDictionary.Items
End Function

Public Property Let Key(Key As Variant, RHS As Variant)
    m_ThisDictionary.Key(Key) = RHS
End Property

Public Function Keys() As Variant
    Keys = m_ThisDictionary.Keys
End Function

Public Sub Remove(Key As Variant)
    Call m_ThisDictionary.Remove(Key)
End Sub

Public Sub RemoveAll()
    Call m_ThisDictionary.RemoveAll
End Sub


'**********************************************************************************
'JUST FOR INTERFACE
'**********************************************************************************
Private Sub Dictionary_Add(Key As Variant, Item As Variant)
    Call Me.Add(Key, Item)
End Sub

Private Property Let Dictionary_CompareMode(ByVal RHS As Scripting.CompareMethod)
    Me.CompareMode = RHS
End Property

Private Property Get Dictionary_CompareMode() As Scripting.CompareMethod
    Dictionary_CompareMode = Me.CompareMode
End Property

Private Property Get Dictionary_Count() As Long
    Dictionary_Count = Me.Count
End Property

Private Function Dictionary_Exists(Key As Variant) As Boolean
    Dictionary_Exists = Me.Exists(Key)
End Function

Private Property Get Dictionary_HashVal(Key As Variant) As Variant
    Dictionary_HashVal = Me.HashVal(Key)
End Property

Private Property Let Dictionary_Item(Key As Variant, RHS As Variant)
    Me.Item(Key) = RHS
End Property

Private Property Get Dictionary_Item(Key As Variant) As Variant
    Dictionary_Item = Me.Item(Key)
End Property

Private Property Set Dictionary_Item(Key As Variant, RHS As Variant)
    Set Me.Item(Key) = RHS
End Property

Private Function Dictionary_Items() As Variant
    Dictionary_Items = Me.Items
End Function

Private Property Let Dictionary_Key(Key As Variant, RHS As Variant)
    Me.Key(Key) = RHS
End Property

Private Function Dictionary_Keys() As Variant
    Dictionary_Keys = Me.Keys
End Function

Private Sub Dictionary_Remove(Key As Variant)
    Call Me.Remove(Key)
End Sub

Private Sub Dictionary_RemoveAll()
    Call Me.RemoveAll
End Sub
