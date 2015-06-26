VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCollectionDecorator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsCollectionDecorator
' Author : Dmitry Gorelenkov
' Date   : 04.11.2014
' Purpose: Wrapper for collection interface objects (such supports "for each")
' Version: 0.9
' Remarks: Supports Arrays, Collection, Dictionary
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/clsCollectionDecorator.cls</file>
'  <test>VBALib/tests/Tests_clsCollectionDecorator.cls</test>
'  <license>VBALib/license.bas</license>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit


Private Enum eColType
    eType_array = 1
    eType_collection = 2
    eType_dictionary = 3
End Enum

Private m_colType As eColType
Private m_vCollection As Variant


Private m_bInitialized As Boolean

Public Function Init(ByRef vCollection As Variant) As Boolean
    If m_bInitialized Then Exit Function
    
    
    If IsObject(vCollection) Then
        Select Case TypeName(vCollection)
            Case "Dictionary"
                m_colType = eType_dictionary
            Case "Collection"
                m_colType = eType_collection
            Case Else
                GoTo Err
        End Select
        
        Set m_vCollection = vCollection
    ElseIf IsArray(vCollection) Then
        m_colType = eType_array
        m_vCollection = vCollection
    Else
        GoTo Err
    End If
    
    m_bInitialized = True
    Exit Function
    
Err:
    Debug.Print "clsCollectionDecorator: wrong argument"
    m_bInitialized = False
    
End Function

'---------------------------------------------------------------------------------------
' Function: contains
'---------------------------------------------------------------------------------------
' Purpose: test if vEmelent contains in loaded collection
' Param  : Variant vElement element to test
' Returns: Variant
' Remarks: "null" element supported
'---------------------------------------------------------------------------------------
Public Function Contains(vElement As Variant)
    Dim e As Variant
    For Each e In m_vCollection
        'in case its dictionary, e is key, -> convert to value
        If TypeName(m_vCollection) = "Dictionary" Then
            If IsObject(m_vCollection.Item(e)) Then
                Set e = m_vCollection.Item(e)
            Else
                e = m_vCollection.Item(e)
            End If
        End If
        
        If IsObject(e) And IsObject(vElement) Then
            If e Is vElement Then
                Contains = True
                Exit Function
            End If
        ElseIf Not IsObject(e) And Not IsObject(vElement) And TypeName(e) = TypeName(vElement) Then
            If e = vElement Or (IsNull(e) And IsNull(vElement)) Then
                Contains = True
                Exit Function
            End If
        End If
    Next e
End Function


'---------------------------------------------------------------------------------------
' Function: length
'---------------------------------------------------------------------------------------
' Purpose: returns length of collection
' Returns: Long length
'---------------------------------------------------------------------------------------
Public Function Length() As Long
    Dim Result As Long
    
    Select Case m_colType
        Case eType_dictionary
            Result = m_vCollection.Count
        Case eType_collection
            Result = m_vCollection.Count
        Case eType_array
            Result = UBound(m_vCollection) - LBound(m_vCollection) + 1
        Case Default
            Result = -1
            Debug.Print "wtf by length in clsCollectionDecorator" 'should never happen
    End Select
    
    Length = Result
End Function

Public Function Add(Value As Variant) As Long
'TODO
End Function

Public Function SetValue(Index As Long, Value As Variant) As Boolean
'TODO
End Function

Public Function Remove(Index As Long) As Variant
'    Select Case m_colType
'        Case eType_dictionary
'            result = m_vCollection.remove(m_vCollection.Keys(index))
'        Case eType_collection
'            While m_vCollection.Count > 0
'                result = result And m_vCollection.remove(0)
'            Wend
'            result = m_vCollection.Count
'        Case eType_array
'            result = UBound(m_vCollection) - LBound(m_vCollection) + 1
'        Case Default
'            result = -1
'            Debug.Print "wtf by clear in clsCollectionDecorator" 'should never happen
'    End Select
End Function

Public Function Clear() As Boolean
'    Dim result As Boolean
'    result = True
'
'    Select Case m_colType
'        Case eType_dictionary
'            result = m_vCollection.RemoveAll
'        Case eType_collection
'            While m_vCollection.Count > 0
'                result = result And m_vCollection.remove(0)
'            Wend
'            result = m_vCollection.Count
'        Case eType_array
'            result = UBound(m_vCollection) - LBound(m_vCollection) + 1
'        Case Default
'            result = -1
'            Debug.Print "wtf by clear in clsCollectionDecorator" 'should never happen
'    End Select
End Function

Public Function ToJSON() As String
'TODO
End Function

Public Function ToArray() As Variant
'TODO
End Function

Public Function ToCollection() As Collection
'TODO
End Function

Public Function ToDictionary() As Scripting.Dictionary
'TODO
End Function

Public Function toString() As String
'TODO
End Function

Public Function IsInitialized() As Boolean
    IsInitialized = m_bInitialized
End Function
