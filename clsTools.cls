VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsTools
' Author : Dmitry Gorelenkov
' Date   : 05.2012
' Purpose: Tools for VBA
' Version: 1.0
' Remarks: TODO refactor all the code. +UnitTests
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/clsTools.cls</file>
'  <license>VBALib/license.bas</license>
'  <ref><name>ADODB</name><major>6</major><minor>1</minor><guid>{B691E011-1797-432E-907A-4D8C69339129}</guid></ref>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'  <ref><name>VBScript_RegExp_55</name><major>5</major><minor>5</minor><guid>{3F4DACA7-160D-11D2-A8E9-00104B365C9F}</guid></ref>
'  <test>VBALib/tests/Tests_clsTools.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Explicit

'API showWindows, nCmdShow = 0 => ausblenden, = 1 => anzeigen
Private Declare Function APIShowWindow Lib "user32" Alias "ShowWindow" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub APISleep Lib "kernel32" Alias "Sleep" (ByVal lngMilliseconds As Long)
Private Declare Function APISetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal Hwnd As Long) As Long
Private Declare Function APIGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Long, ByVal fuWinIni As Long) As Long

Private Declare Function APIGetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long
         
Private Declare Function APIGetTempFileName Lib "kernel32" _
  Alias "GetTempFileNameA" _
  (ByVal lpszPath As String, _
  ByVal lpPrefixString As String, _
  ByVal wUnique As Long, _
  ByVal lpTempFileName As String) As Long
  
Private Declare Function SetWindowPos Lib "user32" _
 (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, _
 ByVal X As Long, ByVal y As Long, _
 ByVal cx As Long, ByVal cy As Long, _
 ByVal wFlags As Long) As Long

'call fn/sub by pointer (address of) + params
'THX: http://www.office-loesung.de/ftopic509235_0_0_asc.php
Private Declare Function APICallWindowProc _
            Lib "user32.dll" Alias "CallWindowProcA" ( _
            ByVal lpPrevWndFunc As Long, _
            Optional ByVal Hwnd As Long, _
            Optional ByVal Msg As Long, _
            Optional ByVal wParam As Long, _
            Optional ByVal lParam As Long) As Long
' ================
' SetWindowPos Info
' ================
Private Enum SWPValues
    SWP_SHOWWINDOW = &H40
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
End Enum
 
Private Const SPI_GETSCREENSAVERRUNNING As Long = &H72

'ein window zu Xpos, Ypos positionieren
Public Function MoveWindow(ByVal Hwnd As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Call SetWindowPos( _
     Hwnd, hWndInsertAfter:=0, X:=Xpos, y:=Ypos, cx:=0, cy:=0, _
     wFlags:=SWP_NOSIZE Or SWP_NOZORDER)
End Function

'checks if screensaver is running
'http://www.mrexcel.com/forum/excel-questions/523494-visual-basic-applications-determine-whether-screen-saver-running.html
Public Function IsScrSaverRunning() As Boolean

    Dim lScrSrvRunning As Long
            
    SystemParametersInfo _
    SPI_GETSCREENSAVERRUNNING, vbNull, lScrSrvRunning, vbNull
    
    IsScrSaverRunning = CBool(lScrSrvRunning)
           
End Function

Public Function GetActiveWindow() As Long
    GetActiveWindow = APIGetActiveWindow
End Function

'window auf vordergrund setzen, bei hWnd
Public Function SetForegroundWindow(ByVal Hwnd As Long)
    Call APISetForegroundWindow(Hwnd)
End Function

'sleep in MS
Public Function Sleep(lngMs As Long) As Boolean
    If lngMs > 0 Then
        Call APISleep(lngMs)
        Sleep = True
    End If
End Function

'fenster anzeigen / verstecken
Public Function ShowWindow(lHandleWnd As Long, Optional bShow As Boolean = True)
    On Error GoTo ErrorHandler
    If bShow Then
        ShowWindow = APIShowWindow(lHandleWnd, 1)
    Else
        ShowWindow = APIShowWindow(lHandleWnd, 0)
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print "Fehler bei clsTools.ShowWindow"
End Function

' Sekunden nach hh:mm:ss umrechnen
Public Function FormatSeconds(ByVal nSeconds As Long, _
     Optional ByVal sFormat As String = "hh:nn:ss") As String
    
     ' Falls anstelle nn mm im Ausgabeformat angegeben wurde
     sFormat = Replace(sFormat, "mm", "nn")
    
     FormatSeconds = Format$(DateAdd("s", nSeconds, CDate("00:00:00")), sFormat)
End Function

'return a copy of object
Public Function ObjCopy(ByVal whatToCopy As Object)
    If Not IsObject(whatToCopy) Then Exit Function
    Set ObjCopy = whatToCopy
End Function

'prueft ob der Wert datum (als text) oder numerische Datum ist
Public Function IsDateExt(vWert)
    On Error Resume Next
    IsDateExt = (IsNumeric(vWert) And Len(vWert) = 5) Or IsDate(vWert)
End Function

'return a copy of array
Public Function ArrayCopy(ByVal whatToCopy As Variant)
    If Not IsArray(whatToCopy) Then Exit Function
    ArrayCopy = whatToCopy
End Function

'returns length of array
Public Function ArrayLength(ByRef aArray As Variant)

    If IsEmpty(aArray) Or Not IsArray(aArray) Then
        ArrayLength = -1
        Exit Function
    End If
    
    ArrayLength = UBound(aArray) - LBound(aArray) + 1
End Function

'---------------------------------------------------------------------------------------
' Function: ArrayPush
'---------------------------------------------------------------------------------------
' Purpose: push Value to aArray array. Size of aArray will be increased by one
' Param  : Array aArray array the value will be pushed to
' Param  : Variant Value value to push
' Returns: Variant modified Array (first argument)
' Remarks: modifies aArray argument!
'---------------------------------------------------------------------------------------
Public Function ArrayPush(ByRef aArray As Variant, Value As Variant) As Variant
    Dim newLength As Long
    Dim minLength As Long
    
    On Error Resume Next
    
    If IsEmpty(aArray) Or Not IsArray(aArray) Then
        ReDim aArray(0)
    ElseIf UBound(aArray) = -1 Then
        ReDim aArray(0)
    Else
        newLength = UBound(aArray) + 1
        minLength = LBound(aArray)
        ReDim Preserve aArray(minLength To newLength)
    End If
    
    Call AssignValue(aArray(UBound(aArray)), Value)
    
    ArrayPush = aArray
    
    On Error GoTo 0
End Function


'---------------------------------------------------------------------------------------
' Function: ArrayPop
'---------------------------------------------------------------------------------------
' Purpose: removes last element of an array
' Param  : Array aArray array where to remove last element
' Returns: Variant removed element, or null if empty array or error
'---------------------------------------------------------------------------------------
Public Function ArrayPop(ByRef aArray() As Variant) As Variant
    ArrayPop = Null
    If ArrayLength(aArray) < 1 Then Exit Function
    'save last element to return
    If Not AssignValue(ArrayPop, aArray(UBound(aArray))) Then
        Exit Function
    End If
    'remove it from array
    Call ArrayRemoveByIndex(aArray, UBound(aArray))
End Function


'---------------------------------------------------------------------------------------
' Function: ArrayRemoveByElement
'---------------------------------------------------------------------------------------
' Purpose: remove first element with Value
' Param  : Array aArray array where to remove
' Param  : Variant Value value of element to remove
' Returns: Boolean success, false if no element removed
' Remarks: modifies aArray
'---------------------------------------------------------------------------------------
Public Function ArrayRemoveByElement(ByRef aArray() As Variant, Value As Variant) As Boolean
    Dim lIndex As Long
    ArrayRemoveByElement = False
    
    lIndex = ArrayIndexByValue(Value, aArray)
    
    If lIndex < 0 Then Exit Function
    
    ArrayRemoveByElement = ArrayRemoveByIndex(aArray, lIndex)
End Function

'---------------------------------------------------------------------------------------
' Function: ArrayRemoveByIndex
'---------------------------------------------------------------------------------------
' Purpose: remove element on lIndex position
' Param  : Array aArray array where to remove
' Param  : Long lIndex index of element to remove
' Returns: Boolean success, false if no element removed
' Remarks: modifies aArray
'---------------------------------------------------------------------------------------
Public Function ArrayRemoveByIndex(ByRef aArray(), lIndex As Long) As Boolean
    Dim i As Long
    
    ArrayRemoveByIndex = False
    If Not IsArray(aArray) Then Exit Function
    If Not ArrayIndexExists(lIndex, aArray) Then Exit Function
    
    'if only one element to remove, return empty array
    If ArrayLength(aArray) = 1 Then
        aArray = Array()
        ArrayRemoveByIndex = True
        Exit Function
    End If
    
    
    For i = lIndex To UBound(aArray) - 1
        Call AssignValue(aArray(i), aArray(i + 1))
    Next i
    
    ReDim Preserve aArray(LBound(aArray) To UBound(aArray) - 1)
    
    ArrayRemoveByIndex = True
End Function

'eleminiert Dublikate in einem array
Public Function ArrayRemoveDublicates(ByVal Arr)
    Dim Element As Variant

    
    If ArrayLength(Arr) <= 0 Then
        ArrayRemoveDublicates = Arr
        Exit Function
    End If
    
    Dim newUniqArray As Variant
    newUniqArray = Array()
    
    For Each Element In Arr
        If Not ArrayContains(newUniqArray, Element) Then
            Call ArrayPush(newUniqArray, Element)
        End If
    Next Element
    
    ArrayRemoveDublicates = newUniqArray
End Function

'return subArray of aArray, with lbound = 0
'lFrom from what element (for example from second = 1), lTo to what element, for example to 4. element = 3
Public Function ArrayGetSubArray(ByVal aArray As Variant, Optional ByVal lFrom As Variant = -1, Optional ByVal lTo As Variant = -1, Optional ByVal lStartBound As Variant = -1)
    Dim lBnd As Long, ubnd As Long
    lBnd = LBound(aArray)
    ubnd = UBound(aArray)
    
    
    
    If lFrom = -1 Then
        lFrom = lBnd
    Else
        lFrom = lBnd + lFrom
    End If
    
    
    If lTo = -1 Then
        lTo = ubnd
    Else
        lTo = lBnd + lTo
    End If
    
    
    If IsEmpty(aArray) Or Not IsArray(aArray) Then
        ArrayGetSubArray = aArray
        Exit Function
    End If
    
    
    If lFrom >= lBnd And lTo <= ubnd And lFrom <= lTo Then
        Dim tmpArray As Variant
        Dim i As Long
        ReDim tmpArray(lFrom To lTo)
        
        For i = lFrom To lTo
            tmpArray(i) = aArray(i)
        Next i
        
        aArray = tmpArray
    End If
    
    'startbound setzen
    If lStartBound >= 0 Then
        Call ArrayChangeBounds(aArray, lStartBound)
    End If
    
    
    ArrayGetSubArray = aArray
End Function

'changes bounds of array.
'todo optional vars
Public Function ArrayChangeBounds(ByRef aArray As Variant, lStartBound As Variant, Optional lEndBound As Variant = -1)
    If Not IsArray(aArray) Then
        ArrayChangeBounds = aArray
        Exit Function
    End If
        
    If lEndBound = -1 Then lEndBound = lStartBound + ArrayLength(aArray) - 1
        
    If lStartBound < 0 Or lEndBound < 0 Or lStartBound > lEndBound Then
        ArrayChangeBounds = aArray
        Exit Function
    End If
    
    
    If IsEmpty(aArray) Then
        ReDim Preserve aArray(lStartBound To lEndBound)
        ArrayChangeBounds = aArray
        Exit Function
    End If
    
    Dim i As Long
    Dim lBnd As Long
    Dim tmpArray As Variant
    ReDim tmpArray(lStartBound To lEndBound)
    
    'werte nacheinanden kopieren
    lBnd = LBound(aArray)
    For i = lBnd To UBound(aArray)
        'nur falls es noch in tmpArray platz gibt
        If (i - lBnd <= lEndBound - lStartBound) Then
            Call AssignValue(tmpArray(lStartBound + i - lBnd), aArray(i))
        End If
    Next i
    
    aArray = tmpArray
    ArrayChangeBounds = tmpArray
End Function

'concatenates 2 arrays, array1 wird groesser, und kopie wird zuruckgegeben 'TODO kopie nur bei einem optional param zurueckgeben?
Public Function ArrayConcat(ByRef array1 As Variant, ByRef array2 As Variant) As Variant
    Dim vResult As Variant
    Dim idx As Long
    Dim Elem As Variant
    
    vResult = Array()
    
    On Error GoTo arrayConcat_Error
    
    'falls nicht arrays exit
    If Not (IsArray(array1) And IsArray(array2)) Then
        Dbg "Falsche argumente bei arrayConcat"
        GoTo Final
    End If
    
    'falls einer der beiden ist leer, return den anderen
    If ArrayLength(array1) = 0 Then
        vResult = ArrayCopy(array2)
        GoTo Final
    ElseIf ArrayLength(array2) = 0 Then
        vResult = ArrayCopy(array1)
        GoTo Final
    End If
    
    vResult = ArrayCopy(array1)
    
    
    idx = UBound(vResult)
    
    'array auf neue geoesse einstellen
    ReDim Preserve vResult(LBound(array1) To idx + ArrayLength(array2))
    
    'jedes elemen von array2 einfuegen
    idx = idx + 1
    For Each Elem In array2
        Call AssignValue(vResult(idx), Elem)
        idx = idx + 1
    Next Elem
    
    
    
Final:
    ArrayConcat = vResult
    
    On Error GoTo 0
    Exit Function
arrayConcat_Error:
    Dbg "arrayConcat of Klassenmodul clsTools"
    GoTo Final
End Function

'array must be sorted!!!
'binary search in array. In case nothing found returns -1
Public Function ArrayBinarySearch(strArray As Variant, strSearch As String) As Long
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim lngMiddle As Long
    Dim bolInverseOrder As Boolean
    If Not IsArray(strArray) Then Exit Function
    lngFirst = LBound(strArray)
    lngLast = UBound(strArray)
    bolInverseOrder = (strArray(lngFirst) > strArray(lngLast))
    ArrayBinarySearch = -1 'lngFirst - 1
    Do
        lngMiddle = (lngFirst + lngLast) \ 2
        If strArray(lngMiddle) = strSearch Then
            ArrayBinarySearch = lngMiddle
            Exit Do
        ElseIf ((strArray(lngMiddle) < strSearch) Xor bolInverseOrder) Then
            lngFirst = lngMiddle + 1
        Else
            lngLast = lngMiddle - 1
        End If
    Loop Until lngFirst > lngLast
End Function

'---------------------------------------------------------------------------------------
' Function: ArrayContains
'---------------------------------------------------------------------------------------
' Purpose: test if vValue contains in aArray
' Param  : Variant vValue value looking for. Must be simple type, (not object), like String or Integer (Cstr must be able to apply on it)
' Param  : Variant aArray array where to search
' Param  : Boolean bAnyTypeCompare (False) set to true, if need to compare e.g. numbers with strings
' Param  : Boolean bTrim (False) set to true for trim values before compare
' Param  : Boolean bCaseInsensive (False) set to true for make the compare case insensive
' Returns: Boolean true if contains, false otherwise
' Remarks:
'---------------------------------------------------------------------------------------
Public Function ArrayContains(ByVal vValue As Variant, ByRef aArray As Variant, _
                              Optional bAnyTypeCompare As Boolean = False, _
                              Optional bTrim As Boolean = False, _
                              Optional bCaseInsensive As Boolean = False) As Boolean
                              
    ArrayContains = (ArrayIndexByValue(vValue, aArray, bAnyTypeCompare, bTrim, bCaseInsensive) >= 0)
End Function

'---------------------------------------------------------------------------------------
' Function: ArrayIndexExists
'---------------------------------------------------------------------------------------
' Purpose: tests if index in the array is valid (exists)
' Param  : Long lIndex index to test
' Param  : Variant vArray array where the index will be tested
' Returns: Boolean true if index exists, false otherwiese
'---------------------------------------------------------------------------------------
Public Function ArrayIndexExists(lIndex As Long, aArray As Variant) As Boolean

    'bad cases
    If Not IsArray(aArray) Or IsObject(aArray) Then
        Err.Raise vbObjectError, "clsTools:ArrayIndexExists", "Wrong argument type, aArray"
    End If
    
    If lIndex <= UBound(aArray) And lIndex >= LBound(aArray) Then
        ArrayIndexExists = True
    Else
        ArrayIndexExists = False
    End If

End Function



'---------------------------------------------------------------------------------------
' Function: ArrayIndexByValue
'---------------------------------------------------------------------------------------
' Purpose: returns index of element first found in array
' Param  : Variant vValue value to search for
' Param  : Variant aArray array where to search
' Param  : Boolean bAnyTypeCompare (False) set to true, if need to compare e.g. numbers with strings
' Param  : Boolean bTrim (False) set to true for trim values before compare
' Param  : Boolean bCaseInsensive (False) set to true for make the compare case insensive
' Returns: Long index of element, or -1 if nothing found
' Remarks: TODO use ArrayAllIndexesByValue ? -perfomance +code
'---------------------------------------------------------------------------------------
Public Function ArrayIndexByValue(vValue As Variant, aArray As Variant, _
                                  Optional bAnyTypeCompare As Boolean = False, _
                                  Optional bTrim As Boolean = False, _
                                  Optional bCaseInsensive As Boolean = False) As Long
    Dim i As Long

    ArrayIndexByValue = -1
    Dim vValueArray As Variant
    Dim vValueSearch As Variant

    

    Call UnifyValueIfPossible(vValue, vValueSearch, bAnyTypeCompare, bTrim, bCaseInsensive)

    For i = LBound(aArray) To UBound(aArray)
        Call UnifyValueIfPossible(aArray(i), vValueArray, bAnyTypeCompare, bTrim, bCaseInsensive)

        If CompareValuesIfPossible(vValueSearch, vValueArray) = 0 Then
            ArrayIndexByValue = i
            Exit For
        End If
    Next i
End Function

'---------------------------------------------------------------------------------------
' Function: UnifyValueIfPossible
'---------------------------------------------------------------------------------------
' Purpose: make value e.g. as trimmed lowcase string, for better compare ability. If possible.
' Param  : Variant vValue value fo unify
' Param  : Variant vResult reference to variable where the result will be saved
' Param  : Boolean bToString (False) set to true, to modify the value to string
' Param  : Boolean bTrim (False) set to true to trim the value
' Param  : Boolean bToLowerCase (False) set to true to make the value lower case
' Returns: Boolean true if value is not array or object, and can be modified
'---------------------------------------------------------------------------------------
Public Function UnifyValueIfPossible(vValue As Variant, ByRef vResult As Variant, _
                                     Optional bToString As Boolean = False, _
                                     Optional bTrim As Boolean = False, _
                                     Optional bCaseInsensive As Boolean = False) As Boolean
                                     
    'nothing to modify
    If IsObject(vValue) Then
        Set vResult = vValue
        UnifyValueIfPossible = False
        Exit Function
    End If
    
    If IsArray(vValue) Then
        vResult = vValue
        UnifyValueIfPossible = False
        Exit Function
    End If

    vResult = vValue
    ' number 5 = string "5"
    If bToString Then
        vResult = CStr(vResult)
    End If
    
    ' " text " = "text"
    If bTrim Then
        vResult = Trim(vResult)
    End If
    
    ' "TeXt" = "texT"
    If bCaseInsensive Then
        vResult = LCase(vResult)
    End If

    UnifyValueIfPossible = True
End Function


'---------------------------------------------------------------------------------------
' Function: CompareValuesIfPossible
'---------------------------------------------------------------------------------------
' Purpose: compare 2 variable if possible
' Param  : Variant vValue1 first value to compare
' Param  : Variant vValue2 second value to compare
' Returns: Long 0 if equal, -1 if not comparable, 1 if are not equal
' Remarks: TODO add less, bigger info return values
'---------------------------------------------------------------------------------------
Public Function CompareValuesIfPossible(vValue1 As Variant, vValue2 As Variant) As Long
    CompareValuesIfPossible = -1
    
    'arrays not comparable
    If IsArray(vValue1) Or IsArray(vValue2) Then Exit Function
    
    If IsObject(vValue1) And IsObject(vValue2) Then

        CompareValuesIfPossible = IIf(vValue1 Is vValue2, 0, 1)

    ElseIf Not IsObject(vValue1) And Not IsObject(vValue2) Then
        
        'if both strings, or not strings and numeric => are comparable
        If (TypeName(vValue1) = "String" And TypeName(vValue2) = "String") _
        Or (TypeName(vValue1) <> "String" And TypeName(vValue2) <> "String" And IsNumeric(vValue1) And IsNumeric(vValue2)) Then
            CompareValuesIfPossible = IIf(vValue1 = vValue2, 0, 1)
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Function: ArrayIsAllocated
'---------------------------------------------------------------------------------------
' Purpose: test whether array allocated or not
' Param  : Variant aArray test object
' Returns: Boolean true if allocated
' Remarks: pearson: http://www.cpearson.com/excel/isarrayallocated.aspx
'---------------------------------------------------------------------------------------
Public Function ArrayIsAllocated(aArray As Variant) As Boolean
    On Error Resume Next
        ArrayIsAllocated = IsArray(aArray) And _
                           Not IsError(LBound(aArray, 1)) And _
                           LBound(aArray, 1) <= UBound(aArray, 1)
                           
    'TODO return Not Not aArray  ?
End Function

'---------------------------------------------------------------------------------------
' Function: ArrayAllIndexesByValue
'---------------------------------------------------------------------------------------
' Purpose: returns array of indexes of element found in array
' Param  : Valiant vValue value to search for
' Param  : Valiant aArray array where to search
' Param  : Boolean bAnyTypeCompare (False) set to true, if need to compare e.g. numbers with strings
' Param  : Boolean bTrim (False) set to true for trim values before compare
' Param  : Boolean bCaseInsensive (False) set to true for make the compare case insensive
' Returns: Variant
' Remarks:
'---------------------------------------------------------------------------------------
Public Function ArrayAllIndexesByValue(vValue As Variant, aArray As Variant, _
                                       Optional bAnyTypeCompare As Boolean = False, _
                                       Optional bTrim As Boolean = False, _
                                       Optional bCaseInsensive As Boolean = False) As Variant
    Dim i As Long
    Dim retArray As Variant

    ArrayAllIndexesByValue = Array()
    Dim vValueArray As Variant
    Dim vValueSearch As Variant

    If Not IsArray(aArray) Then Exit Function

    Call UnifyValueIfPossible(vValue, vValueSearch, bAnyTypeCompare, bTrim, bCaseInsensive)

    For i = LBound(aArray) To UBound(aArray)
        Call UnifyValueIfPossible(aArray(i), vValueArray, bAnyTypeCompare, bTrim, bCaseInsensive)

        If CompareValuesIfPossible(vValueSearch, vValueArray) = 0 Then
            Call ArrayPush(retArray, i)
        End If
    Next i

    ArrayAllIndexesByValue = retArray

End Function


'---------------------------------------------------------------------------------------
' Function: ArrayToCollection
'---------------------------------------------------------------------------------------
' Purpose: convert an array to collection
' Param  : aArray array to convert
' Returns: Collection
' Remarks: key in collection is index in array as string. e.g. "0", "1", "2" ..
'---------------------------------------------------------------------------------------
Public Function ArrayToCollection(aArray As Variant) As Collection
    Dim colReturn As Collection
    Set colReturn = New Collection
    Dim i As Long
    
    For i = LBound(aArray) To UBound(aArray)
        Call colReturn.Add(aArray(i), CStr(i))
    Next i
    
    Set ArrayToCollection = colReturn
End Function


'---------------------------------------------------------------------------------------
' Function: ArrayToDictionary
'---------------------------------------------------------------------------------------
' Purpose: converts an array to dictionary
' Param  : aArray array to convert
' Returns: Scripting.Dictionary
' Remarks: key will be the index in array, as long
'---------------------------------------------------------------------------------------
Public Function ArrayToDictionary(aArray As Variant) As Scripting.Dictionary
    Dim dic As Scripting.Dictionary
    Set dic = New Scripting.Dictionary
    Dim i As Long
    
    For i = LBound(aArray) To UBound(aArray)
        Call dic.Add(i, aArray(i))
    Next i
    
    Set ArrayToDictionary = dic
End Function

'Collection in ein array umwandeln
Public Function CollectionToArray(colCollection As Collection) As Variant
    If colCollection Is Nothing Then Exit Function
    Dim wert As Variant
    Dim aReturn() As Variant
    aReturn = Array()
    For Each wert In colCollection
        Call ArrayPush(aReturn, wert)
    Next wert
    CollectionToArray = aReturn
End Function




'---------------------------------------------------------------------------------------
' Function: arrayXDimTo1Dim
'---------------------------------------------------------------------------------------
' Purpose: 'push each element from X-dimensional array to new 1-dim array
' Param  : Variant arr array to merge
' Returns: Variant new array with 1 dimension
'---------------------------------------------------------------------------------------
Public Function ArrayXDimTo1Dim(ByVal aArray As Variant) As Variant
    Dim tmpArray()
    Dim Element As Variant
    Dim tmpArrayOfElem As Variant
    
    If Not IsArray(aArray) Then
        Err.Raise vbObjectError, "clsTools.arrayXDimTo1Dim", "Wrong type of arg, is not an array"
    End If
    
    tmpArray = Array()
    
    For Each Element In aArray
        If IsArray(Element) Then 'if array again, merge recursive
            tmpArrayOfElem = ArrayXDimTo1Dim(Element)
            tmpArray = ArrayConcat(tmpArray, tmpArrayOfElem)  'add new array to current array
        Else
            Call ArrayPush(tmpArray, Element) 'add only element to current array
        End If
        
    Next Element
    
    ArrayXDimTo1Dim = tmpArray
End Function

'---------------------------------------------------------------------------------------
' Function: ArrayShuffle
'---------------------------------------------------------------------------------------
' Purpose: make copy of array, shuffeled
' Param  : Array InArray source array
' Returns: Variant shuffeled array
' Remarks: http://www.cpearson.com/excel/ShuffleArray.aspx, modified
'---------------------------------------------------------------------------------------
Public Function ArrayShuffle(InArray() As Variant) As Variant()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ArrayShuffle
' This function returns the values of InArray in random order. The original
' InArray is not modified.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Arr() As Variant
    
    Arr = ArrayCopy(InArray)
    Call ArrayShuffleInPlace(Arr)
    
    ArrayShuffle = Arr
End Function

'---------------------------------------------------------------------------------------
' Sub: ArrayShuffleInPlace
'---------------------------------------------------------------------------------------
' Purpose: shuffle array in place
' Param  : Array InArray source array
' Remarks: http://www.cpearson.com/excel/ShuffleArray.aspx, modified
'---------------------------------------------------------------------------------------
Public Sub ArrayShuffleInPlace(InArray() As Variant)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ArrayShuffleInPlace
' This shuffles InArray to random order, randomized in place.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim n As Long
    Dim Temp As Variant
    Dim j As Long
   
    Randomize
    For n = LBound(InArray) To UBound(InArray)
        j = Int(((UBound(InArray) - n + 1) * Rnd) + n)
        If n <> j Then
            Call AssignValue(Temp, InArray(n))
            Call AssignValue(InArray(n), InArray(j))
            Call AssignValue(InArray(j), Temp)
        End If
    Next n
End Sub


'max of both values
Public Function Max(value1 As Variant, Value2 As Variant) As Variant
    If IsEmpty(value1) Or IsNull(value1) Then
        Max = Value2
        Exit Function
    ElseIf IsEmpty(Value2) Or IsNull(Value2) Then
        Max = value1
        Exit Function
    Else
        If IsNumeric(value1) And IsNumeric(Value2) Then
            If value1 > Value2 Then
                Max = value1
            Else
                Max = Value2
            End If
        End If
    End If
End Function

'min of both values
Public Function Min(value1 As Variant, Value2 As Variant) As Variant
    If IsEmpty(value1) Or IsNull(value1) Then
        Min = Value2
        Exit Function
    ElseIf IsEmpty(Value2) Or IsNull(Value2) Then
        Min = value1
        Exit Function
    Else
        If IsNumeric(value1) And IsNumeric(Value2) Then
            If value1 < Value2 Then
                Min = value1
            Else
                Min = Value2
            End If
        End If
    End If
End Function

'filePicker or FolderPicker dialog. 'todo miltiselect option ?
'required Office 11? Object Library
Public Function DialogGetPath(Optional bFileOrFolder As Boolean = True, Optional sTitle As String, Optional filterTyp As String, _
                        Optional filterExt As String, Optional sInitial As String, Optional bMulti As Boolean = False, Optional sButtonName As String)
                        
    Dim fd As FileDialog

    'if true -> file
    If bFileOrFolder Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        'falls filter werde vorhanden
        If IsSet(filterTyp) And IsSet(filterExt) Then
            fd.Filters.Add filterTyp, filterExt
        End If
        'extra button title
        If IsSet(sButtonName) Then
            fd.ButtonName = sButtonName
        End If
    
    'folder
    Else
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    
    'anfangspfad
    If IsSet(sInitial) Then fd.InitialFileName = sInitial
    
    'if multiselected
    If bMulti Then
        fd.AllowMultiSelect = True
    End If
    
    If IsSet(sTitle) Then
        fd.Title = sTitle
    End If
    
    
    
    Dim vrtResult As Variant
    Dim vrtSelectedItem As Variant
    
    If fd.Show = -1 Then

        'Step through each string in the FileDialogSelectedItems collection.
        For Each vrtSelectedItem In fd.SelectedItems
            'extra slash fuer ordner
            If Not bFileOrFolder Then
                If Right(vrtSelectedItem, 1) <> "\" Then vrtSelectedItem = vrtSelectedItem & "\"
            End If
        
            'falls multidatei, dann als Array zuruckgeben
            If bMulti Then
                Call ArrayPush(vrtResult, vrtSelectedItem)
            Else
                vrtResult = vrtSelectedItem
                Exit For 'zur sicherheit
            End If
            
            
            
        Next vrtSelectedItem
    'The user pressed Cancel.
    Else
        vrtResult = ""
    End If

        
    DialogGetPath = vrtResult
End Function

'prueft ob vCheck ein wert hat, also ob es nicht 0 und nicht "" ist.
'bTrim ob auch trim eingesetz werden muss /default ja
Public Function IsSet(ByVal vCheck As Variant, Optional bTrim As Boolean = True)

    If IsNull(vCheck) Then
        IsSet = False
        Exit Function
    End If
    
    If IsObject(vCheck) Then
        If vCheck Is Nothing Then
            IsSet = False
        Else
            IsSet = True
        End If
        Exit Function
    End If
    
    If IsArray(vCheck) Then
        IsSet = True
        Exit Function
    End If
    
    
    If IsEmpty(vCheck) Then
        IsSet = False
        Exit Function
    End If
    
    If bTrim Then
        vCheck = Trim(vCheck)
    End If
    
    If Not (CStr(vCheck) = vbNullString Or vCheck = 0 Or Len(CStr(vCheck)) = 0) Then
        IsSet = True
        Exit Function
    End If
    
    IsSet = False
End Function

'prueft isSet fuer alle variables in Array
Public Function IsSetAll(ByVal aTxtCheck As Variant, Optional bTrim As Boolean = True)
    Dim Element As Variant
    IsSetAll = True
    For Each Element In aTxtCheck
        If Not IsSet(Element, bTrim) Then
            IsSetAll = False
            Exit For
        End If
    Next Element
End Function

'wenn value keinen wert hat dann wird vInCaseNoValue zuruckgeben, sonst das vValue
Public Function IfNoValue(vValue As Variant, vInCaseNoValue As Variant)
    If IsSet(vValue) Then
        Call AssignValue(IfNoValue, vValue)
    Else
        Call AssignValue(IfNoValue, vInCaseNoValue)
    End If
End Function

'check if directory exists
'https://svn.access-codelib.net/svn/codelib/branches/draft/file/FileTools.bas
Public Function DirExists(ByVal FullPath As String) As Boolean

   If Right$(FullPath, 1) <> "\" Then
      FullPath = FullPath & "\"
   End If

   DirExists = (Dir$(FullPath, vbDirectory Or vbReadOnly Or vbHidden Or vbSystem) = ".")
   Dir$ "\"
End Function

'check if file exists
Public Function FileExists(sPath As String)
    On Error Resume Next
    FileExists = True
    
    If sPath = VBA.vbNullString Or VBA.Right$(sPath, 1) = "\" Then
        FileExists = False
        Exit Function
    End If
    
    If VBA.Dir(sPath) = "" Then
        FileExists = False
    End If
End Function

'check if all files exists
Public Function FileExistsAll(ParamArray aPaths() As Variant)
    On Error GoTo ErrorHandler
    Dim sPath As Variant
    
    'if no parameters
    If UBound(aPaths) < 0 Then
        FileExistsAll = True
    End If
    'in case first arg is array, than use this array
    If IsArray(aPaths(0)) Then
        aPaths = aPaths(0)
    End If
    
    
    FileExistsAll = True
    For Each sPath In aPaths
        If Not FileExists(CStr(sPath)) Then
            FileExistsAll = False
            Exit For
        End If
    Next sPath

    
    Exit Function
ErrorHandler:
    FileExistsAll = False
End Function

'---------------------------------------------------------------------------------------
' Function: FilesList
'---------------------------------------------------------------------------------------
' Purpose: get array with files/folders found by the pattern
' Param  : String sPattern pattern to search for
' Param  : VbFileAttribute FileAttr file/folder types
' Returns: Variant
' Remarks: uses Dir
'---------------------------------------------------------------------------------------
Public Function FilesList(sPattern As String, Optional FileAttr As VbFileAttribute = vbNormal) As Variant
    Dim aFiles As Variant
    aFiles = Array()
    Dim sFilePath As String
    sFilePath = Dir$(sPattern, FileAttr)
    
    While Len(sFilePath) <> 0
        Call ArrayPush(aFiles, sFilePath)
        sFilePath = Dir$()
    Wend

    FilesList = aFiles
End Function

'quicksort, wegen kaskadenrekursion nur fuer kleine array benutzen..
Public Sub QuickSort(ByRef sArray As Variant, Optional ByVal MinElem As Long = -1, Optional MaxElem As Long = -1)
'
' QuickSort()
'
' Sortieren eines Arrays mit dem QuickSort-Algorithmus, dem wohl schnellsten
' Sortieralgorithmus von Welt.
'
' IN:   sArray      Array das sortiert werden soll
'       MinElem     erstes Element des Arrays (oder Teil-Arrays) default: LBound(sArray)
'       MaxElem     letztes Element des Arrays (oder Teil-Arrays) default: UBound(sArray)
'

If IsMissing(MinElem) Or MinElem = -1 Then MinElem = LBound(sArray)
If IsMissing(MaxElem) Or MaxElem = -1 Then MaxElem = UBound(sArray)

Dim Mitte As Long
Dim vDummy As Variant
Dim vMitte As Variant
Dim i As Long, j As Long
    '
    ' Abbruchbedingung der Rekursion prüfen
    '
    If MinElem > MaxElem Then
        '
        ' Rekursion beenden
        '
        Exit Sub
    End If
    '
    ' Ermittlung der Mitte des Arrays
    '
    Mitte = (MinElem + MaxElem) \ 2
    vMitte = sArray(Mitte)
    '
    ' Für die Prüfung der linken und rechten
    ' Seite die Zähler initialisieren
    '
    i = MinElem
    j = MaxElem
    Do
        '
        ' Von links bis zur Mitte prüfen
        '
        Do While sArray(i) < vMitte
            i = i + 1
        Loop
        '
        ' Von rechts bis zur Mitte prüfen
        '
        Do While sArray(j) > vMitte
            j = j - 1
        Loop
        
        If i <= j Then
            '
            ' Die beiden gefundenen, falsch einsortierten
            ' Elemente vertauschen
            '
            vDummy = sArray(j)
            sArray(j) = sArray(i)
            sArray(i) = vDummy
            '
            ' Prüfung bei der nächsten Position fortsetzen
            '
            i = i + 1
            j = j - 1
        End If
        
    Loop Until i > j
    '
    ' Rekursiver Aufruf mit den Teil-Arrays
    '
    QuickSort sArray, MinElem, j
    QuickSort sArray, i, MaxElem
End Sub

'sorts first elements of multidim. array(x,1) 0,1
Public Sub QuickSort2d(ByRef sArray As Variant, ByVal MinElem As Long, MaxElem As Long)
'
' QuickSort()
'
' Sortieren eines Arrays mit dem QuickSort-Algorithmus, dem wohl schnellsten
' Sortieralgorithmus von Welt.
'
' IN:   sArray      Array das sortiert werden soll
'       MinElem     erstes Element des Arrays (oder Teil-Arrays)
'       MaxElem     letztes Element des Arrays (oder Teil-Arrays)
'
Dim Mitte As Long
Dim vDummy As Variant
Dim vDummy2 As Variant
Dim vMitte As Variant
Dim i As Long, j As Long
    '
    ' Abbruchbedingung der Rekursion prüfen
    '
    If MinElem > MaxElem Then
        '
        ' Rekursion beenden
        '
        Exit Sub
    End If
    '
    ' Ermittlung der Mitte des Arrays
    '
    Mitte = (MinElem + MaxElem) \ 2
    vMitte = sArray(Mitte, 0)
    '
    ' Für die Prüfung der linken und rechten
    ' Seite die Zähler initialisieren
    '
    i = MinElem
    j = MaxElem
    Do
        '
        ' Von links bis zur Mitte prüfen
        '
        Do While sArray(i, 0) < vMitte
            i = i + 1
        Loop
        '
        ' Von rechts bis zur Mitte prüfen
        '
        Do While sArray(j, 0) > vMitte
            j = j - 1
        Loop
        
        If i <= j Then
            '
            ' Die beiden gefundenen, falsch einsortierten
            ' Elemente vertauschen
            '
            vDummy = sArray(j, 0)
            vDummy2 = sArray(j, 1)
            sArray(j, 0) = sArray(i, 0)
            sArray(j, 1) = sArray(i, 1)
            sArray(i, 0) = vDummy
            sArray(i, 1) = vDummy2
            '
            ' Prüfung bei der nächsten Position fortsetzen
            '
            i = i + 1
            j = j - 1
        End If
        
    Loop Until i > j
    '
    ' Rekursiver Aufruf mit den Teil-Arrays
    '
    QuickSort2d sArray, MinElem, j
    QuickSort2d sArray, i, MaxElem
End Sub

'Heap Sort. 'http://www.vbforums.com/showpost.php?p=2909250&postcount=7
Public Function HeapSort(ByRef pvarArray As Variant)
        Dim i As Long
        Dim iMin As Long
        Dim iMax As Long
        Dim varSwap As Variant
       
        iMin = LBound(pvarArray)
        iMax = UBound(pvarArray)
        For i = (iMax + iMin) \ 2 To iMin Step -1
            HeapSortHelper pvarArray, i, iMin, iMax
        Next i
        For i = iMax To iMin + 1 Step -1
            varSwap = pvarArray(i)
            pvarArray(i) = pvarArray(iMin)
            pvarArray(iMin) = varSwap
            HeapSortHelper pvarArray, iMin, iMin, i - 1
        Next i
        HeapSort = pvarArray
    End Function

Private Function HeapSortHelper(ByRef pvarArray As Variant, ByVal i As Long, iMin As Long, iMax As Long)
        Dim lngLeaf As Long
        Dim varSwap As Variant
       
        Do
            lngLeaf = i + i - (iMin - 1)
            Select Case lngLeaf
                Case Is > iMax: Exit Do
                Case Is < iMax: If pvarArray(lngLeaf + 1) > pvarArray(lngLeaf) Then lngLeaf = lngLeaf + 1
            End Select
            If pvarArray(i) > pvarArray(lngLeaf) Then Exit Do
            varSwap = pvarArray(i)
            pvarArray(i) = pvarArray(lngLeaf)
            pvarArray(lngLeaf) = varSwap
            i = lngLeaf
        Loop
    End Function

'http://www.freevbcode.com/ShowCode.asp?ID=1086
Public Function IsAlphaBetical(TestString As String) As Boolean
    
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in a string are alphabetical
    'returns false otherwise or for empty string
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[A-Za-z]" Then Exit Function
        Next
    
    IsAlphaBetical = True
    End If
    
End Function

'http://www.freevbcode.com/ShowCode.asp?ID=1086
Public Function IsAlphaNumeric(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in a string are alphabetical
    '   or numeric
    'returns false otherwise or for empty string
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9A-Za-z]" Then Exit Function
        Next
    
    IsAlphaNumeric = True
    End If
    
End Function

'http://www.freevbcode.com/ShowCode.asp?ID=1086
Public Function IsNumericOnly(TestString As String) As Boolean
    
    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in string are numeric
    'returns false otherwise or for empty string
    
    'this is different than VB's isNumeric
    'isNumeric returns true for something like 90.09
    'This Public Function will return false
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9]" Then Exit Function
        Next
    
    IsNumericOnly = True
    End If
    
End Function

'numeric type BEGIN http://www.ms-office-forum.net/forum/showthread.php?t=129184 ---------
'modified by me
Public Function IsDouble(var As Variant) As Boolean
   If IsNull(var) Then
      IsDouble = False
   Else
      IsDouble = ((VBA.VarType(var) = vbDouble) Or Me.IsSingle(var))
   End If
End Function

Public Function IsSingle(var As Variant) As Boolean
   If IsNull(var) Then
      IsSingle = False
   Else
      IsSingle = ((VBA.VarType(var) = vbSingle) Or Me.isLong(var))
   End If
End Function

Public Function isLong(var As Variant) As Boolean
   If IsNull(var) Then
      isLong = False
   Else
      isLong = ((VBA.VarType(var) = vbLong) Or Me.IsInteger(var))
   End If
End Function

Public Function IsInteger(var As Variant) As Boolean
   If IsNull(var) Then
      IsInteger = False
   Else
      IsInteger = (VBA.VarType(Val(var)) = vbInteger Or Me.IsByte(var))
   End If
End Function

Public Function IsByte(var As Variant) As Boolean
   If IsNull(var) Then
      IsByte = False
   Else
      IsByte = ((VBA.VarType(Val(var)) = vbByte))
   End If
End Function
'numeric type END ------------------------------------------------------------------------

Public Function RandRange(num1 As Variant, num2 As Variant)
    
    If IsNumeric(num1) And IsNumeric(num2) Then
        RandRange = Int(((num2 - num1 + 1) * Rnd) + num1)
    Else
        Debug.Print "RandRange, nicht numerische arguments"
    End If
    
End Function

'TODO
'tries to modify to date, if failed returns input, or 2nd arg of input
Public Function CdateTry(something As Variant, Optional inFailCaseReturn As Variant) As Variant
    Dim sReplaced As String
    Dim Delimiter As Variant 'string
    Dim aSplit As Variant
    Dim sDay As String
    
    If TypeName(something) = "Date" Or IsDate(something) Then
        CdateTry = CDate(something)
        Exit Function
    Else
        On Error Resume Next
        sReplaced = CStr(something)
        For Each Delimiter In Array(".", ",", "'")
            sReplaced = Replace(sReplaced, Delimiter, "/")
        Next Delimiter
        
        'replace month with day, for formats dd.mm.yyyy
        aSplit = Split(sReplaced, "/")
        sDay = aSplit(0)
        aSplit(0) = aSplit(1)
        aSplit(1) = sDay
        sReplaced = Join(aSplit, "/")
        
        'try to parse date now
        CdateTry = CDate(sReplaced)
        
        If IsDate(CdateTry) Then Exit Function
        On Error GoTo 0
    End If

    If IsMissing(inFailCaseReturn) Then
        Call AssignValue(CdateTry, something)
    Else
        Call AssignValue(CdateTry, inFailCaseReturn)
    End If
End Function

'tries to modify to Long, if failed returns input, or 2nd arg of input
Public Function ClngTry(something As Variant, Optional inFailCaseReturn As Variant) As Variant
    If isLong(something) Then
        ClngTry = CLng(something)
        Exit Function
    ElseIf Not IsEmpty(something) And Not IsNull(something) Then
        On Error Resume Next
        ClngTry = CLng(something)
        If isLong(ClngTry) Then Exit Function
        On Error GoTo 0
    End If
    
    
     If IsMissing(inFailCaseReturn) Then
        Call AssignValue(ClngTry, something)
    Else
        Call AssignValue(ClngTry, inFailCaseReturn)
    End If
End Function

'bestaetigungs frage
Public Function Ask(sString As String, Optional sTitle = " ") As Boolean
  
    If MsgBox(sString, vbYesNo, sTitle) = vbYes Then
        Ask = True
    Else
        Ask = False
    End If
End Function

'prints some info in debug window, with date
Public Sub Dbg(sInfo As String)
   Debug.Print Format(Now, "hh:mm:ss") & " - " & sInfo
End Sub

' Umlaute und Sonderzeichen verschlüsseln
'
' UsePlusRatherThanHexForSpace:
'  False:  Leerzeichen als %32 verschlüsseln
'  True :  Leerzeichen als + verschlüsseln
'http://www.vbarchiv.net/tipps/tipp_139-urlencode-und-urldecode.html  :: Matthias Zirngibl
Public Function UrlEncode(StringToEncode As String, Optional _
  UsePlusRatherThanHexForSpace As Boolean = False) As String
 
  Dim TempAns As String
  Dim CurChr As Integer
 
  CurChr = 1
  Do Until CurChr - 1 = Len(StringToEncode)
    Select Case Asc(Mid$(StringToEncode, CurChr, 1))
      Case 48 To 57, 65 To 90, 97 To 122
        TempAns = TempAns & Mid$(StringToEncode, CurChr, 1)
      Case 32
        If UsePlusRatherThanHexForSpace = True Then
          TempAns = TempAns & "+"
        Else
          TempAns = TempAns & "%" & Hex(32)
        End If
      Case Else
        TempAns = TempAns & "%" & Hex(Asc(Mid$(StringToEncode, _
          CurChr, 1)))
    End Select
    CurChr = CurChr + 1
  Loop
  UrlEncode = TempAns
End Function

' Entschlüsseln von mit URLEncode
' verschlüsselten Zeichenketten
'http://www.vbarchiv.net/tipps/tipp_139-urlencode-und-urldecode.html :: Matthias Zirngibl
Public Function UrlDecode(StringToDecode As String) As String
  Dim TempAns As String
  Dim CurChr As Integer
 
  CurChr = 1
  Do Until CurChr - 1 = Len(StringToDecode)
    Select Case Mid$(StringToDecode, CurChr, 1)
      Case "+"
        TempAns = TempAns & " "
      Case "%"
        TempAns = TempAns & Chr(Val("&h" & _
          Mid$(StringToDecode, CurChr + 1, 2)))
        CurChr = CurChr + 2
      Case Else
        TempAns = TempAns & Mid$(StringToDecode, CurChr, 1)
    End Select
    CurChr = CurChr + 1
  Loop
  UrlDecode = TempAns
End Function

'get path of string
Public Function PathGetDir(sPath As String)
    PathGetDir = Left$(sPath, InStrRev(sPath, "\"))
End Function

'get filename of sPath path
Public Function PathGetFileName(sPath As String)
    PathGetFileName = Mid$(sPath, InStrRev(sPath, "\") + 1)
End Function

'get extension of sPath path 'TODO "/path/pa.th/"
Public Function PathGetExtension(sPath As String)
    PathGetExtension = Mid$(sPath, InStrRev(sPath, "."))
End Function

'zeit als string oder Datum in SQL save string modifizieren
Public Function SQLString_formatTimeToSQLString(dDate As Variant) As String
    Dim saveDate As Date
    Dim testDate As Variant
    testDate = CdateTry(dDate, -1)
    
    If testDate <> -1 Then
        saveDate = CDate(testDate)
        SQLString_formatTimeToSQLString = Format(saveDate, "\#yyyy-mm-dd HH\:MM\:SS\#")
        Exit Function
    End If
    
    SQLString_formatTimeToSQLString = ""
End Function


'---------------------------------------------------------------------------------------
' Property: TempPath
'---------------------------------------------------------------------------------------
' Purpose: Get temp folder path
' Returns: String Temp folder
' Remarks: got from AccLib, https://svn.access-codelib.net/svn/codelib/trunk/file/modFiles.bas
' Remarks: modified by Dmitry Gorelenkov
Public Function TempPath() As String
    Const m_DEFAULT_TEMPPATH_NoEnv As String = "C:\"
    Const m_MAXPATHLEN As Long = 255
    
    Dim strTemp As String

    On Error GoTo HandleErr

    strTemp = Space$(m_MAXPATHLEN)
    APIGetTempPath m_MAXPATHLEN, strTemp
    strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
    If Len(strTemp) = 0 Then
        strTemp = m_DEFAULT_TEMPPATH_NoEnv
    End If
    TempPath = strTemp
    
    Exit Function
HandleErr:
    Debug.Print "TempPath error"
    TempPath = ""
End Function


'---------------------------------------------------------------------------------------
' Function: TempFileName
'---------------------------------------------------------------------------------------
' Purpose: get temp file path (file created)
' Param  : String sPrefix prefix for filename
' Param  : String sExtension extension for filename
' Returns: String temp file fullpath
' Remarks: https://svn.access-codelib.net/svn/codelib/trunk/file/modFiles.bas, modified
'---------------------------------------------------------------------------------------
Public Function TempFileName( _
  Optional sPrefix As String = "VBA", _
  Optional sExtension As String = "") As String
  Dim sTmpPath As String * 512
  Dim sTmpName As String * 576
  Dim nRet As Long
  Dim f As String
  nRet = APIGetTempPath(512, sTmpPath)
  If (nRet > 0 And nRet < 512) Then
    nRet = APIGetTempFileName(sTmpPath, sPrefix, 0, sTmpName)
    If nRet <> 0 Then f = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
    If sExtension > "" Then
      Kill f
      If Right(f, 4) = ".tmp" Then f = Left(f, Len(f) - 4)
      f = f & sExtension
    End If
    TempFileName = f
  End If
End Function


'---------------------------------------------------------------------------------------
' Function: CreateDir
'---------------------------------------------------------------------------------------
' Purpose: creates dir and all subdirs
' Param  : String sPath path of dir to create
' Returns: Boolean true if dir created, false otherwise
' Remarks: got from AccLib, https://svn.access-codelib.net/svn/codelib/trunk/file/modFiles.bas, modified
'---------------------------------------------------------------------------------------
Public Function CreateDirectory(ByVal sPath As String) As Boolean
   Dim strPathBefore As String

   If Right$(sPath, 1) = "\" Then
      sPath = Left$(sPath, Len(sPath) - 1)
   End If

   If Len(Dir$(sPath, vbDirectory)) > 0 Then 'Directory already created
      CreateDirectory = False
      Dir$ "\" ' to remove pointer
      Exit Function
   End If

   strPathBefore = Mid$(sPath, 1, InStrRev(sPath, "\") - 1)
   If Len(Dir$(strPathBefore, vbDirectory)) = 0 Then
      If CreateDirectory(strPathBefore) = False Then
         CreateDirectory = False
         Exit Function
      End If
   End If

   MkDir sPath

   CreateDirectory = True

End Function


'filters all letters except numeric
Public Function StrFilterNumeric(sString As String) As String
    Dim letterPos As Long
    Dim retString As String
    Dim letter As String
    
    For letterPos = 1 To Len(sString)
        letter = Mid(sString, letterPos, 1)
        If IsNumeric(letter) Then
            retString = retString & letter
        End If
    Next letterPos
    
    StrFilterNumeric = retString
End Function

'---------------------------------------------------------------------------------------
' Function: StrReplaceAll
'---------------------------------------------------------------------------------------
' Purpose: replaces all chars of array from the given string
' Param  : String sString source string
' Param  : Variant aChars array with not allowed chars
' Param  : String vReplacement (optional) char to replace, default "", or array with chars to replace
' Returns: String new string with replaced chars
' Remarks: in case vReplacement is an Array, length must be the same as aChars
'---------------------------------------------------------------------------------------
Public Function StrReplaceAll(ByVal sString As String, aChars As Variant, Optional vReplacement As Variant = "") As String
    Dim sToReplace As String
    
    'for single value
    If Not IsArray(vReplacement) Then
        On Error Resume Next
        sToReplace = CStr(vReplacement)
        
        'if Cstr failed.
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise vbObjectError, "clsTools:StrReplaceAll", "Wrong argument type vReplacement"
        End If
        
        On Error GoTo 0
        
        StrReplaceAll = StrReplaceAllByOne(sString, aChars, sToReplace)
        
    Else
        'for array replacement
        StrReplaceAll = StrReplaceAllByArray(sString, aChars, vReplacement)
    End If
End Function

'---------------------------------------------------------------------------------------
' Function: StrReplaceAllByOne
'---------------------------------------------------------------------------------------
' Purpose: replace all chars of array from the given string
' Param  : String sString source string
' Param  : Variant aChars array with not allowed chars
' Param  : String sReplacement (optional) char to replace, default ""
' Returns: String new string with replaced chars
'---------------------------------------------------------------------------------------
Private Function StrReplaceAllByOne(ByVal sString As String, aChars As Variant, Optional sReplacement As String = "") As String
    Dim sChar As Variant 'String
    Dim sResult As String
    sResult = sString
    
    For Each sChar In aChars
        sResult = Replace(sResult, sChar, sReplacement)
    Next
    
    StrReplaceAllByOne = sResult
End Function

'---------------------------------------------------------------------------------------
' Function: StrReplaceAllByArray
'---------------------------------------------------------------------------------------
' Purpose: replace all chars of array from the given string with all chars (compared by index) of replacement array
' Param  : String sString source string
' Param  : Variant aChars array with not allowed chars
' Param  : Variant aReplacement array with chars to replace
' Returns: String new string with replaced chars
'---------------------------------------------------------------------------------------
Private Function StrReplaceAllByArray(ByVal sString As String, aChars As Variant, aReplacement As Variant) As String
    Dim sResult As String
    Dim i As Long
    If Not IsArray(aReplacement) Then
        Err.Raise vbObjectError, "clsTools:StrReplaceAllByArray", "arg aReplacement is not an array"
    End If
    If (ArrayLength(aChars) <> ArrayLength(aReplacement)) Then
        Err.Raise vbObjectError, "clsTools:StrReplaceAllByArray", "array lengths are not equal, aChars: " & ArrayLength(aChars) & _
                                                                    "aReplacement" & ArrayLength(aReplacement)
    End If
    
    sResult = sString
    For i = LBound(aReplacement) To UBound(aReplacement)
        sResult = Replace(sResult, aChars(i), aReplacement(i))
    Next
    
    StrReplaceAllByArray = sResult
End Function


'---------------------------------------------------------------------------------------
' Function: StrUnescapeUTF8
'---------------------------------------------------------------------------------------
' Purpose: unescape utf8 escaped string
' Param  : string StringToDecode string to unescape
' Returns: String unescaped string
' Remarks: not fully tested, got from http://www.xtremevbtalk.com/showthread.php?t=152882
'---------------------------------------------------------------------------------------
Public Function StrUnescapeUTF8(ByVal StringToDecode As String) As String
    Dim i As Long
    Dim acode As Integer, sTmp As String
    
    On Error Resume Next
    
    If InStr(1, StringToDecode, "\") = 0 And InStr(1, StringToDecode, "%") = 0 Then
        StrUnescapeUTF8 = StringToDecode
        Exit Function
    End If
    For i = Len(StringToDecode) To 1 Step -1
        acode = Asc(Mid$(StringToDecode, i, 1))
        Select Case acode
        Case 48 To 57, 65 To 90, 97 To 122
            ' don't touch alphanumeric chars
        Case 92, 37: ' Decode \ or % value with uXXXX format
            If Mid$(StringToDecode, i + 1, 1) = "u" Then
                sTmp = CStr(CLng("&H" & Mid$(StringToDecode, i + 2, 4)))
                If IsNumeric(sTmp) Then
                    StringToDecode = Left$(StringToDecode, i - 1) & ChrW$(CInt("&H" & Mid$(StringToDecode, i + 2, 4))) & Mid$(StringToDecode, i + 6)
                End If
            End If
            
        Case 37: ' % not %uXXXX but %XX format
            
            sTmp = CStr(CLng("&H" & Mid$(StringToDecode, i + 1, 2)))
            If IsNumeric(sTmp) Then
                StringToDecode = Left$(StringToDecode, i - 1) & ChrW$(CInt("&H" & Mid$(StringToDecode, i + 1, 2))) & Mid$(StringToDecode, i + 3)
            End If
            
        End Select
    Next

    StrUnescapeUTF8 = StringToDecode
End Function



'---------------------------------------------------------------------------------------
' Function: FileGetContentSimple
'---------------------------------------------------------------------------------------
' Purpose: get file content, simple method
' Param  : String sName file path
' Returns: String file content
'---------------------------------------------------------------------------------------
Public Function FileGetContentSimple(sName As String) As String
    Dim intUnit As Integer
     
    On Error GoTo ErrGetFileContent
    intUnit = FreeFile
    Open sName For Input As intUnit
    FileGetContentSimple = Input(LOF(intUnit), intUnit)
ErrGetFileContent:
    Close intUnit
    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Function: FileContent
'---------------------------------------------------------------------------------------
' Purpose: Returns file content with set charset
' Param  : String sPath path of file
' Param  : String sCharset ["UTF-8"] charset, default UTF-8
' Returns: String content of file
' Remarks: http://cboden.de/index.php/softwareentwicklung/vba/tipps-tricks/19-utf8-einlesen
'---------------------------------------------------------------------------------------
Public Function FileGetContent(sPath As String, Optional sCharset As String = "UTF-8") As String
On Error GoTo ErrExit
    Dim adoStream As Object 'ADODB.Stream
        
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = sCharset
    adoStream.Open
    adoStream.LoadFromFile sPath
    FileGetContent = adoStream.ReadText
ErrExit:
    adoStream.Close
    Set adoStream = Nothing
End Function



'---------------------------------------------------------------------------------------
' Function: RegEx_getStringMatched
'---------------------------------------------------------------------------------------
' Purpose: 'returns string that is matched by sSearchRegex in sContent text
' Param  : String sSearchRegex regex for search. MUST contains 1 expression in brackets (searching for this submatch)
' Param  : String sContent text where to search
' Param  : Long subMatchNr [0] optional number of submatch
' Returns: String founded text
' Remarks: todo error handling?
'---------------------------------------------------------------------------------------
Public Function RegEx_getStringMatched(sPattern As String, sContent As String, Optional subMatchNr As Long = 0) As String
    Static objRegex As Object
    Dim aMatch As Object
    On Error GoTo Err
    
    If objRegex Is Nothing Then
        Set objRegex = CreateObject("vbscript.regexp")
    End If
    
    With objRegex
        .Multiline = True
        .IgnoreCase = True
        .Pattern = sPattern
        If .Test(sContent) Then
            Set aMatch = .Execute(sContent)
            RegEx_getStringMatched = aMatch(0).SubMatches(subMatchNr)
        Else
            RegEx_getStringMatched = ""
        End If
    End With

    Exit Function
Err:
    RegEx_getStringMatched = ""
End Function


'return array of strings, with file paths
'bOnlyNames if you need names of files only
Public Function DirList(ByVal sFolder As String, _
                    Optional sFilePattern As String = "*.*", _
                    Optional faFileAttr As VbFileAttribute = vbNormal, _
                    Optional bOnlyNames = False) As String()
    Dim sNewFile As String
    Dim retArray() As String

    
    
    If Right(sFolder, 1) <> "\" Then
        sFolder = sFolder & "\"
    End If
    
    sNewFile = Dir(sFolder & sFilePattern, faFileAttr)
    
    While sNewFile <> ""
        If sNewFile <> "." And sNewFile <> ".." Then
            
            ' Mit bit-weisem Vergleich sicherstellen, daß Name1 ein
            ' Verzeichnis ist.
            If (GetAttr(sFolder & sNewFile) And faFileAttr) = faFileAttr Then
                If bOnlyNames Then
                    Call Me.ArrayPush(retArray, sNewFile)
                Else
                    Call Me.ArrayPush(retArray, sFolder & sNewFile)
                End If
            End If
        
            
        End If
        sNewFile = Dir()
    Wend
    
    DirList = retArray
End Function



'---------------------------------------------------------------------------------------
' Function: CSVStringToArrayOfArrays
'---------------------------------------------------------------------------------------
' Purpose: parse csv content into 2d array. array(rows, columns)
' Param  : String strInput csv content
' Param  : String RowDelimiter delimiter for rows
' Param  : String FieldDelimiter delimiter for values
' Param  : Boolean bSkipHeader true if need to skip header (first row)
' Returns: array of arrays (rows arrays)
' Remarks: got some code from http://stackoverflow.com/questions/12259595/load-csv-file-into-a-vba-array-rather-than-excel-sheet
'---------------------------------------------------------------------------------------
Public Function CSVStringToArrayOfArrays(ByRef sInput As String, _
                       Optional sRowDelimiter As String = vbCr, _
                       Optional sFieldDelimiter As String = vbTab, _
                       Optional bSkipHeader As Boolean = False, _
                       Optional bTrim As Boolean = True) As Variant

    On Error Resume Next

    Dim i As Long, j As Long

    Dim lBoundRows As Long
    Dim uBoundRows As Long
    Dim arrRow() As String
    Dim arrRows As Variant
    Dim arrResult As Variant
    
    arrResult = Array()
    
    arrRows = Split(sInput, sRowDelimiter)
    If ArrayLength(arrRows) = 0 Then
        CSVStringToArrayOfArrays = Array()
        Exit Function
    End If
    
    lBoundRows = LBound(arrRows)
    uBoundRows = UBound(arrRows)
    
    While VBA.LenB(arrRows(uBoundRows)) <= 0 And uBoundRows >= lBoundRows ' clip out empty last row: common artifact in files with a terminating row delimiter
        uBoundRows = uBoundRows - 1
    Wend
    
    For i = IIf(bSkipHeader, lBoundRows + 1, lBoundRows) To uBoundRows
        'row as array
        arrRow = Split(arrRows(i), sFieldDelimiter)
        
        'trim if needed
        If bTrim Then
            For j = LBound(arrRow) To UBound(arrRow)
                arrRow(j) = Trim(arrRow(j))
            Next j
        End If
        
        'save to total array
        Call ArrayPush(arrResult, arrRow)
    Next i

    CSVStringToArrayOfArrays = arrResult
End Function


'---------------------------------------------------------------------------------------
' Function: AssignValue
'---------------------------------------------------------------------------------------
' Purpose: assign value from variant. If vFROM is an oject, "set" will be used.
' Param  : Variant vTO element that gets new value
' Param  : Variant vFROM element which value need to be assigned
'---------------------------------------------------------------------------------------
Public Function AssignValue(ByRef vTO As Variant, ByRef vFROM As Variant) As Boolean
    On Error GoTo ErrorHandler

    If IsObject(vFROM) Then
        Set vTO = vFROM
    Else
        vTO = vFROM
    End If
    
    AssignValue = True

    Exit Function
ErrorHandler:
    AssignValue = False
    Err.Raise Err.Number, "clsTools.AssignValue: " & Err.Source, Err.Description
End Function


'---------------------------------------------------------------------------------------
' Function: GetDBNameFromConnString
'---------------------------------------------------------------------------------------
' Purpose: get the database name if the input connection string points to a SQL Server DB,
'          or the DB path if the connection string points to a Jet (Access) DB
' Param  : String connString connection string
' Returns: String db name or path
' Remarks: http://www.devx.com/vb2themax/Tip/19608 , refactored for VBA
'---------------------------------------------------------------------------------------
Public Function GetDBNameFromConnString(ByVal connString As String) As String
' Examples:
'   Debug.WriteLine(GetDatabaseName("Provider=Microsoft.Jet.OLEDB.4.0;Data
' Source=D:\MyDB.mdb;")) ' => D:\MyDB.mdb;
'   Debug.WriteLine(GetDatabaseName("Provider=SQLOLEDB.1;Data Source=.;User
' ID=sa;Password=;Initial Catalog=MyDB")) ' => MyDB
'   Debug.WriteLine(GetDatabaseName("server=(local);database=MyDB;uid=sa;pwd=;")
' ) ' => MyDB
    Dim lcConnString As String
    Dim endIndex As Integer
    
    lcConnString = LCase(connString)

    ' if this is a Jet database, find the index of the "data source" setting
    Dim startIndex As Integer
    If InStr(1, lcConnString, "jet.oledb", vbTextCompare) > -1 Then
        startIndex = InStr(1, lcConnString, "data source=", vbTextCompare)
        If startIndex > -1 Then startIndex = startIndex + 12
    Else
        ' if this is a SQL Server database, find the index of the "initial
        ' catalog" or "database" setting
        startIndex = InStr(1, lcConnString, "initial catalog=", vbTextCompare)
        If startIndex > -1 Then
            startIndex = startIndex + 16
        Else ' if the "Initial Catalog" setting is not found,
             '  try with "Database"
            startIndex = InStr(1, lcConnString, "database=", vbTextCompare)
            If startIndex > -1 Then startIndex = startIndex + 9
        End If
    End If

    ' if the "database", "data source" or "initial catalog" values are not
    ' found, return an empty string
    If startIndex = -1 Then GetDBNameFromConnString = "": Exit Function

    ' find where the database name/path ends
    endIndex = InStr(startIndex, lcConnString, ";", vbTextCompare)
    If endIndex = -1 Then endIndex = Len(lcConnString)

    ' return the substring with the database name/path
    GetDBNameFromConnString = Mid$(connString, startIndex, endIndex - startIndex)
End Function


'---------------------------------------------------------------------------------------
' Function: CallProc
'---------------------------------------------------------------------------------------
' Purpose: call any sub or function of MODULE! (not class)
' Param  : Long ptrProc adress of  procedure. If params defined, must have 4 args!
' Param  : Variant vParam1 1. parameter for procedure etc..
' Returns: Variant pointer or return value of ptrProc
' Remarks: usage .CallProc(AddressOf procName, param1, param2 ..)
' Remarks: VERY DANGEROUS FUNCTION!! crash 90%, use in special well tested cases
'---------------------------------------------------------------------------------------
Public Function CallProc(ptrProc As Long, _
                        Optional vParam1 As Variant = 0, _
                        Optional vParam2 As Variant = 0, _
                        Optional vParam3 As Variant = 0, _
                        Optional vParam4 As Variant = 0) As Variant
    'TODO define what vars and use VarPtr, StrPtr, ObjPtr
    CallProc = APICallWindowProc(ptrProc, VarPtr(vParam1), VarPtr(vParam2), VarPtr(vParam3), VarPtr(vParam4))
End Function
