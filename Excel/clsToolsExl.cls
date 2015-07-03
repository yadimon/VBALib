VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToolsExl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsToolsExl
' Author : Dmitry Gorelenkov
' Date   : 05.2012
' Purpose: excel tools
' Version: 0.1
' Remarks: alpha version, many todos
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/Excel/clsToolsExl.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/clsTools.cls</use>
'  <ref><name>Excel</name><major>1</major><minor>7</minor><guid>{00020813-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Explicit

Private t As clsTools

'returns element from aElements with hightest priority.
'aPrioElements elements with aPriority priority mark.
'if aPriority is nothing, than last element of array has hightest prio, first lowest
Public Function GetProirityElement(ByVal aElements As Variant, ByVal aPrioElements As Variant, Optional ByVal aPriority As Variant)
    Dim Element As Variant
    Dim lIndex As Long, lBestPrioIndex As Long

    If IsEmpty(aElements) Or Not IsArray(aElements) Then Exit Function
    If IsEmpty(aPrioElements) Or Not IsArray(aPrioElements) Then Exit Function


    'falls priority array nicht uebergeben wurde, ein erstellen mit werten 0,1,2,3,...
    If IsMissing(aPriority) Then
        ReDim aPriority(LBound(aPrioElements) To UBound(aPrioElements))
        Dim i As Long
        For i = 0 To UBound(aPriority)
            aPriority(i) = i
        Next i
    End If

    'priority array muss mindestens gleiche groesse haben
    If t.ArrayLength(aPriority) < t.ArrayLength(aPrioElements) Then Exit Function

    lBestPrioIndex = -1
    'elements travisieren, hoechste prio finden
    For Each Element In aElements
        lIndex = t.ArrayIndexByValue(Element, aPrioElements)
        If lIndex = -1 Then GoTo nextElem
        If lBestPrioIndex > -1 Then
            If aPriority(lBestPrioIndex) < aPriority(lIndex) Then
                lBestPrioIndex = lIndex
            End If
        Else    'erste schrumpf
            lBestPrioIndex = lIndex
        End If
nextElem:
    Next Element


    If lBestPrioIndex = -1 Then
        GetProirityElement = ""
    Else
        GetProirityElement = aPrioElements(lBestPrioIndex)
    End If
End Function

'paste an 2d Array to specified range (1 cell)
Public Function ArrayPasteToRange(aValues As Variant, rngPasteRange As Excel.Range, Optional bSetHochkomma As Boolean = False)
    On Error GoTo Err:
    If IsArray(aValues) Then
        If UBound(aValues) - LBound(aValues) > 0 Then
            If rngPasteRange.Count = 1 Then
                Dim offsetRows As Long
                Dim offsetCols As Integer
                Dim i As Long
                Dim i2 As Long
                Dim lbound1 As Long
                Dim ubound1 As Long
                Dim lbound2 As Long
                Dim ubound2 As Long
                offsetRows = UBound(aValues) - LBound(aValues)
                offsetCols = UBound(aValues, 2) - LBound(aValues, 2)
                lbound1 = LBound(aValues, 1)
                ubound1 = UBound(aValues, 1)
                lbound2 = LBound(aValues, 2)
                ubound2 = UBound(aValues, 2)

                If bSetHochkomma Then
                    For i = lbound1 To ubound1
                        For i2 = lbound2 To ubound2
                            If VarType(aValues(i, i2)) = vbString Then
                                aValues(i, i2) = "'" & aValues(i, i2)
                            End If
                        Next i2
                    Next i
                End If
                rngPasteRange.Parent.Range(rngPasteRange, rngPasteRange.Offset(offsetRows, offsetCols)).Value = aValues
                ArrayPasteToRange = True
                Exit Function
            End If
        End If
    End If

Err:
    ArrayPasteToRange = False
End Function

'prueft ob sSheetName worksheet in wkbBook vorhanden ist
Public Function IsExistsWorksheet(sSheetName As String, wkbBook As Excel.Workbook) As Boolean
    Dim Sheet As Variant
    IsExistsWorksheet = False
    
    If wkbBook Is Nothing Then
        Set wkbBook = ThisWorkbook
    End If
    
    For Each Sheet In wkbBook.Sheets
        If VBA.StrComp(Sheet.Name, sSheetName, vbTextCompare) = 0 Then
            IsExistsWorksheet = True
            Exit Function
        End If
    Next Sheet
End Function

'wandelt eine Buchstabe z.B. "B" in eine SpaltenNr also z.B. "2" um
Public Function LetterToCol(ByVal sLetter As String) As Variant
    If IsNumeric(sLetter) Then
        LetterToCol = CInt(sLetter)
        Exit Function
    End If
    
    Dim num As Integer
    Dim letnum As Integer
    Dim i As Integer
    Dim currentLetter As String
    letnum = Len(sLetter)
    
    For i = 1 To letnum
        currentLetter = Mid(sLetter, i, 1)
        'falls nummer vorkommt, das input zuruckgeben
        If Not t.IsAlphaBetical(currentLetter) Or IsNumeric(currentLetter) Then
            LetterToCol = sLetter
            Exit Function
        End If
        num = num * 26 + Asc(UCase(currentLetter)) - 64
    Next i
    LetterToCol = num
End Function

'wandelt SpaltenNr in eine Buchstabe um
Public Function ColToLetter(ByVal lCol As Variant) As String
    If Not IsNumeric(lCol) Then
        ColToLetter = CStr(lCol)
        Exit Function
    End If
    
    Dim columnString
    Dim currentLetterNumber
        While (lCol > 0)
            currentLetterNumber = (lCol - 1) Mod 26
            columnString = Chr(currentLetterNumber + 65) & columnString
            lCol = (lCol - (currentLetterNumber + 1)) / 26
        Wend
    ColToLetter = columnString
End Function

'gibt array von werten zurueck die in eine Spalte "column" (nummer oder buchstabe) sich befinden.
'Ab zeile "fromFow" bis "toRow" in eine "sheet" tabelle
Public Function GetArrayOfColumn(ByVal vColumn As Variant, Optional fromRow As Variant = -1, Optional toRow As Variant = -1, Optional wksSheet As Excel.Worksheet)
    Dim aArray
    If wksSheet Is Nothing Then
        Set wksSheet = Excel.ActiveSheet
    End If

    If fromRow = -1 Then fromRow = GetFirstRow(wksSheet, LetterToCol(vColumn))    'sheet.UsedRange.row
    If toRow = -1 Then toRow = GetLastRow(wksSheet, LetterToCol(vColumn))    'sheet.UsedRange.Rows.Count

    If IsNumeric(vColumn) Then vColumn = ColToLetter(vColumn)

    If fromRow > toRow Then toRow = fromRow

    'falls erste und leer, dann leere array uebergeben
    If fromRow = toRow And wksSheet.Cells(fromRow, LetterToCol(vColumn)).Value = vbNullString Then
        GetArrayOfColumn = Array()
    Else
        aArray = wksSheet.Range(vColumn & fromRow & ":" & vColumn & toRow).Value2
        If Not IsArray(aArray) Then
            Dim tmpret(1 To 1, 1 To 1)
            tmpret(1, 1) = aArray
            aArray = tmpret
        End If
        GetArrayOfColumn = t.ArrayXDimTo1Dim(aArray)
    End If

    'getArrayOfColumn = t.array2DimTo1Dim(aArray)
End Function


'gibt array von werten zurueck die in eine Zeile "row" sich befinden.
'Ab spalte "fromCol" bis "toCol" in eine "sheet" tabelle
Public Function GetArrayOfRow(ByVal Row As Variant, Optional fromCol As Variant = -1, Optional toCol As Variant = -1, Optional Sheet As Excel.Worksheet)
    Dim aArray
    If Sheet Is Nothing Then
        Set Sheet = Excel.ActiveSheet
    End If


    fromCol = LetterToCol(fromCol)
    toCol = LetterToCol(toCol)

    'row und cols muessen nun numerisch sein
    If Not (IsNumeric(Row) And IsNumeric(fromCol) And IsNumeric(toCol)) Then Exit Function

    If fromCol = -1 Then fromCol = GetFirstCol(Sheet)
    If toCol = -1 Then toCol = GetLastCol(Sheet)
    With Sheet
        aArray = .Range(.Cells(Row, fromCol), .Cells(Row, toCol)).Value2
    End With

    GetArrayOfRow = t.ArrayChangeBounds(t.ArrayXDimTo1Dim(aArray), 1)
End Function

'fuellt ein Spalte mit werten von dem aArray, ab fromRow bis toRow
Public Function FillColumnByArray(ByVal Column As Variant, ByVal aArray As Variant, Optional fromRow As Variant = 1, Optional toRow As Variant = -1, Optional Sheet As Excel.Worksheet, Optional frmlaLocal = True)
'Dim aArrayOut
    If toRow = -1 Then toRow = fromRow + UBound(aArray) - LBound(aArray)
    If Sheet Is Nothing Then
        Set Sheet = Excel.ActiveSheet
    End If
    If IsNumeric(Column) Then Column = ColToLetter(Column)

    With Sheet
        If frmlaLocal Then
            .Range(Column & fromRow & ":" & Column & toRow).FormulaLocal = Excel.WorksheetFunction.Transpose(aArray)
        Else
            '.range(column & fromRow & ":" & column & toRow).FormulaArray = WorksheetFunction.Transpose(aArray)
            'maximum 255 zeichen bei FromulaArray .. deswegen ..
            Dim i As Long
            For i = fromRow To toRow
                .Cells(i, Column).Formula = aArray(i - fromRow)
            Next i
        End If
    End With

End Function

'no need? :(
Public Function ModifyValuesToPercentString(ByVal aArray)
    Dim i As Integer
    Dim newArr() As String
    ReDim newArr(LBound(aArray) To UBound(aArray))
    For i = LBound(aArray) To UBound(aArray)
        newArr(i) = CStr(Format(CStr(aArray(i)), "0.0%"))
    Next i

    ModifyValuesToPercentString = newArr
End Function

'letzte Spalte in bestimmter zeile "fromRow"
Public Function GetLastCol(Optional ByRef oTabelle As Excel.Worksheet, Optional fromRow As Variant = 1)
    If oTabelle Is Nothing Then
        Set oTabelle = Excel.ActiveSheet
    End If

    With oTabelle
        GetLastCol = .Cells(fromRow, .Columns.Count).End(xlToLeft).Column
    End With
End Function


'letzte Zeile in bestimmte Spalte "fromCol"
Public Function GetLastRow(Optional ByRef oTabelle As Excel.Worksheet, Optional fromCol As Variant = 1)
    If oTabelle Is Nothing Then
        Set oTabelle = Excel.ActiveSheet
    End If

    If Not IsNumeric(fromCol) Then fromCol = LetterToCol(fromCol)

    With oTabelle
        GetLastRow = .Cells(.Rows.Count, fromCol).End(xlUp).Row
    End With
End Function

'erste Zeile in bestimmte Spalte "fromCol" 'TODO bessere methode?
Public Function GetFirstRow(Optional ByRef oTabelle As Excel.Worksheet, Optional fromCol As Variant = 1)
    Dim i As Long
    If oTabelle Is Nothing Then
        Set oTabelle = Excel.ActiveSheet
    End If

    For i = 1 To GetLastRow(oTabelle, fromCol)
        If oTabelle.Cells(i, fromCol) <> vbNullString Then
            GetFirstRow = i
            Exit For
        End If
    Next i

    GetFirstRow = i
End Function

'erste Spalte in bestimmter zeile "fromRow"
Public Function GetFirstCol(Optional ByRef oTabelle As Excel.Worksheet, Optional fromRow As Long = 1)
    Dim i As Long
    If oTabelle Is Nothing Then
        Set oTabelle = Excel.ActiveSheet
    End If

    For i = 1 To GetLastCol(oTabelle, fromRow)
        If oTabelle.Cells(fromRow, i) <> "" Then
            GetFirstCol = i
            Exit For
        End If
    Next i

    GetFirstCol = i
End Function

'todo better own func ?
'modify Array of Array to 2dArray. Beginns with array(1,1)
Public Function ArrOfArrsToArray2d(aArray As Variant)
    Dim retArray As Variant

    If Not IsArray(aArray) Then
        ArrOfArrsToArray2d = aArray
        Exit Function
    End If

    If t.ArrayLength(aArray) = 0 Then
        ArrOfArrsToArray2d = Array()
        Exit Function
    End If

    retArray = Excel.WorksheetFunction.Index(aArray, 0, 0)

    ArrOfArrsToArray2d = retArray
End Function

'workbook finden, falls schon geoeffnet, oder oeffnen falls nicht geoeffnet
Public Function FindOpenedDocOrOpen(FilePath As String) As Excel.Workbook
    Dim wkbReturn As Excel.Workbook

    Set wkbReturn = FindOpenedDoc(FilePath)
    
    If wkbReturn Is Nothing Then
        Set FindOpenedDocOrOpen = Excel.Workbooks.Open(FilePath)
    Else
        Set FindOpenedDocOrOpen = wkbReturn
    End If
    
End Function

'sucht geoeffnete datei mit dem Pfad, oder mit dem Namen
Public Function FindOpenedDoc(FilePath As String, Optional bByName As Boolean = False, Optional bLikeName As Boolean = False) As Excel.Workbook
    Dim curWkb As Excel.Workbook
    If bByName Then
        Set FindOpenedDoc = FindOpenedDocByName(FilePath, bLikeName)
    Else
        For Each curWkb In Excel.Workbooks
            If StrComp(curWkb.FullName, FilePath, vbTextCompare) = 0 Then
                Set FindOpenedDoc = curWkb
                Exit For
            End If
        Next curWkb
    End If
End Function

'geoeffnetes Dok finden, dessen name ist gleich "sName" ODER falls bLikeName true ist, dessen name sName drin hat.
Public Function FindOpenedDocByName(sName As String, Optional bLikeName As Boolean = False) As Excel.Workbook
    Dim curWkb As Excel.Workbook
    For Each curWkb In Excel.Workbooks
        If curWkb.Name = sName Then
            Set FindOpenedDocByName = curWkb
            Exit For
        ElseIf bLikeName And InStr(1, curWkb.Name, sName) > 0 Then
            Set FindOpenedDocByName = curWkb
            Exit For
        End If
    Next curWkb
End Function

'screenUpdating und Formulasberechnung Ein/Ausschalten
Public Function SetScreenCalculationProtect(bOnOff As Boolean, Optional sProtPW As String = "0000")
'reinfolge beim einschalten, ausschalten 'no need? Oo
    If bOnOff Then
        If IsEmpty(sProtPW) Then
            Call ProtectWorkbook(bOnOff)
        Else
            Call ProtectWorkbook(bOnOff, sProtPW)
        End If
        Call ExcelCalculation(bOnOff)
        Call SetScreenUpdating(bOnOff)

    Else
        Call SetScreenUpdating(bOnOff)
        ExcelCalculation (bOnOff)
        If IsEmpty(sProtPW) Then
            Call ProtectWorkbook(bOnOff)
        Else
            Call ProtectWorkbook(bOnOff, sProtPW)
        End If
    End If

End Function

'screen und berechnung ausschalten. oefter benoetigt..
Public Function SetCalculationAndScreen(bOnOff As Boolean)
    Call SetScreenUpdating(bOnOff)
    Call ExcelCalculation(bOnOff)
End Function

'screenUpdating Ein/Ausschalten
Public Function SetScreenUpdating(bOnOff As Boolean)
    If (CBool(bOnOff)) Then
        Excel.Application.ScreenUpdating = True
    Else
        Excel.Application.ScreenUpdating = False
    End If
End Function

'Formulasberechnung Ein/Ausschalten
Public Function ExcelCalculation(bOnOff As Boolean)
    If Excel.Application.Workbooks.Count <= 0 Then Exit Function    ' falls nichts geoeffnet, exit

    If (CBool(bOnOff)) Then
        Excel.Application.Calculation = xlCalculationAutomatic
    Else
        Excel.Application.Calculation = xlCalculationManual
    End If
End Function

'open dialog to pick a file, return "" on error.
Public Function GetFilePathDialog(sFileFilter As String, sTitle As String) As String
'Pfad zur Datei bestimmen
'On Error Resume Next
    On Error GoTo ErrorHandler

    GetFilePathDialog = Excel.Application.GetOpenFilename(FileFilter:=sFileFilter, Title:=sTitle)
    'todo? international?
    If GetFilePathDialog = "Falsch" Then GetFilePathDialog = ""

    Exit Function
ErrorHandler:
    GetFilePathDialog = ""
End Function

'filePicker or FolderPicker dialog. 'todo miltiselect option ?
Public Function DialogGetPath(Optional bFileOrFolder As Boolean = True, Optional sTitle As String, Optional filterTyp As String, Optional filterExt As String, Optional sInitial As String, Optional bMulti As Boolean = False, Optional sButtonName As String)
    Dim fd As FileDialog
    Dim sInitDir As String
    sInitDir = CurDir ' Store initial directory
    'see http://www.mrexcel.com/forum/excel-questions/294728-browse-folder-visual-basic-applications-2.html
    'need to change dir to let it not be locked while excel running

    'if true -> file
    If bFileOrFolder Then
        Set fd = Excel.Application.FileDialog(msoFileDialogFilePicker)

        'falls filter werde vorhanden
        If t.IsSet(filterTyp) And t.IsSet(filterExt) Then
            fd.Filters.Clear
            fd.Filters.Add filterTyp, filterExt, 1
        End If
        
        'extra button title
        If t.IsSet(sButtonName) Then
            fd.ButtonName = sButtonName
        End If

        'folder
    Else
        Set fd = Excel.Application.FileDialog(msoFileDialogFolderPicker)
    End If

    'anfangspfad
    If t.IsSet(sInitial) Then fd.InitialFileName = sInitial

    'if multiselected
    If bMulti Then
        fd.AllowMultiSelect = True
    End If

    If t.IsSet(sTitle) Then
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
                Call t.ArrayPush(vrtResult, vrtSelectedItem)
            Else
                vrtResult = vrtSelectedItem
                Exit For    'zur sicherheit
            End If



        Next vrtSelectedItem
        'The user pressed Cancel.
    Else
        vrtResult = ""
    End If


    DialogGetPath = vrtResult
    
    
    ' Reset directory before exit, ignore pat errors
    On Error Resume Next
    ChDrive sInitDir ' Return to the Initial Drive
    ChDir sInitDir ' Resets directory for Initial Drive
    On Error GoTo 0
End Function

'Blattschutz aufheben, oder wieder einschalten
Public Sub ProtectWorkbook(bYesNo As Boolean, Optional sProtPW As String = "0000", Optional ByRef wkbWorkbook As Excel.Workbook)
    On Error Resume Next
    Dim i As Integer

    If wkbWorkbook Is Nothing Then Set wkbWorkbook = Excel.ThisWorkbook

    With wkbWorkbook
        If bYesNo Then
            For i = 1 To .Worksheets.Count
                With Worksheets(i)
                    .Protect Password:=sProtPW, UserInterfaceOnly:=True
                    '.EnableSelection = xlUnlockedCells
                    '.EnableSelection = xlNoRestrictions
                End With
            Next i
        Else
            For i = 1 To .Worksheets.Count
                With Excel.Worksheets(i)
                    .Unprotect Password:=sProtPW
                    '.EnableSelection = xlNoRestrictions
                End With
            Next i
        End If
    End With
End Sub

'recalculates all sheets in Workbook
Public Function RecalcAllSheets(Optional wkbWorkbook As Excel.Workbook)
    Dim wksSheet As Variant
    If wkbWorkbook Is Nothing Then Set wkbWorkbook = Excel.ThisWorkbook

    For Each wksSheet In wkbWorkbook.Sheets
        wksSheet.Calculate
    Next wksSheet
End Function

'calculate array of sheets objects
Public Function RefreshTables(ByRef aSheets As Excel.Sheets)
    Dim wksTable As Variant
    For Each wksTable In aSheets
        wksTable.Calculate
    Next wksTable
End Function

'worksheet hinzufuegen, am ende. TODO : index wohin.
Public Function AddWorksheet(Optional sWksName As String, Optional wkbWorkbook As Excel.Workbook)
    If wkbWorkbook Is Nothing Then
        Set wkbWorkbook = Excel.ThisWorkbook
    End If

    If IsExistsWorksheet(sWksName, wkbWorkbook) Then
        AddWorksheet = Null
        Exit Function
    Else
        Set AddWorksheet = wkbWorkbook.Worksheets.Add(After:=wkbWorkbook.Worksheets(wkbWorkbook.Worksheets.Count))
        If sWksName <> vbNullString Then
            AddWorksheet.Name = sWksName
        End If
    End If

End Function


'array von worksheets ausdrucken, worksheets mussen in einer arbeitsmappe sein!!
'bEinzeln - einzeln ausdrucken
'bFirstPrint - drucker auswahlfenster ja/nein
'bOrdered ob alle sheets in richtige reihenfolge gedruckt werden sollen
Public Function PrintSheets(aWorksheets As Variant, Optional bEinzeln As Boolean = False, _
                            Optional bFirstPrint As Boolean = True, Optional bOrdered As Boolean = True, Optional wkbWorkbook As Excel.Workbook)
    On Error GoTo PrintSheets_Error

    If Not IsArray(aWorksheets) Then Exit Function
    Dim wksFromArray As Variant
    Dim bFirstSelect As Boolean
    Dim bDisplayPageBreaksB4 As Boolean
    If wkbWorkbook Is Nothing Then Set wkbWorkbook = Excel.ThisWorkbook
    Dim aReihnfolgeB4 As Variant    'array of worksheets refs
    Dim xlApp As Excel.Application
    
    bDisplayPageBreaksB4 = aWorksheets(0).DisplayPageBreaks    'todo save for each page?

    'case must be ordered, save original order, and sort by new array
    If bOrdered Then
        aReihnfolgeB4 = GetWksArray(wkbWorkbook)
        Call WorksheetsOrder(aWorksheets)
    End If

    If Not bEinzeln Then    'print all merged

        bFirstSelect = True    'to prevent print of selections b4
        For Each wksFromArray In aWorksheets
            wksFromArray.Select bFirstSelect
            bFirstSelect = False
        Next wksFromArray

        If bFirstPrint Then
            xlApp.Dialogs(xlDialogPrint).Show
        Else
            ActiveWindow.SelectedSheets.PrintOut
        End If

    Else    'print separate
        For Each wksFromArray In aWorksheets
            wksFromArray.Select

            If bFirstPrint Then
                xlApp.Dialogs(xlDialogPrint).Show
            Else
                ActiveWindow.SelectedSheets.PrintOut
            End If

            bFirstPrint = False
        Next wksFromArray
    End If

    PrintSheets = True

Final:
    'case must be ordered, order original
    If bOrdered Then
        Call WorksheetsOrder(aReihnfolgeB4)
    End If

    'pagebreaks, hide?
    If IsArray(aWorksheets) Then
        For Each wksFromArray In aWorksheets
            wksFromArray.DisplayPageBreaks = bDisplayPageBreaksB4
        Next wksFromArray
    End If

    Exit Function

PrintSheets_Error:
    PrintSheets = False
    Debug.Print "PrintSheets - ERROR"
    GoTo Final
End Function

'returns array of refs of worksheets in the wkbWorkbook
Public Function GetWksArray(Optional wkbWorkbook As Excel.Workbook)
    Dim wksSheet As Excel.Worksheet
    Dim aReturnArray As Variant
    If wkbWorkbook Is Nothing Then Set wkbWorkbook = Excel.ThisWorkbook

    For Each wksSheet In wkbWorkbook.Worksheets
        Call t.ArrayPush(aReturnArray, wksSheet)
    Next wksSheet

    GetWksArray = aReturnArray
End Function

'orders worksheets as order of array
Public Function WorksheetsOrder(orderedArrayOfWks As Variant)
    Dim i As Integer
    Dim iArrLbnd As Integer
    Dim iArrUbnd As Integer

    WorksheetsOrder = False
    If Not IsArray(orderedArrayOfWks) Then Exit Function
    If t.ArrayLength(orderedArrayOfWks) <= 1 Then Exit Function    'no need to sort 1 element

    iArrLbnd = LBound(orderedArrayOfWks)
    iArrUbnd = UBound(orderedArrayOfWks)

    For i = iArrLbnd + 1 To iArrUbnd
        orderedArrayOfWks(i).Move After:=orderedArrayOfWks(i - 1)
    Next i

    WorksheetsOrder = True
End Function

Private Sub Class_Initialize()
    Set t = New clsTools
End Sub

Private Sub Class_Terminate()
    Set t = Nothing
End Sub
