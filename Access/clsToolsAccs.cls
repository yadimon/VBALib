VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToolsAccs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsToolsAccs
' Author : Dmitry Gorelenkov
' Date   : 16.05.2014
' Purpose: Access tools
' Version: 0.9
' Remarks: many not used, bad, not tested methods here, TODO cleanup
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/Access/clsToolsAccs.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/clsTools.cls</use>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'  <ref><name>Excel</name><major>1</major><minor>7</minor><guid>{00020813-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'

Option Compare Database
Option Explicit

Private t As clsTools
Private Const LONGMAX As Long = 2 ^ 31 - 1

Private Sub Class_Initialize()
    Set t = New clsTools
End Sub

Private Sub Class_Terminate()
    Set t = Nothing
End Sub

'sucht nach einem Tag, der mit komma getrennt in "Marke" drin ist
Public Function GetCtrlsByTag(ByRef frmForm As Form, ByVal sTag As String) As Scripting.Dictionary

    
    On Error GoTo getCtrlsByTag_Error

    If frmForm Is Nothing Then Exit Function
    
    Dim ctrControl As Object
    Dim aArrTags As Variant
    Dim retArrayCol As New Scripting.Dictionary
    
    For Each ctrControl In frmForm.Controls
        aArrTags = Split(ctrControl.Tag, ",", , vbTextCompare)
        If t.ArrayContains(sTag, aArrTags, True, True, True) Or sTag = vbNullString Then
            Call retArrayCol.Add(ctrControl, ctrControl.Name)
        End If
        
    Next ctrControl
    
    Set GetCtrlsByTag = retArrayCol
    

Finally:
    On Error Resume Next
    Set GetCtrlsByTag = Nothing
    On Error GoTo 0
    Exit Function

getCtrlsByTag_Error:

    Err.Raise vbObjectError, "getCtrlsByTag of Class Module clsToolsAccs"
    Resume Finally

End Function



'add tooltip mit steuerelement-inhalten
Sub AddToolTips(frmAktForm As Form)
    Dim ctlInForm As Control
    On Error Resume Next
    
    For Each ctlInForm In frmAktForm.Controls
      With ctlInForm
        If .ControlType = acTextBox Then
          .ControlTipText = IIf(IsNull(.Value), "", .Value)
        ElseIf .ControlType = acComboBox Or _
          .ControlType = acListBox Then
          If .ColumnCount > 1 Then
            .ControlTipText = IIf(IsNull(.Value), "", .Column(.BoundColumn))
          Else
            .ControlTipText = IIf(IsNull(.Value), "", .Value)
          End If
        End If
      End With
    Next ctlInForm

End Sub

'sets properties back to default
'frmForm are controls (but a Form will be as argument)
Public Function SetDefaultFormValues(ByRef frmForm As Variant)
    Dim ctl As Control
    On Error Resume Next
    For Each ctl In frmForm
        Select Case ctl.ControlType
            Case acComboBox, acListBox, acOptionGroup, acTextBox, acCheckBox    ' add other types if necessary
                If Not ctl.Locked Then
                    ctl.Value = QuotTrim(ctl.DefaultValue)
                End If
        End Select
    Next
End Function

'removes quotation at beginn and end of the string, if exists
Private Function QuotTrim(ByVal vValue As Variant)

    If VarType(vValue) = vbString Then
        If Left(vValue, 1) = """" Then
            vValue = Right(vValue, Len(vValue) - 1)
        End If
        
        If Right(vValue, 1) = """" Then
            vValue = Left(vValue, Len(vValue) - 1)
        End If
    End If
    
    QuotTrim = vValue
    
End Function


Public Function ModifyAllPossible(ByRef frmForm As Variant, sWhatToModify As String, Optional bWhatValue As Boolean = True)
    Dim ctl As Variant
    On Error Resume Next
    
    For Each ctl In frmForm
        Select Case LCase(sWhatToModify)
            Case "value"
                ctl.Value = ""

            Case "controlsource"
                ctl.ControlSource = ""
           
            Case "locked"
                ctl.Locked = bWhatValue
            
            Case "enabled"
                ctl.Enabled = bWhatValue
                
            Case "visible"
                ctl.Visible = bWhatValue
        
        End Select

        
    Next
End Function

'select 1 listeneintrag mit index
Public Function SetListIndex(ByRef ctlList As Variant, Index As Integer) As Boolean
    SetListIndex = False
    
    If TypeName(ctlList) <> "ComboBox" And TypeName(ctlList) <> "ListBox" Then
        Debug.Assert False
        Exit Function
    End If
    
    With ctlList
        'wenn genug eintraege fuer den gesetzten wert da sind
        If .ListCount > Index Then
            'falls multiselect bei listbox
            If TypeName(ctlList) = "ListBox" Then
                If .MultiSelect = 1 Then
                    .Selected(Index) = True
                    SetListIndex = True
                    Exit Function
                End If
            End If
            
            'sonst mit ItemData versuchen
            .Value = .ItemData(Index)
            SetListIndex = True
        End If
    End With
End Function

'list eintrag markieren bei Wert sValue, in spalte iColumn
Public Function SelectListByValue(ByRef lstList As Access.ListBox, sValue As String, Optional iColumn As Integer = -1) As Boolean
    Dim i As Integer
    Dim tmpValue As Variant
    SelectListByValue = False
    
    'fuer jeden eintrag
    For i = 0 To lstList.ListCount
        'falls keine Spalte eingegeben, gebundete Spalte lesen
        If iColumn = -1 Then
            tmpValue = lstList.ItemData(i)
        ElseIf iColumn < lstList.ColumnCount - 1 Then
            tmpValue = lstList.Column(iColumn, i)
        End If
        'falls uebereinstimmt, markieren
        If StrComp(CStr(tmpValue), sValue, vbTextCompare) = 0 Then
            SelectListByValue = SetListIndex(lstList, i)
            Exit For
        End If
    Next i
End Function

'mehrere zeilen in ListBox selektieren, bei Werten in aValues array
Public Function SelectListByValues(ByRef lstList As Access.ListBox, aValues As Variant, Optional iColumn As Integer = -1)
    Dim bReturnResult As Boolean
    bReturnResult = False
    If Not IsArray(aValues) Then
        Debug.Assert False
        Exit Function
    End If
    
    If lstList.MultiSelect <> 1 Then
        Debug.Assert False
        Exit Function
    End If
    
    Dim vValue As Variant
    bReturnResult = True
    For Each vValue In aValues
        bReturnResult = bReturnResult & SelectListByValue(lstList, CStr(vValue), iColumn)
    Next vValue
    
    SelectListByValues = bReturnResult
End Function

'deselect all items in ListBox.
'http://allenbrowne.com/func-12.html  modified
Function UnselectListBox(lstList As Access.ListBox) As Boolean
On Error GoTo Err_ClearList
    'Purpose:   Unselect all items in the listbox.
    'Return:    True if successful
    'Author:    Allen Browne. http://allenbrowne.com  June, 2006.
    Dim varItem As Variant

    If lstList.MultiSelect = 0 Then
        lstList.Value = Null
    Else
        For Each varItem In lstList.ItemsSelected
            lstList.Selected(varItem) = False
        Next
    End If

    UnselectListBox = True

    Exit Function

Err_ClearList:
    Debug.Print "error by clsToolsAccs.unselectListBox"
End Function

Public Function GetSelectedFromListBox(ByRef lstList As Access.ListBox, Optional lColumn As Long = 0) As Variant
    Dim lSelected As Variant
    Dim vValue As Variant
    Dim aResultArray As Variant
    aResultArray = Array()
    
    If lstList.MultiSelect = 0 Then
        Call t.ArrayPush(aResultArray, lstList.Value)
    Else
        For Each lSelected In lstList.ItemsSelected
            vValue = CStr(lstList.Column(lColumn, lSelected))
            Call t.ArrayPush(aResultArray, vValue)
        Next
    End If
    
    GetSelectedFromListBox = aResultArray
    
End Function

'try to scale formular
Public Function ResizeFormular(ByRef frmForm As Variant, dFaktor As Double)
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In frmForm
        If Not ctrl Is Nothing Then
            ctrl.Height = ctrl.Height * dFaktor
            ctrl.Width = ctrl.Width * dFaktor
            ctrl.Top = ctrl.Top * dFaktor
            ctrl.Left = ctrl.Left * dFaktor
            ctrl.FontSize = ctrl.FontSize * dFaktor
            
            'for subform add recursive scale
            If ctrl.ControlType = acSubform Then
                Call ResizeFormular(ctrl, dFaktor)
            End If
        End If
    Next ctrl

End Function

'datenbank pfad fuer alle (bzw mit bestimmten prefix) verknuepfte tabellen aendern
Public Function ChangeSourceDB(sPath As String, Optional sTblPrefix As String = vbNullString) As Integer
    Dim DB As DAO.Database
    Dim sPrefLen As String
    Dim i As Integer
    

    
    On Error GoTo changeSourceDB_Error

    Set DB = CurrentDb()
    ChangeSourceDB = 0
    
    
    
    sPrefLen = Len(sTblPrefix)
    
    For i = 0 To DB.TableDefs.Count - 1
        If DB.TableDefs(i).Attributes = dbAttachedTable Then '.connect <> ""
            If Left(DB.TableDefs(i).Name, sPrefLen) = sTblPrefix Then
                If Mid(DB.TableDefs(i).Connect, 11) <> sPath Then
                    DB.TableDefs(i).Connect = ";database=" & sPath
                    DB.TableDefs(i).RefreshLink
                    ChangeSourceDB = 1 'set changed
                End If
            End If
        End If
    Next i
    
    Exit Function

changeSourceDB_Error:

    Err.Raise vbObjectError, "changeSourceDB of Class Module clsToolsAccs"
    ChangeSourceDB = -1

End Function

'select datenbank dialog
'sDefaultPath vordefiniertes Pfad
Public Function SelectDBDialog(sPrefTbl As String, Optional sDefaultPath As String) As String
    Dim sDBPath As String
    
    If Not t.FileExists(sDefaultPath) Then
        sDefaultPath = ""
    End If
    
    sDBPath = t.DialogGetPath(True, "BackEnd Datenbank Pfad", "Access DB Files", "*.accd*", sDefaultPath, , "Auswählen")
    If Not t.IsSet(sDBPath) Then Exit Function
    
    If ChangeSourceDB(sDBPath, sPrefTbl) <> -1 Then
        MsgBox "Erfolgreich geändert"
    Else
        MsgBox "Fehler aufgetreten"
    End If
    
    SelectDBDialog = sDBPath
End Function

'get DBpath of all attached tables
Public Function GetBackEndPath(Optional dbDatabase As DAO.Database) As String
    Dim DB As DAO.Database
    Dim i As Integer
    Dim sFoundPath As String


    On Error GoTo getBackEndPath_Error

    If dbDatabase Is Nothing Then Set DB = CurrentDb()

    For i = 0 To DB.TableDefs.Count - 1
        If DB.TableDefs(i).Attributes = dbAttachedTable Then
            If sFoundPath = "" Then
                sFoundPath = Mid(DB.TableDefs(i).Connect, 11)
            End If
            'falls andere path gefunden, error, return 0
            If Mid(DB.TableDefs(i).Connect, 11) <> sFoundPath Then
                GetBackEndPath = "0"
            End If
        End If
    Next i

    GetBackEndPath = sFoundPath

    Exit Function

getBackEndPath_Error:

    Err.Raise vbObjectError, "getBackEndPath of Class Module clsToolsAccs"
    GetBackEndPath = ""

End Function

'prueft alle verknuepfte mit BackEnd Tabellen, ob die BackEnd DB zugreifbar ist
'sehr langsam!!!
Public Function CheckDBTblLinks(Optional dbDatabase As DAO.Database) As Boolean
    Dim DB As DAO.Database
    Dim i As Integer
    
    On Error GoTo checkTblLinks_Error

    If dbDatabase Is Nothing Then Set DB = CurrentDb()
    
    For i = 0 To DB.TableDefs.Count - 1
        If DB.TableDefs(i).Attributes = dbAttachedTable Then
            DB.TableDefs(i).RefreshLink
        End If
    Next i
    

    CheckDBTblLinks = True
    
    Exit Function

checkTblLinks_Error:
    If Err.Number = 3024 Then 'cannot refreshlink
        CheckDBTblLinks = False
    Else
        Err.Raise vbObjectError, "checkTblLinks of Class Module clsToolsAccs"
    End If
End Function

'prueft ob eine tabelle existiert
'from http://www.office-loesung.de/ftopic1371_0_0_asc.php
Public Function TableExists(sTableName As String, Optional DB As DAO.Database) As Boolean
    On Error GoTo Fehlerbehandlung
    Dim td As DAO.TableDef
    If DB Is Nothing Then Set DB = CurrentDb
    
    Set td = DB.TableDefs(sTableName)
    TableExists = True
    Exit Function
Fehlerbehandlung:
    TableExists = False
End Function

'prueft ob in tabelle sTable sColName spalte existiert
'http://www.office-loesung.de/ftopic24463_0_0_asc.php  - modified
Public Function ColumnExists(sTable As String, sColName As String)
    Dim tdf As DAO.TableDef
    Dim fldField As DAO.Field
    Dim DB As DAO.Database
    Set DB = CurrentDb
    
    
    'Fehlerbehandlung abschalten
    On Error Resume Next
    
    Set tdf = DB.TableDefs(sTable)

    
    ColumnExists = True
    
    Set fldField = tdf.Fields(sColName)
    
    If Err <> 0 Then
        ColumnExists = False
    End If
    
    On Error GoTo 0
End Function


'Recordset nach Excel exportieren
'http://www.ms-office-forum.net/forum/showthread.php?t=242723  - modified
Public Function Recordset2Excel(ByRef rstData As Object, _
                                Optional ByVal StartRowNr As Long = 1, _
                                Optional ByVal StartColNr As Long = 1, _
                                Optional ByVal bWithRecordsetHeaders As Boolean = True) As Boolean

'rstData = ADODB oder DAO-Recordset-Referenz

    'tests
    If StartRowNr < 1 Or StartColNr < 1 Or _
       (LCase(TypeName(rstData)) <> LCase("Recordset") And LCase(TypeName(rstData)) <> LCase("Recordset2")) Then

        t.Dbg "Recordset2Excel: falsche parameter"
        Exit Function
    End If


    Dim xlApp As Object    '.. = Excel.Application
    Dim xlBook As Object   '.. = Excel.Workbook
    Dim xlSheet As Object  '.. = Excel.WorkSheet
    Dim rngRange As Object '.. = Excel.Range

    On Error GoTo Recordset2Excel_Error

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    xlApp.Visible = True
    Set xlSheet = xlBook.Worksheets(1)
    Set rngRange = xlSheet.Cells(StartRowNr, StartColNr)
    Call Recordset2ExcelRange(rstData, rngRange, bWithRecordsetHeaders)
    
    Recordset2Excel = True
    Exit Function

Recordset2Excel_Error:
    Err.Raise vbObjectError, "Recordset2Excel of Klassenmodul clsToolsAccs"


End Function

'---------------------------------------------------------------------------------------
' Function: Recordset2ExcelRange
'---------------------------------------------------------------------------------------
' Purpose: puts data from recordset to range
' Param  : Recordset rs recordset to put
' Param  : Excel.Range rngRange start cell range of excel worksheet where rs will be put
' Param  : Boolean bWithRecordsetHeaders (optional) true if need the headers too. Deafult true
' Remarks: based on http://www.ms-office-forum.net/forum/showthread.php?t=242723
'---------------------------------------------------------------------------------------
Public Function Recordset2ExcelRange(rstData As Object, rngRange As Object, Optional ByVal bWithRecordsetHeaders As Boolean = True)
    Dim lngRstDataCols As Long
    Dim strRstDataColsName() As String
    Dim i As Long
    
    'headers
    If bWithRecordsetHeaders Then
        With rstData
            lngRstDataCols = .Fields.Count
            ReDim strRstDataColsName(lngRstDataCols - 1)
            
            'fill header strings array
            For i = 0 To (lngRstDataCols - 1)
                strRstDataColsName(i) = .Fields(i).Name
            Next i
            
            'put in excel worksheet
            rngRange.Worksheet.Range(rngRange, rngRange.Offset(0, lngRstDataCols)).Value = strRstDataColsName
            
            'offset 1 row, for the data
            Set rngRange = rngRange.Offset(1, 0)
        End With
    End If
    
    'copy from recordset
    rngRange.CopyFromRecordset rstData
    
End Function


'aktuelle datensatz nach bestimmte bedingungen umstellen/suchen
Sub FormMoveToRecord(ByRef frmForm As Access.Form, sFindFirst As String)
    Dim rsRS As Recordset
    Set rsRS = frmForm.RecordsetClone
    rsRS.FindFirst (sFindFirst)
    If Not rsRS.EOF Then frmForm.Bookmark = rsRS.Bookmark
    Set rsRS = Nothing
End Sub


'refresh form, and try to set pointer to last position
Function FormReload(sFormName As String, Optional sByWhatField As String = "0", Optional bStayByIndex = False)
    Dim Sid As String
    
    On Error GoTo formReload_Error
    FormReload = False
    
    With Forms(sFormName)
        'falls numerisch, versuchen spaltenname zu bekommen
        If IsNumeric(sByWhatField) Then
            sByWhatField = .Recordset.Fields(Val(sByWhatField)).Name
        End If
        
        'key spalte speichern
        On Error Resume Next 'falls gelöscht oder sonst irgendwas, ignorieren
        Sid = t.IfNoValue(.Recordset.Fields(sByWhatField).Value, "")
        .Requery
        'zu id springen, falls moeglich
        Call FormMoveToRecord(Forms(sFormName), sByWhatField & " = " & Sid)
        .Recalc
        On Error GoTo formReload_Error
    End With

    FormReload = True
    Exit Function

formReload_Error:
    Err.Raise vbObjectError, "formReload of Modul clFuncsAccs"
    
End Function

'prueft ob form sFormName geladen ist
Function IsFormLoaded(ByVal strFormName As String) As Boolean
'    On Error Resume Next
'    Dim frm As Access.Form
'
'    Set frm = Forms(strFormName)
'
'    isFormLoaded = Not frm Is Nothing
'    Err.Clear


'http://www.aadconsulting.com/aadtips.html
    Const conObjStateClosed = 0
    Const conDesignView = 0

    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) > conObjStateClosed Then
        If Forms(strFormName).CurrentView > conDesignView Then
            IsFormLoaded = True
        End If
    End If
End Function

    'prueft ob form sFormName existiert
Function FormExists(sFormName As String) As Boolean
    On Error Resume Next
    Dim frmDummy As Variant
    If sFormName = vbNullString Then Exit Function
    
    Set frmDummy = Application.CurrentProject.AllForms(sFormName)
    FormExists = (Err.Number = 0)
    Err.Clear
End Function

'try to get control reference from form
Public Function TryGetControl(ByRef vFrmForm_p As Variant, sCtrlName As String, Optional inFailCase As Variant = Nothing) As Control
'    On Error Resume Next
'    Dim retCtrl As Variant
'    Set retCtrl = Access.Forms(vFrmForm_p).Controls(sCtrlName)
'    If err.Number <> 0 Then
'        Set retCtrl = inFailCase
'    End If
'
'    err.Clear
    
    Dim ctrl As Control
    Dim bFound As Boolean
    Dim retCtrl As Variant
    Dim frmForm As Access.Form
    
    If TypeName(vFrmForm_p) = "String" Then
        Set frmForm = Access.Forms(vFrmForm_p)
    Else
        Set frmForm = vFrmForm_p
    End If
    
    For Each ctrl In frmForm.Controls
        If ctrl.Name = sCtrlName Then
            Set retCtrl = ctrl
            bFound = True
            Exit For
        End If
    Next ctrl
    
    If Not bFound Then
        Set retCtrl = inFailCase
    End If
    
    Set TryGetControl = retCtrl
End Function

'set value vValue to property sProp_p, for every Control in Controls Set (Scripting.Dictionary).
Public Function SetPropertyForControls(dictControlsSet As Dictionary, sProp_p As String, vValue As Variant)
    Dim ctrl As Variant 'control
    On Error Resume Next
    For Each ctrl In dictControlsSet.GetValues
        ctrl.Properties(sProp_p) = vValue
    Next ctrl
    
    Err.Clear
End Function

'check if control exists on a form
Public Function IsControlExists(ByRef vForm_p As Variant, sCtrlName As String) As Boolean
    Dim frmForm As Access.Form
    
    If TypeName(vForm_p) = "String" Then
        Set frmForm = Access.Forms(vForm_p)
    Else
        Set frmForm = vForm_p
    End If
    
    IsControlExists = Not (TryGetControl(frmForm, sCtrlName, Nothing) Is Nothing)
End Function

'moves controls in collection as Block of controls to coord left, top.
Public Function MoveControls(dictControls As Scripting.Dictionary, lLeftTw As Long, lTopTw As Long)
    Dim ctrl As Variant
    Dim lastLeft As Long
    Dim lastTop As Long
    Dim diffLeft As Long
    Dim diffTop As Long
    lastLeft = LONGMAX
    lastTop = LONGMAX
    
    On Error Resume Next
    
    For Each ctrl In dictControls.GetValues
        lastTop = t.Min(lastTop, ctrl.Top)
        lastLeft = t.Min(lastLeft, ctrl.Left)
    Next ctrl
    
    'difference between controlsSet most left-top point and new point to move to
    diffLeft = lLeftTw - lastLeft
    diffTop = lTopTw - lastTop
    
    'offset each control
    Call OffsetControls(dictControls, diffLeft, diffTop)
    
    'offset each control
'    For Each ctrl In dictControls.GetValues
'        ctrl.top = ctrl.top + diffTop
'        ctrl.left = ctrl.left + diffLeft
'    Next ctrl
    
    If Err.Number <> 0 Then
        Debug.Print "Fehler bei moveControls " & vbCrLf & _
        "Err.num:" & Err.Number & " Err.Description: " & Err.Description
    End If
    
    Err.Clear
End Function

'change all controls properties LEFT and TOP to LEFT+lLeftTw, TOP+lTopTw
Public Function OffsetControls(dictControls As Scripting.Dictionary, lLeftTw As Long, lTopTw As Long)
    Dim ctrl As Variant
    On Error Resume Next
    
    Dim arrayOfParents As Variant
    arrayOfParents = Array()
    Dim aCtrlAndSize As Variant
    
    'OptionGroup Fix  'TODO all possible parent controls?
    'alle elemente durchsuchen, OptionGroup suchen
    For Each ctrl In dictControls.GetValues
        If ctrl.ControlType = acOptionGroup Then
            'original Width und Height parameter speichern
            Call t.ArrayPush(arrayOfParents, Array(ctrl, ctrl.Width, ctrl.Height, ctrl.Left, ctrl.Top))
        End If
    Next ctrl

    
    'offset each control
    For Each ctrl In dictControls.GetValues
            ctrl.Top = ctrl.Top + lTopTw
            ctrl.Left = ctrl.Left + lLeftTw
    Next ctrl
    
    
    'fuer optionGroup versuchen groesse wiederherstellen
    For Each aCtrlAndSize In arrayOfParents
        'verschen richtige position und groesse nun zuweisen
        aCtrlAndSize(0).Left = aCtrlAndSize(3) + lLeftTw
        aCtrlAndSize(0).Width = aCtrlAndSize(1)
        
        aCtrlAndSize(0).Top = aCtrlAndSize(4) + lTopTw
        aCtrlAndSize(0).Height = aCtrlAndSize(2)
    Next aCtrlAndSize
    
    
    
    If Err.Number <> 0 Then
        Debug.Print "Fehler bei offsetControls " & vbCrLf & _
        "Err.num:" & Err.Number & " Err.Description: " & Err.Description
    End If
    
    Err.Clear
End Function


'---------------------------------------------------------------------------------------
' Function: createAttached
'---------------------------------------------------------------------------------------
' Purpose: Create an attached table in the current database from a table in a different MDB file
' Param  : String strTable name of linked table to create
' Param  : String strPath path and name of DB file containing the table
' Param  : String strBaseTable name of table in strPath DB
' Param  : DAO.Database daoDB (optional) where to create the attached tables. Default = currentDB
' Param  : Boolean bReplace true if need to replace exists table
' Returns: Boolean True/False, indicating success
' Remarks: http://www.microsoftaccessexpert.com/Microsoft-Access-Code-LinkTable.aspx , modified.
'---------------------------------------------------------------------------------------
Function CreateAttachedByPath(strTable As String, _
                        strPath As String, _
                        strBaseTable As String, _
                        Optional daoDB As DAO.Database, _
                        Optional bReplace As Boolean = False) As Boolean
    On Error GoTo CreateAttachedByPathError

    Dim tdf As DAO.TableDef
    Dim dbCurrentDB As DAO.Database

    Set dbCurrentDB = IIf(daoDB Is Nothing, CurrentDb, daoDB)
    Set tdf = dbCurrentDB.CreateTableDef(strTable)

    With tdf
        .Connect = ";DATABASE=" & strPath
        .SourceTableName = strBaseTable
    End With
    
    'if need to replace, find and remove old one
    If bReplace Then
        If TableExists(tdf.Name, daoDB) Then
            dbCurrentDB.TableDefs.Delete tdf.Name
        End If
    End If
    
    dbCurrentDB.TableDefs.Append tdf


CreateAttachedByPathExit:
    CreateAttachedByPath = True
    Exit Function

CreateAttachedByPathError:
    If Err = 3110 Then
        Resume CreateAttachedByPathExit
    Else
        If Err = 3011 Then
            Resume Next
        End If
    End If

End Function
