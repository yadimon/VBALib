Attribute VB_Name = "mdlDevelopHelpFuncs"
'---------------------------------------------------------------------------------------
' Module : mdlDevelopHelpFuncs
' Author : Dmitry Gorelenkov
' Date   : 16.05.2014
' Purpose: help functions for development
' Version: 0.9
' Remarks:
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/Access/mdlDevelopHelpFuncs.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/Access/clsToolsAccs.cls</use>
'  <ref><name>ADODB</name><major>6</major><minor>1</minor><guid>{B691E011-1797-432E-907A-4D8C69339129}</guid></ref>
'  <ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>
'  <ref><name>VBScript_RegExp_55</name><major>5</major><minor>5</minor><guid>{3F4DACA7-160D-11D2-A8E9-00104B365C9F}</guid></ref>
'  <ref><name>VBIDE</name><major>5</major><minor>3</minor><guid>{0002E157-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

' Tables fuer update version
Private Const m_sSETTINGS_TABLE_NAME As String = "Settings"
Private Const m_sCHANGE_LOG_TABLE_NAME As String = "ChangeLog"


'---------------------------------------------------------------------------------------
' Function: closeAllForms
'---------------------------------------------------------------------------------------
' Purpose: closes all opened forms with acSaveNo option
' Returns: Variant
'---------------------------------------------------------------------------------------
Public Function CloseAllForms()
    Dim frmForm As Variant
    For Each frmForm In Forms
        DoCmd.Close acForm, frmForm.Name, acSaveNo
    Next frmForm
End Function

'---------------------------------------------------------------------------------------
' Sub: Get_References_in_Project
'---------------------------------------------------------------------------------------
' Purpose: get all references of access-project. Print it in debug window. Prints CodeLib tags too.
'---------------------------------------------------------------------------------------
Sub Get_References_in_Project()
    Dim bRefBroken As Boolean
    Dim sAllInclude As String
    Dim i1
    For i1 = 1 To Application.VBE.ActiveVBProject.References.Count
        With Application.VBE.ActiveVBProject.References(i1)
        ' Get the Name of the Reference
        Debug.Print "REF NAME:" & .Name

        ' Get the Description of Reference
        Debug.Print "DESCRIPTION:" & .Description
        Debug.Print "GUID:" & .Guid
        Debug.Print "Major:" & .Major
        Debug.Print "Minor:" & .Minor

        'Returns a Boolean value indicating whether or not the Reference object points to a valid reference in the registry. Read-only.
        bRefBroken = False
        If .IsBroken = True Then
            bRefBroken = True
        End If
        
        Debug.Print "IS BROKEN?:" & CStr(bRefBroken) & vbNewLine & vbNewLine
        'example: <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
        sAllInclude = sAllInclude _
        & "<ref><name>" & .Name & "</name><major>" & .Major & "</major><minor>" & .Minor & "</minor><guid>" & .Guid & "</guid></ref>" _
        & vbNewLine
        End With
    Next i1
    
    Debug.Print "CodeLib tags:" & vbNewLine & sAllInclude

End Sub



'prints an Array if bToDbg not defined, and returns printed string. bHoriz for horizontal printing
Public Function PrintArray(ByRef aArray, Optional bHoriz As Boolean = False, Optional bToDbg As Boolean = True) As String
    Dim elem As Variant
    Dim elemStr As String
    Dim counter As Long
    Dim retString As String
    
    counter = LBound(aArray)
    On Error Resume Next
    retString = "Array(" & IIf(bHoriz, "", vbCrLf)
    For Each elem In aArray
        If IsArray(elem) Then
            elemStr = PrintArray(elem, True, False)
        ElseIf IsNull(elem) Then
            elemStr = "%<NULL>%"
        ElseIf IsEmpty(elem) Then
            elemStr = "%<Empty>%"
        ElseIf IsObject(elem) Then
            elemStr = "%<Object>%"
        Else
            elemStr = CStr(elem)
            If Err.Number <> 0 Then
                elemStr = "%<Error>%"
            End If
        End If
        retString = retString & counter & " => " & elemStr & IIf(bHoriz, "; ", vbCrLf)
        counter = counter + 1
    Next elem
    retString = retString & ")"
    
    If bToDbg Then Debug.Print retString
    PrintArray = retString
End Function


'---------------------------------------------------------------------------------------
' Function: updateVersion
'---------------------------------------------------------------------------------------
' Purpose: Update current version. Save changes in changelog
' Param  : String sDescription description
' Returns: Variant
' Remarks:
Public Function UpdateVersion(sDescription As String)
    Dim tAccs As New clsToolsAccs
    Dim dbase As Database
    Dim changelogTableRs As Recordset
    Dim sVersion As String
    Dim aNewVersionSplitted As Variant
    On Error GoTo updateVersion_Error
    
    'TODO create tables
    If Not tAccs.TableExists(m_sCHANGE_LOG_TABLE_NAME) Then
        MsgBox "Log Table: " & m_sCHANGE_LOG_TABLE_NAME & " doesnt exists"
        Exit Function
    End If
    
    If Not tAccs.TableExists(m_sSETTINGS_TABLE_NAME) Then
        MsgBox "Settings Table: " & m_sSETTINGS_TABLE_NAME & " doesnt exists"
        Exit Function
    End If
    
    Set dbase = CurrentDb()

    Set changelogTableRs = dbase.TableDefs(m_sCHANGE_LOG_TABLE_NAME).OpenRecordset(dbOpenDynaset)
    
    aNewVersionSplitted = Split(CurrentVersion, ".")
    aNewVersionSplitted(2) = CStr(CLng(aNewVersionSplitted(2)) + 1)
    sVersion = Join(aNewVersionSplitted, ".")
    
    With changelogTableRs
        .AddNew
'        .Edit
        !Version = sVersion
        !CommitDate = Now
        !Description = sDescription
        .Update
        .Close
    End With
    
    'current version update
    'Dim DBset As New ApplicationHandler_DBSettings
    'Call DBset.SetSetting("Version", sVersion)



Finally:
    On Error Resume Next
    dbase.Close
    'Set DBset = Nothing
    Set tAccs = Nothing
    Set dbase = Nothing
    On Error GoTo 0
    Exit Function

updateVersion_Error:

    Call hErr("updateVersion of Module mdlApplication", , , aclibErrMsgBox)
    Resume Finally
End Function

'---------------------------------------------------------------------------------------
' Property: CurrentVersion
'---------------------------------------------------------------------------------------
' Purpose: get current version from Settings table
' Returns: Variant version
'---------------------------------------------------------------------------------------
Property Get CurrentVersion()
    CurrentVersion = DLookup("Version", m_sCHANGE_LOG_TABLE_NAME, "ID = " & DMax("ID", m_sCHANGE_LOG_TABLE_NAME))
End Property


'---------------------------------------------------------------------------------------
' Sub: ExportAll
'---------------------------------------------------------------------------------------
' Purpose: export all files to filesystem
' Remarks: uses ACLibImportWizard
'---------------------------------------------------------------------------------------
Public Sub ExportAll()
   Application.Run "C:\Users\" & Environ("username") & "\AppData\Roaming\Microsoft\AddIns\ACLibImportWizard.ExportAllCodeLibElements", False
End Sub


'---------------------------------------------------------------------------------------
' Sub: Commit
'---------------------------------------------------------------------------------------
' Purpose: extract all files with ImportWizard and commit everytig with git + updates current Tool version
' Param  : String sMsg commit message
' Param  : Boolean bGlobalSettings [True] define if global settings used or not. If not "Author" wii be used form last commit
' Remarks: uses windows Shell, to commit. Addin must be loaded/opened b4. TODO: last commited author dont work. (because of for loop?)
'---------------------------------------------------------------------------------------
Public Sub Commit(sMsg As String, Optional bGlobalSettings As Boolean = True)
    Dim sGitExePath As String
    Dim sPathToExportedFiles As String
    Dim sAddCommand As String
    Dim sComitCommand As String
    Dim sFullCommand As String
    Dim sSetAuthorVarCommand As String
    Dim sGetAuthorCommand As String
    Dim sCMDSep As String 'command separator
    Dim sAuthorDef As String
    Dim wsh As Object
    
    sCMDSep = " && "
    sPathToExportedFiles = CurrentProject.Path '"H:\PruefungsDB"
    sGitExePath = "C:\Users\z003c9tt\AppData\Local\GitHub\PortableGit_054f2e797ebafd44a30203088cd3d58663c627ef\cmd\git.exe"
    
    sGetAuthorCommand = sGitExePath & " --no-pager show -s --format^=""%an <%ae>""" '^= statt = wegen batch schleife
    sSetAuthorVarCommand = "for /f ""delims=;"" %i in ('" & sGetAuthorCommand & "') do set AUTHOR=%i" '""delims=;"" da ich leerzeichen brauche
    sAuthorDef = IIf(bGlobalSettings, "", " --author=""%AUTHOR%""")
    
    sAddCommand = sGitExePath & " add -A"
    sComitCommand = sGitExePath & " commit -a" & sAuthorDef & " -m """ & sMsg & """"
    
    'command to execute
    If bGlobalSettings Then
        sFullCommand = sAddCommand & sCMDSep & sComitCommand
    Else
        sFullCommand = sSetAuthorVarCommand & _
                    sCMDSep & sAddCommand & _
                    sCMDSep & sComitCommand
    End If
    
    'Shell object
    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.CurrentDirectory = sPathToExportedFiles
    
    'export all modules etc
    ExportAll
    
    'update local version
    Call UpdateVersion(sMsg)
    
    
    'tmp ignore dont work
    'commit over git
'    Dim sIgnoreThisProjFile As String
'    Dim sIgnoreThisProjTmpFile As String
'    sIgnoreThisProjFile = CurrentProject.FullName
'    sIgnoreThisProjTmpFile = replace(CurrentProject.FullName, ".accdb", ".laccdb")
    'ignore curent opened project
    
'    Call wsh.Run("cmd.exe /k " & sGitExePath & " update-index --assume-unchanged " & sIgnoreThisProjFile, 1, True)
'    Call wsh.Run("cmd.exe /k " & sGitExePath & " update-index --assume-unchanged " & sIgnoreThisProjTmpFile, 1, True)

    'commit
    Call wsh.Run("cmd.exe /k " & sFullCommand, 1, True)
    
    
    'undo ignore current project
'    Call wsh.Run("cmd.exe /k " & sGitExePath & " update-index --no-assume-unchanged " & sIgnoreThisProjFile, 1, True)
'    Call wsh.Run("cmd.exe /k " & sGitExePath & " update-index --no-assume-unchanged " & sIgnoreThisProjTmpFile, 1, True)
End Sub


Sub GenerateCreateObject()
    Dim Component As Variant 'VBComponent
    Dim sResult As String
    Dim sSelectCase As String
    
    For Each Component In Application.VBE.ActiveVBProject.VBComponents
        If Component.Type = vbext_ct_ClassModule Then 'if class
            sSelectCase = sSelectCase & vbTab & vbTab & "Case """ & Component.Name & """" & vbCrLf & _
            vbTab & vbTab & vbTab & "Set CreateObject = New " & Component.Name & vbCrLf
        End If
        
    Next Component
    
    
    sResult = _
"'Generated by ""GenerateCreateObject"" sub" & vbCrLf & _
"Public Function CreateObject(sClassName As String, Optional sServerName As String) As Object" & vbCrLf & _
"    Select Case sClassName" & vbCrLf & _
        sSelectCase & vbCrLf & _
"        Case Else" & vbCrLf & _
"            Set CreateObject = VBA.CreateObject(sClassName, sServerName)" & vbCrLf & _
"    End Select" & vbCrLf & _
"End Function"

Debug.Print sResult

End Sub


Sub ListOfNotDescribedModules()
    Dim Component As VBIDE.VBComponent
    Dim sResult As String
    Const sSearchString As String = "'<codelib>"
    Const sSearchString2 As String = "' Author : Dmitry Gorelenkov"
    
    For Each Component In Application.VBE.ActiveVBProject.VBComponents
        If Component.Type = vbext_ct_ClassModule Or Component.Type = vbext_ct_StdModule Or Component.Type = vbext_ct_MSForm Then 'if class, module or form
            If Not Component.CodeModule.Find(sSearchString, 1, 1, -1, -1, True) And _
                Not Component.CodeModule.Find(sSearchString2, 1, 1, -1, -1, True) Then
                Debug.Print Component.Name
            End If
        End If
        
    Next Component
End Sub

'veeery bad, but usable. TODO make with arrays, search for needed reference/class, output possible solutions etc... and maybe as access plugin
Sub FindMissingReferencesInCode()
    Dim Component As VBIDE.VBComponent
    Const sPattern As String = "^[^'\n]*(((Dim|Private|Public) [a-zA-Z0-9_]+ as (New )?)|(Set .+? = New ))([a-zA-Z0-9_]+)(\.[a-zA-Z0-9_]+)?$"    'still bad :/
    'Const sPattern As String = "^[^'\n]*(((Dim|Private|Public) [a-zA-Z_]+ as (New )?)|(Set .+? = New ))([a-zA-Z_\.]*)$"
    Dim sCodeContent As String
    Dim aMatch As MatchCollection
    Dim FoundMatch As Match
    Dim sFoundType As String
    Dim sFoundAll As String
    Dim sMustBePresent As String
    Dim t As New clsTools
    Dim aIgnoredTypes As Variant
    Dim dicAllRefs As Scripting.Dictionary
    Set dicAllRefs = FindAllRefsAndFileDefs()
    
    aIgnoredTypes = Array("Variant", "String", "Date", "Integer", "Long", "Double", "Object", "Collection", "Boolean", "Currency", _
                            "Database", "Recordset", "Control", "DAO", "FileDialog", "ApplicationHandlerResumeModes", "Access")
    Dim mustContainIf As New Scripting.Dictionary

    Const sAdo = "<ref><name>ADODB</name><major>6</major><minor>1</minor><guid>{B691E011-1797-432E-907A-4D8C69339129}</guid></ref>"
    Const sAdoX = "<ref><name>ADOX</name><major>6</major><minor>0</minor><guid>{00000600-0000-0010-8000-00AA006D2EA4}</guid></ref>"
    Const sScripting = "<ref><name>Scripting</name><major>1</major><minor>0</minor><guid>{420B2830-E718-11CF-893D-00A0C9054228}</guid></ref>"
    Const sRegExp = "<ref><name>VBScript_RegExp_55</name><major>5</major><minor>5</minor><guid>{3F4DACA7-160D-11D2-A8E9-00104B365C9F}</guid></ref>"
    Const sVBIDE = "<ref><name>VBIDE</name><major>5</major><minor>3</minor><guid>{0002E157-0000-0000-C000-000000000046}</guid></ref>"
    Const sDao = "<ref><name>DAO</name><major>12</major><minor>0</minor><guid>{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}</guid></ref>"
    Const sOffice = "<ref><name>Office</name><major>2</major><minor>5</minor><guid>{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}</guid></ref>"
    Const sExcel = "<ref><name>Excel</name><major>1</major><minor>7</minor><guid>{00020813-0000-0000-C000-000000000046}</guid></ref>"
    Const sMscoree = "<ref><name>mscoree</name><major>2</major><minor>0</minor><guid>{5477469E-83B1-11D2-8B49-00A0C9B7C9C4}</guid></ref>"
    Const sMscorlib = "<ref><name>mscorlib</name><major>2</major><minor>0</minor><guid>{BED7F4EA-1A96-11D2-8F08-00A0C9A6186D}</guid></ref>"
    Const sAccUnit_Integration = "<ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>"
    Const sComObjectsLib = "<ref><name>ComObjectsLib</name><major>1</major><minor>0</minor><guid>{7B2A26C5-9D8E-4830-9CA2-3CD5FC3227ED}</guid></ref>"
    
    With mustContainIf
        'default included
        'Call .Add("DAO", sDao)
        'Call .Add("FileDialog", sOffice)
        
        Call .Add("ADODB", sAdo)
        Call .Add("ADOX", sAdoX)
        Call .Add("Scripting", sScripting)
        Call .Add("FileSystemObject", sScripting)
        Call .Add("RegExp", sRegExp)
        Call .Add("Match", sRegExp)
        Call .Add("MatchCollection", sRegExp)
        Call .Add("VBScript_RegExp_55", sRegExp)
        Call .Add("VBIDE", sVBIDE)
        Call .Add("Excel", sExcel)
        Call .Add("mscoree", sMscoree)
        Call .Add("mscorlib", sMscorlib)
        Call .Add("AccUnit_Integration", sAccUnit_Integration)
        Call .Add("ComObjectsLib", sComObjectsLib)
    End With

    Dim objRegex As New RegExp
    With objRegex
        .Multiline = True
        .IgnoreCase = True
        .Pattern = sPattern
        .Global = True
    End With
    
    For Each Component In Application.VBE.ActiveVBProject.VBComponents
        If t.ArrayContains(Component.Type, Array(vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_Document)) Then 'if class, module or form
        'If Component.Type = vbext_ct_ClassModule Or Component.Type = vbext_ct_StdModule Or Component.Type = vbext_ct_MSForm Or Component.Type = vbext_ct_Document Then
            sCodeContent = Component.CodeModule.Lines(1, Component.CodeModule.CountOfLines)
            If Not sCodeContent Like "*'<codelib>*" Then GoTo nextComp 'only components with codelib comments
            
            'find all refs, and test if they are described
            With objRegex
                If .Test(sCodeContent) Then
                    sFoundAll = ""
                    Set aMatch = .Execute(sCodeContent)
                    For Each FoundMatch In aMatch
                        sFoundType = FoundMatch.SubMatches(5)
                        'if should not be ignored
                        If Not t.ArrayContains(sFoundType, aIgnoredTypes, True, True, True) Then
                            'test if enum or type
                            If sCodeContent Like "*enum " & sFoundType & "*" Then GoTo nextMatch
                            If sCodeContent Like "*type " & sFoundType & "*" Then GoTo nextMatch
                            
                            'define description, must be present
                            If mustContainIf.Exists(sFoundType) Then
                                sMustBePresent = "'  " & mustContainIf.Item(sFoundType)
                            Else
                                'like /clsTools.???</use>
                                sMustBePresent = "/" & sFoundType & ".???</use>"
                            End If
                            'show only if not ref to it in header
                            If Not (sCodeContent Like "*" & sMustBePresent & "*") Then
                                'if we have full link, add it
                                If dicAllRefs.Exists(sFoundType) Then
                                    sFoundType = dicAllRefs.Item(sFoundType)
                                End If
                                'add to result only if not added already
                                If Not CBool(InStr(1, sFoundAll, sFoundType)) Then
                                    sFoundAll = sFoundAll & sFoundType & vbNewLine
                                End If
                            End If

                        End If
nextMatch:
                    Next FoundMatch

                    If Len(sFoundAll) > 0 Then
                        Debug.Print Component.Name & ":" & vbNewLine & sFoundAll
                    End If
                End If
            End With

        End If
nextComp:
    Next Component


    Set objRegex = Nothing
    Set t = Nothing
End Sub

'dirty fast code
Private Function FindAllRefsAndFileDefs() As Scripting.Dictionary
    Const sFileRef As String = "^('\s*<file>.*/(.{1,}?)\..{3}</file>)$"
    Const sRef As String = "^('\s*<ref><name>([a-zA-Z0-9_]{1,})\.?.*?</name><maj.*?</guid></ref>)$"
    
    Dim dicResult As New Scripting.Dictionary
    Dim sFoundType As String
    Dim sFoundString As String
    Dim Component As VBIDE.VBComponent
    Dim sCodeContent As String
    Dim aMatch As MatchCollection
    Dim FoundMatch As Match
    Dim sFoundAll As String
    
    Dim objRegexFile As New RegExp
    Dim objRegexRefs As New RegExp
    
    With objRegexFile
        .Multiline = True
        .IgnoreCase = True
        .Pattern = sFileRef
        .Global = True
    End With
    
     With objRegexRefs
        .Multiline = True
        .IgnoreCase = True
        .Pattern = sRef
        .Global = True
    End With
    
    
    For Each Component In Application.VBE.ActiveVBProject.VBComponents
        If t.ArrayContains(Component.Type, Array(vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_Document)) Then 'if class, module or form
        'If Component.Type = vbext_ct_ClassModule Or Component.Type = vbext_ct_StdModule Or Component.Type = vbext_ct_MSForm Or Component.Type = vbext_ct_Document Then
            sCodeContent = Component.CodeModule.Lines(1, Component.CodeModule.CountOfLines)
            If Not sCodeContent Like "*'<codelib>*" Then GoTo nextComp 'only components with codelib comments
            With objRegexFile
                If .Test(sCodeContent) Then
                    Set aMatch = .Execute(sCodeContent)
                    For Each FoundMatch In aMatch
                        sFoundType = FoundMatch.SubMatches(1)
                        sFoundString = FoundMatch.SubMatches(0)
                        If Not dicResult.Exists(sFoundType) Then
                            Call dicResult.Add(sFoundType, Replace(sFoundString, "file>", "use>"))
                        End If
                    Next FoundMatch
                End If
            End With
            
            With objRegexRefs
                If .Test(sCodeContent) Then
                    Set aMatch = .Execute(sCodeContent)
                    For Each FoundMatch In aMatch
                        sFoundType = FoundMatch.SubMatches(1)
                        sFoundString = FoundMatch.SubMatches(0)
                        If Not dicResult.Exists(sFoundType) Then
                            Call dicResult.Add(sFoundType, sFoundString)
                        End If
                    Next FoundMatch
                End If
            End With
            

        End If
nextComp:
    Next Component
    
    Set FindAllRefsAndFileDefs = dicResult
End Function
