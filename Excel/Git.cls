VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Git"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : Git
' Author : Dmitry Gorelenkov
' Date   : 23.03.2015
' Purpose: use Git from VBA
' Version: 0.9
' Remarks: not ready yet. TODO: update from source. Save/Load to/from in file defined sub folders.
' Remarks: uses "Microsoft Visual Basic for Application Extensibility 5.3" reference
' Remarks: singleton! Attribute VB_PredeclaredId = True
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/Git.cls</file>
'  <license>VBALib/license.bas</license>
'  <use>VBALib/clsTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Public WorkingDir As String
Public SubFolderForExport As String
Private t As clsTools
'---------------------------------------------------------------------------------------
' Sub: SaveCodeModules
'---------------------------------------------------------------------------------------
' Purpose: save code modules for version control
' Remarks: thx http://stackoverflow.com/a/2003792/2519073
'---------------------------------------------------------------------------------------
Function SaveCodeModules()
    Dim i%, sName$, Ext$, SubFolder$, PathToSave$ 'performance not importatnt, +short code
    
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                sName$ = .VBComponents(i%).CodeModule.Name
                
                Select Case .VBComponents(i%).Type
                    Case vbext_ct_Document
                        SubFolder = "Worksheets"
                        Ext = "cls"
                    Case vbext_ct_StdModule
                        SubFolder = "Modules"
                        Ext = "bas"
                    Case vbext_ct_ClassModule
                        SubFolder = "Classes"
                        Ext = "cls"
                    Case vbext_ct_MSForm
                        SubFolder = "Forms"
                        Ext = "frm"
                End Select
                
                PathToSave = WorkingDir & "\" & SubFolderForExport & "\" & SubFolder & "\"
                Call t.CreateDirectory(PathToSave) 'ensure dir exists, or create otherwise
                Call .VBComponents(i%).Export(PathToSave & sName$ & "." & Ext)
                
            End If
        Next i
    End With
    
    Debug.Print "Ready"
End Function

Public Function Commit(Optional sMessage As String = "")
    Call SaveCodeModules
    Call GitCall("add .")
    Call GitCall("commit -m """ & sMessage & """")
End Function

Public Function Add()
    Call SaveCodeModules
    Call GitCall("add")
End Function

Public Function GitCall(sCommand As String)
    
    Dim bWaitOnReturn As Boolean: bWaitOnReturn = True
    Dim iWindowStyle As Integer: iWindowStyle = 1
    Call RunShell(GetGitPathToExe() & " " & sCommand)
    
End Function

'if GitHub is installed
Private Function GetGitPathToExe() As String
    Dim PathToLocal As String
    Dim sGitFolderName As String
    
    PathToLocal = Environ("LOCALAPPDATA")
    
    sGitFolderName = Dir(PathToLocal & "\GitHub\PortableGit_*", vbDirectory)
    
    GetGitPathToExe = PathToLocal & "\GitHub\" & sGitFolderName & "\bin\git.exe"
End Function

'TODO add to tools?
Private Function RunShell(sCommand As String, Optional sWorkingDir As String, _
                        Optional bWaitOnReturn As Boolean = True, Optional iWindowStyle As Integer = 1)
                        
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    
    wsh.CurrentDirectory = IIf(Len(sWorkingDir), sWorkingDir, Me.WorkingDir)
    
    Call wsh.Run("%comspec% /C " & sCommand, iWindowStyle, bWaitOnReturn)
    
    Set wsh = Nothing
End Function


Private Sub Class_Initialize()
    WorkingDir = ThisWorkbook.Path 'default
    SubFolderForExport = "src"
    Set t = New clsTools
End Sub

Private Sub Class_Terminate()
    Set t = Nothing
End Sub
