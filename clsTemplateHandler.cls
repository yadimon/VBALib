VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTemplateHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module : clsTemplateHandler
' Author : Dmitry Gorelenkov
' Date   : 23.05.2014
' Purpose: work with templates like "{param} .. {value}"
' Version: 0.9
' Remarks: uses {} to identify variables. TODO for '{' symbol use {{
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>VBALib/clsTemplateHandler.cls</file>
'  <license>VBALib/license.bas</license>
'  <use></use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Function: apply
'---------------------------------------------------------------------------------------
' Purpose: apply parameters from dictVars to sTemplate
' Param  : String sTemplate template
' Param  : Scripting.Dictionary dictVars variables in dictionary ("replace me(key)", "with me(value)")
' Returns: String result text
' Remarks: TODO values type recogniiton
'---------------------------------------------------------------------------------------
Public Function apply(sTemplate As String, dictVars As Scripting.Dictionary) As String
    Dim sKey As Variant
    
    For Each sKey In dictVars.Keys
        sTemplate = Replace$(sTemplate, "{" & sKey & "}", CStr(dictVars.Item(sKey)), , , vbTextCompare)
    Next sKey
    
    apply = sTemplate
End Function
