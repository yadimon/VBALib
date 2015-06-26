VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDirectCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: CDirectCOM (29.06.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DirectCOM-Code von Olaf Schmidt
' </summary>
' <remarks>
' DirectCOM.dll wird benötigt (z.B. ins Anwendungsverzeichnis kopieren
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'


'Helper-Module for the usage of DirectCOM.Dll, wich allows Object-Instantiation
'without using the Registry and also has builtin Threading-Capabilities
'DirectCOM.Dll is not a COM-Dll itself - put it beside your own binaries
'(into your App.Path) and you're done (regarding deployment).
'Normally you shouldn't change the content of this Module (if you go the route,
'as shown in the Demo)
'Have fun! ;-)                         Olaf Schmidt (sss@online.de)
Option Explicit

Private m_hLibModule As Long
Private Const m_conLibFileName As String = "DirectCOM.dll"
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'******** Obj-Instantiation without using the registry *********
Private Declare Function mGetInstance Lib "DirectCOM" Alias "GETINSTANCE" (FName As String, ClassName As String) As Object 'CreateObject-Replacement (FileBased)
Private Declare Function mGETINSTANCELASTERROR Lib "DirectCOM" () As String 'if GETINSTANCE returns Nothing
'----------------------------------------------------------

'******** Thread-Obj-Creation and Threading-Infos *********
Private Type ThreadInit
  Param As Long 'User-Param (given into the ThreadMain-Method of the COM-Dll)
  hThread As Long 'filled from Thread (instantly, after STARTCOMOBJECT() returns)
  ThreadID As Long 'filled from Thread (if fully initialized, some msec after STARTCOMOBJECT())
  Reserved(511) As Byte 'reserved for Thread-Internal usage (don't touch this)
End Type
Private Declare Function STARTCOMOBJECT Lib "DirectCOM" (FName As String, ClassName As String, TI As ThreadInit) As Long
Private Declare Function GETTHREADCOUNT Lib "DirectCOM" () As Long
Private Declare Function GETTHREADSTATUS Lib "DirectCOM" (ByVal hThread&) As Long '259=Thread is running; 258=Couldn't create ThreadObject (Threadfunction already exited); all other Values are set by the exiting ThreadMain()-Method
Private Declare Function CLOSETHREADHANDLE Lib "DirectCOM" (hThread&) As Long
'----------------------------------------------------------

'******** Thread-Synchronizing *********
Private Type CRITICAL_SECTION
  DebugInfo As Long
  LockCount As Long
  RecursionCount As Long
  OwningThreadID As Long
  LockSemaphore As Long
  SpinCount As Long
End Type
Private Declare Sub InitializeCriticalSection Lib "kernel32" (CS As CRITICAL_SECTION)
Private Declare Sub EnterCriticalSection Lib "kernel32" (ByVal pCRITICAL_SECTION&)
Private Declare Sub LeaveCriticalSection Lib "kernel32" (ByVal pCRITICAL_SECTION&)
Private Declare Sub DeleteCriticalSection Lib "kernel32" (CS As CRITICAL_SECTION)
Private dh_CrSct As CRITICAL_SECTION, p_dh_CrSct As Long
'----------------------------------------------------------

'******** useful stuff (Thread-Communicaiton, -Termination, etc. *********
Private Type TMsg
  hWnd As Long
  Message As Long
  wParam As Long
  lParam As Long
  time As Long
  PtX As Long
  PtY As Long
End Type
Private Const THREAD_PRIORITY_IDLE& = -15
Private Const THREAD_PRIORITY_LOWEST& = -2
Private Const THREAD_PRIORITY_BELOW_NORMAL& = -1
Private Const THREAD_PRIORITY_NORMAL& = 0
Private Const THREAD_PRIORITY_ABOVE_NORMAL& = 1
Private Const THREAD_PRIORITY_HIGHEST& = 2
Private Const THREAD_PRIORITY_TIME_CRITICAL& = 15
Private Const ThrMsg& = &H444 'some value above WM_USER (&H400)
Private Declare Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ArrPtr& Lib "msvbvm60" Alias "VarPtr" (P() As Any)
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread&, ByVal dwExitCode&) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As TMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
'We rename the Sleep-API, to allow Public Declaration (and usage).
'The other Public declared APIs above are normally not very usual and so don't
'have a high potential, to conflict with already Public defined "Duplicates".
Private Declare Sub dh_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Dst As Any, Src As Any, ByVal cb&)
'----------------------------------------------------------

'************* Some Public Helper-Functions ***************
Public Sub InitCrSctVars()
  InitializeCriticalSection dh_CrSct
  p_dh_CrSct = VarPtr(dh_CrSct)
End Sub
Public Sub CleanupCrSctVars()
  DeleteCriticalSection dh_CrSct
End Sub

Public Sub BindArrayToPointer(ByVal pDst As Long, ByVal pSrc As Long)
  RtlMoveMemory ByVal pDst, ByVal pSrc, 4
End Sub
Public Sub ReleaseArrayBinding(ByVal pArray As Long)
  RtlMoveMemory ByVal pArray, 0&, 4
End Sub

Public Function WaitForMessages(Optional ByVal msecTimeOut& = -1) As Boolean
  WaitForMessages = MsgWaitForMultipleObjects(0, ByVal 0&, 0, msecTimeOut, 255) <> 258
End Function
'----------------------------------------------------------

'Klassen-Umformung

Public Function GetInstance(FName As String, ClassName As String) As Object
    If InitLibDLL Then
        Set GetInstance = mGetInstance(FName, ClassName)
    Else
        Set GetInstance = Nothing
    End If
'Private Declare Function GETINSTANCELASTERROR Lib "DirectCom" () As String 'if GETINSTANCE returns Nothing
End Function

Public Function GetInstanceLastError() As String
    If InitLibDLL Then
        GetInstanceLastError = mGETINSTANCELASTERROR()
    Else
        GetInstanceLastError = "Die Datei " & m_conLibFileName & " wurde nicht gefunden."
    End If
End Function

Private Sub Class_Initialize()
    InitLibDLL
End Sub

Private Sub Class_Terminate()
    FreeLibDLL
End Sub

Function InitLibDLL() As Boolean
    If m_hLibModule = 0 Then
        m_hLibModule = LoadLibrary(m_conLibFileName)     'Systempfad?
        If m_hLibModule = 0 Then
            m_hLibModule = LoadLibrary(CurrentProject.Path & "\" & m_conLibFileName)    'Datenbankpfad?
        End If
        If m_hLibModule = 0 Then
            MsgBox "Die benötigte Datei " & m_conLibFileName & " konnte nicht gefunden werden", vbCritical
        End If
    End If
    InitLibDLL = (m_hLibModule <> 0)
End Function

Function FreeLibDLL()
    If m_hLibModule <> 0 Then
        Call FreeLibrary(m_hLibModule)
    End If
End Function

