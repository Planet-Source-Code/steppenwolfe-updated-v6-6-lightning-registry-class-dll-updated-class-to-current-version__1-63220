VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegBackup"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//key and security constants
Private Const TOKEN_QUERY                     As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES         As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED            As Long = &H2
Private Const REG_FORCE_RESTORE              As Long = 8&
Private Const READ_CONTROL                   As Long = &H20000
Private Const SYNCHRONIZE                    As Long = &H100000
Private Const STANDARD_RIGHTS_READ           As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE          As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL            As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL            As Long = &HFFFF
Private Const KEY_QUERY_VALUE                As Long = &H1
Private Const KEY_SET_VALUE                  As Long = &H2
Private Const KEY_CREATE_SUB_KEY             As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS         As Long = &H8
Private Const KEY_NOTIFY                     As Long = &H10
Private Const KEY_CREATE_LINK                As Long = &H20
Private Const KEY_READ                       As Double = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS                 As Double = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'//account access levels
Private Const SE_ASSIGN_PRIMARY_TOKEN       As String = "SeAssignPrimaryTokenPrivilege"
Private Const SE_AUDIT                      As String = "SeAuditPrivilege"
Private Const SE_BACKUP_NAME                As String = "SeBackupPrivilege"
Private Const SE_CHANGE_NOTIFY              As String = "SeChangeNotifyPrivilege"
Private Const SE_CREATE_PERMANENT           As String = "SeCreatePermanentPrivilege"
Private Const SE_CREATE_TOKEN               As String = "SeCreateTokenPrivilege"
Private Const SE_INCREASE_BASE_PRIORITY     As String = "SeIncreaseBasePriorityPrivilege"
Private Const SE_INCREASE_PAGE_FILE         As String = "SeCreatePagefilePrivilege"
Private Const SE_INCREASE_QUOTA             As String = "SeIncreaseQuotaPrivilege"
Private Const SE_LOAD_DRIVER                As String = "SeLoadDriverPrivilege"
Private Const SE_LOCK_MEMORY                As String = "SeLockMemoryPrivilege"
Private Const SE_MACHINE_ACCOUNT            As String = "SeMachineAccountPrivilege"
Private Const SE_PROFILE_SINGLE_PROC        As String = "SeProfileSingleProcessPrivilege"
Private Const SE_REMOTE_SHUTDOWN            As String = "SeRemoteShutdownPrivilege"
Private Const SE_RESTORE_NAME               As String = "SeRestorePrivilege"
Private Const SE_SECURITY                   As String = "SeSecurityPrivilege"
Private Const SE_SERVICE_LOGON              As String = "SeServiceLogonRight"
Private Const SE_SHUT_DOWN                  As String = "SeShutdownPrivilege"
Private Const SE_SYSTEM_ENVIROMENT          As String = "SeSystemEnvironmentPrivilege"
Private Const SE_SYSTEM_PROFILE             As String = "SeSystemProfilePrivilege"
Private Const SE_SYSTEM_TIME                As String = "SeSystemtimePrivilege"
Private Const SE_TAKE_OWNERSHIP             As String = "SeTakeOwnershipPrivilege"
Private Const SE_TCB                        As String = "SeTcbPrivilege"
Private Const SE_UNSOLICITED_INPUT          As String = "SeUnsolicitedInputPrivilege"

'//se access enum
Public Enum eSEUserRights
    eAssignPrimaryToken = 0
    eAudit = 1
    eBackupName = 2
    eChangeNotify = 3
    eCreatePermanent = 4
    eCreateToken = 5
    eIncreaseBasePriority = 6
    eIncreasePageFile = 7
    eIncreseQuota = 8
    eLoadDriver = 9
    eLockMmeory = 10
    eMachineAccount = 11
    eProfileSingleProc = 12
    eRemoteShutdown = 13
    eRestoreName = 14
    eSecurity = 14
    eServiceLogon = 16
    eShutdown = 17
    eSystemEnviroment = 18
    eSystemProfile = 19
    eSystemTime = 20
    eTakeOwnership = 21
    eTcb = 22
    eUnsolicitedInput = 23
End Enum

'//token structures
Private Type tLuid
    lowpart                                      As Long
    highpart                                     As Long
End Type

Private Type tAttributes
    pLuid                                        As tLuid
    Attributes                                   As Long
End Type

Private Type tPrivilege
    PrivilegeCount                               As Long
    Privileges                                   As tAttributes
End Type

'//reg api
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpFile As String, _
                                                                            lpSecurityAttributes As Any) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long

Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, _
                                                                                  ByVal lpFile As String, _
                                                                                  ByVal dwFlags As Long) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, _
                                                                   ByVal DisableAllPriv As Long, _
                                                                   NewState As tPrivilege, _
                                                                   ByVal BufferLength As Long, _
                                                                   PreviousState As tPrivilege, _
                                                                   ReturnLength As Long) As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, _
                                                                                                ByVal lpName As String, _
                                                                                                lpLuid As tLuid) As Long
'//process api
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, _
                                                              ByVal DesiredAccess As Long, _
                                                              TokenHandle As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long


Private Function cSet_Priv(ByVal Access As String) As Long

'//set user privilege for backup
Dim lReturn             As Long
Dim lToken              As Long
Dim lLength             As Long
Dim tpeLuid             As tLuid
Dim tpeTokenPriv        As tPrivilege
Dim tpePrevPriv         As tPrivilege

On Error GoTo Handler

    '//open current process token
    lReturn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lToken)
    If lReturn = 0 Then
        cSet_Priv = 0
        Exit Function
    ElseIf Err.LastDllError <> 0 Then
        Exit Function
    End If
    '//test access level
    lReturn = LookupPrivilegeValue(0&, Access, tpeLuid)
    If lReturn = 0 Then
        cSet_Priv = 0
        Exit Function
    End If
    
    With tpeTokenPriv
        .PrivilegeCount = 1
        .Privileges.Attributes = SE_PRIVILEGE_ENABLED
        .Privileges.pLuid = tpeLuid
    End With
    '//if token change is allowed, return is non 0
    cSet_Priv = AdjustTokenPrivileges(lToken, False, tpeTokenPriv, Len(tpePrevPriv), tpePrevPriv, lLength)

Handler:

End Function

Private Function cTest_Priv(ByVal Access As String) As Boolean

'//Get user privilege for level

Dim lReturn             As Long
Dim lToken              As Long
Dim lLength             As Long
Dim tpeLuid             As tLuid
Dim tpeTokenPriv        As tPrivilege
Dim tpePrevPriv         As tPrivilege

On Error GoTo Handler

    '//open current process token
    lReturn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lToken)
    If lReturn = 0 Then
        cTest_Priv = 0
        Exit Function
    ElseIf Err.LastDllError <> 0 Then
        Exit Function
    End If
    '//test access level
    lReturn = LookupPrivilegeValue(0&, Access, tpeLuid)
    If lReturn = 0 Then
        cTest_Priv = 0
        Exit Function
    End If
    
    cTest_Priv = True

Handler:

End Function

Public Function cAccess_Level(AccessType As eSEUserRights) As Boolean

Dim sAccessType     As String

On Error GoTo Handler

    '//determine particular access rights
    '//based on se request
    Select Case AccessType
        Case 0
            sAccessType = SE_ASSIGN_PRIMARY_TOKEN
        Case 1
            sAccessType = SE_AUDIT
        Case 2
            sAccessType = SE_BACKUP_NAME
        Case 3
            sAccessType = SE_CHANGE_NOTIFY
        Case 4
            sAccessType = SE_CREATE_PERMANENT
        Case 5
            sAccessType = SE_CREATE_TOKEN
        Case 6
            sAccessType = SE_INCREASE_BASE_PRIORITY
        Case 7
            sAccessType = SE_INCREASE_PAGE_FILE
        Case 8
            sAccessType = SE_INCREASE_QUOTA
        Case 9
            sAccessType = SE_LOAD_DRIVER
        Case 10
            sAccessType = SE_LOCK_MEMORY
        Case 11
            sAccessType = SE_MACHINE_ACCOUNT
        Case 12
            sAccessType = SE_PROFILE_SINGLE_PROC
        Case 13
            sAccessType = SE_REMOTE_SHUTDOWN
        Case 14
            sAccessType = SE_RESTORE_NAME
        Case 15
            sAccessType = SE_SECURITY
        Case 16
            sAccessType = SE_SERVICE_LOGON
        Case 17
            sAccessType = SE_SHUT_DOWN
        Case 18
            sAccessType = SE_SYSTEM_ENVIROMENT
        Case 19
            sAccessType = SE_SYSTEM_PROFILE
        Case 20
            sAccessType = SE_SYSTEM_TIME
        Case 21
            sAccessType = SE_TAKE_OWNERSHIP
        Case 22
            sAccessType = SE_TCB
        Case 23
            sAccessType = SE_UNSOLICITED_INPUT
    End Select
    
    If cTest_Priv(sAccessType) <> 0 Then
        cAccess_Level = True
    End If

Handler:

End Function

Public Function cSave_Key(ByVal KeyPar As Long, _
                          ByVal SubKey As String, _
                          ByVal File As String) As Boolean

'//Save registry key to Snapshot
Dim hKey    As Long
Dim lResult As Long

On Error GoTo Handler

    '//test privilege level
    If cSet_Priv(SE_BACKUP_NAME) <> 0 Then
        '//open key
        RegOpenKeyEx KeyPar, SubKey, 0&, KEY_ALL_ACCESS, hKey
        '//if backup file exists, remove it
        If LenB(Dir(File)) Then Kill File
        '//save the key
        lResult = RegSaveKey(hKey, File, ByVal 0&)
        If lResult = 0 Then cSave_Key = True
        '//close the key
        RegCloseKey hKey
    End If

Handler:

End Function

Public Function cRestore_Key(ByVal KeyPar As Long, _
                             ByVal SubKey As String, _
                             File As String) As Boolean
                           
'//Restore a registry key with Snapshot
Dim hKey        As Long
Dim lResult     As Long

On Error GoTo Handler

    '//test privilege level
    If cSet_Priv(SE_RESTORE_NAME) <> 0 Then
        '//open the key
        RegOpenKeyEx KeyPar, SubKey, 0&, KEY_ALL_ACCESS, hKey
        '//test for file
        If Len(Dir(File)) = 0 Then Exit Function
        '//start restore
        lResult = RegRestoreKey(hKey, File, REG_FORCE_RESTORE)
        If lResult = 0 Then cRestore_Key = True
        '//close the key
        RegCloseKey hKey
    End If

Handler:

End Function


