Attribute VB_Name = "mRestore"
Option Explicit
'//found the system restore api on the net somewhere
'//wrote the rest..

'//type structure for clsMrusearch options
Public Type tChoice
    tChoice0 As Boolean
    tChoice1 As Boolean
    tChoice2 As Boolean
    tChoice3 As Boolean
    tChoice4 As Boolean
    tChoice5 As Boolean
    tChoice6 As Boolean
    tChoice7 As Boolean
    tChoice8 As Boolean
End Type

Private lSeqNum                              As Long
Public bRestore                              As Boolean
Private Const BEGIN_SYSTEM_CHANGE            As Integer = 100
Private Const END_SYSTEM_CHANGE              As Integer = 101
Private Const BEGIN_NESTED_SYSTEM_CHANGE     As Integer = 102
Private Const END_NESTED_SYSTEM_CHANGE       As Integer = 103
Private Const DESKTOP_SETTING                As Integer = 2
Private Const ACCESSIBILITY_SETTING          As Integer = 3
Private Const OE_SETTING                     As Integer = 4
Private Const APPLICATION_RUN                As Integer = 5
Private Const WINDOWS_SHUTDOWN               As Integer = 8
Private Const WINDOWS_BOOT                   As Integer = 9
Private Const MAX_DESC                       As Integer = 64
Private Const MAX_DESC_W                     As Integer = 256
Public Const VER_PLATFORM_WIN32_WINDOWS      As Integer = 1
Public Const VER_PLATFORM_WIN32_NT           As Integer = 2

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize                          As Long
    dwMajorVersion                               As Long
    dwMinorVersion                               As Long
    dwBuildNumber                                As Long
    dwPlatformId                                 As Long
    szCSDVersion                                 As String * 128
    wServicePackMajor                            As Integer
    wServicePackMinor                            As Integer
    wSuiteMask                                   As Integer
    wProductType                                 As Byte
    wReserved                                    As Byte
End Type

Public Enum RestoreType
    APPLICATION_INSTALL = 0
    APPLICATION_UNINSTALL = 1
    MODIFY_SETTINGS = 12
    CANCELLED_OPERATION = 13
    RESTORE = 6
    CHECKPOINT = 7
    DEVICE_DRIVER_INSTALL = 10
    FIRSTRUN = 11
    BACKUP_RECOVERY = 14
End Enum

Private Type RESTOREPTINFOA
    dwEventType                                  As Long
    dwRestorePtType                              As Long
    llSequenceNumber                             As Currency
    szDescription                                As String * MAX_DESC
End Type

Private Type RESTOREPTINFOW
    dwEventType                                  As Long
    dwRestorePtType                              As Long
    llSequenceNumber                             As Currency
    szDescription                                As String * MAX_DESC_W
End Type

Private Type SMGRSTATUS
    nStatus                                      As Long
    llSequenceNumber                             As Currency
End Type

Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpFile As String, _
                                                                            lpSecurityAttributes As Any) As Long

Private Declare Function SRSetRestorePointA Lib "srclient.dll" (pRestorePtSpec As RESTOREPTINFOA, _
                                                                pSMgrStatus As SMGRSTATUS) As Long

Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Public Function End_Restore() As Boolean

'//End the System Restore

Dim tRPI        As RESTOREPTINFOA
Dim tStatus     As SMGRSTATUS
Dim lResult     As Long

    tRPI.dwEventType = END_SYSTEM_CHANGE
    tRPI.llSequenceNumber = lSeqNum
    lResult = SRSetRestorePointA(tRPI, tStatus)
    If lResult = 0 Then End_Restore = True
    lSeqNum = 0


End Function

Public Function Restore_Available() As Boolean

'//Check if OS supports System Restore

Dim OSV As OSVERSIONINFOEX

    OSV.dwOSVersionInfoSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        Select Case OSV.dwPlatformId
        Case VER_PLATFORM_WIN32_WINDOWS
            If OSV.dwMinorVersion = 90 Then    '//windows me
                Restore_Available = True
            End If
        Case VER_PLATFORM_WIN32_NT
            Select Case OSV.dwMajorVersion
            Case 5                              '//2000/xp
                If OSV.dwMinorVersion = 1 Then
                    Restore_Available = True
                End If
            Case Is > 5                         '//future os
                Restore_Available = True
            End Select
        End Select
    End If

End Function

Public Function Start_Restore(ByVal sDescription As String, _
                              Optional ByVal lType As RestoreType = CHECKPOINT) As Boolean

'//Start the System Restore

Dim tRPI    As RESTOREPTINFOA
Dim tStatus As SMGRSTATUS

    If lSeqNum <> 0 Then
        Err.Raise 100001, , "You must End the previous restore point first."
    Else
        DoEvents
        tRPI.dwEventType = BEGIN_SYSTEM_CHANGE
        DoEvents

        With tRPI
            .dwRestorePtType = lType
            .llSequenceNumber = 0
            .szDescription = sDescription
        End With

        If SRSetRestorePointA(tRPI, tStatus) Then
            Start_Restore = True
            lSeqNum = tStatus.llSequenceNumber
        End If
    End If

End Function
