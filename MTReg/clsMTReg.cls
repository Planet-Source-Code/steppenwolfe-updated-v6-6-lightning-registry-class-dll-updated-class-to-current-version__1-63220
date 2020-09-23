VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMTReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'// Recursion routine using async thread operations.
'// Demonstrates potential performance gain by releasing
'// classes from parent object dependancy
'// Could, (with a little work and imagination..)
'// be added into the main dll, and used to generate
'// multiple in process worker threads for the recursion
'// routine, rather then just the one async thread

Implements Runnable

Public Enum HKEY_Type
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Private Const ERROR_NONE                   As Long = &H0
Private Const ERROR_BADDB                  As Long = &H1
Private Const ERROR_BADKEY                 As Long = &H2
Private Const ERROR_CANTOPEN               As Long = &H3
Private Const ERROR_CANTREAD               As Long = &H4
Private Const ERROR_CANTWRITE              As Long = &H5
Private Const ERROR_OUTOFMEMORY            As Long = &H6
Private Const ERROR_ARENA_TRASHED          As Long = &H7
Private Const ERROR_ACCESS_DENIED          As Long = &H8
Private Const ERROR_INVALID_PARAMETERS     As Long = &H57
Private Const ERROR_MORE_DATA              As Long = &HEA
Private Const ERROR_NO_MORE_ITEMS          As Long = &H103

'//access paramaters
Private Const KEY_ALL_ACCESS               As Long = &HF003F
Private Const KEY_CREATE_LINK              As Long = &H20
Private Const KEY_CREATE_SUB_KEY           As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS       As Long = &H8
Private Const KEY_EXECUTE                  As Long = &H20019
Private Const KEY_NOTIFY                   As Long = &H10
Private Const KEY_QUERY_VALUE              As Long = &H1
Private Const KEY_READ                     As Long = &H20019
Private Const KEY_SET_VALUE                As Long = &H2
Private Const KEY_WRITE                    As Long = &H20006
Private Const REG_OPTION_NON_VOLATILE      As Long = &H0
Private Const REG_ERR_OK                   As Long = &H0
Private Const REG_ERR_NOT_EXIST            As Long = &H1
Private Const REG_ERR_NOT_STRING           As Long = &H2
Private Const REG_ERR_NOT_DWORD            As Long = &H4

'//time structure
Private Type FILETIME
    dwLowDateTime                          As Long
    dwHighDateTime                         As Long
End Type

'//security structure
Private Type SECURITY_ATTRIBUTES
    nLength                                As Long
    lpSecurityDescriptor                   As Long
    bInheritHandle                         As Boolean
End Type

Private cLKeyList                          As New Collection
Private m_cGKeyList                        As New Collection

Public Event Complete()
Public Event Cancelled()

Private m_lRoot                            As Long
Private m_sSubKey                          As String
Private m_lInterval                        As Long
Private m_bRunning                         As Boolean

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
                                                                                
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                                                                                ByVal dwIndex As Long, _
                                                                                ByVal lpName As String, _
                                                                                lpcbName As Long, _
                                                                                lpReserved As Long, _
                                                                                ByVal lpClass As String, _
                                                                                lpcbClass As Long, _
                                                                                lpftLastWriteTime As FILETIME) As Long

Public Property Get cGKeyList() As Collection
    Set cGKeyList = m_cGKeyList
End Property

Public Property Set cGKeyList(PropVal As Collection)
    Set m_cGKeyList = PropVal
End Property

Public Property Get Root() As HKEY_Type
    Root = m_lRoot
End Property

Public Property Let Root(ByVal lRoot As HKEY_Type)
    m_lRoot = lRoot
End Property

Public Property Get SubKey() As String
    SubKey = m_sSubKey
End Property

Public Property Let SubKey(ByVal sSubKey As String)
    m_sSubKey = sSubKey
End Property

Public Property Get Interval() As Long
    Interval = m_lInterval
End Property

Public Property Let Interval(ByVal lInterval As Long)
    m_lInterval = lInterval
End Property

Public Sub Start()

    If Not m_bRunning Then
        m_bRunning = True
        mMain.Start Me
    End If

End Sub

Private Sub Runnable_Start()

Dim cRoot   As Long
Dim cSubKey As String
Dim vKey    As Variant

    '//key values
    cSubKey = SubKey
    cRoot = Root
    '//call recursion loop
    Lib_Recurse_Keys cRoot, cSubKey
    
    For Each vKey In cLKeyList
        cGKeyList.Add vKey
    Next vKey

    RaiseEvent Complete
    m_bRunning = False

End Sub

Private Function Lib_List_Keys(ByVal lHKey As HKEY_Type, _
                               ByVal SubKey As String) As Collection

'//list all keys and add to collection
Dim KeyName   As String
Dim keylen    As Long
Dim classname As String
Dim classlen  As Long
Dim lastwrite As FILETIME
Dim hKey      As Long
Dim RetVal    As Long
Dim Index     As Long
Dim cKeyList  As New Collection

    On Error GoTo Handler

    Set cKeyList = New Collection
    '//open key
    RetVal = RegOpenKeyEx(lHKey, SubKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
    If Not RetVal = ERROR_NONE Then
        Set cKeyList = Nothing
        Exit Function
    End If
    Index = 0
    
    '//loop through keys and add to collection
    Do
        KeyName = Space$(255)
        keylen = 255
        classname = Space$(255)
        classlen = 255
        RetVal = RegEnumKeyEx(hKey, Index, KeyName, keylen, ByVal 0, classname, classlen, lastwrite)
        If RetVal = ERROR_NONE Then
            KeyName = Left$(KeyName, keylen)
            cKeyList.Add KeyName
        End If
        Index = Index + 1
    Loop Until Not RetVal = 0
    '//close
    Set Lib_List_Keys = cKeyList
    Set cKeyList = Nothing

Handler:
    RetVal = RegCloseKey(hKey)

End Function

Public Function Lib_Recurse_Keys(ByVal lHKey As HKEY_Type, _
                                 ByVal sSubKey As String)

Dim cKey     As Variant

On Error GoTo Handler

    For Each cKey In Lib_List_Keys(lHKey, sSubKey)
        If Not IsEmpty(cKey) Then
            cLKeyList.Add sSubKey & Chr$(92) & cKey
            Lib_Recurse_Keys lHKey, sSubKey & Chr$(92) & cKey
        End If
    Next cKey

Handler:
    On Error GoTo 0

End Function

Private Sub Class_Initialize()
    Set cLKeyList = New Collection
    Set m_cGKeyList = New Collection
End Sub

Private Sub Class_Terminate()
    Set cLKeyList = Nothing
End Sub


