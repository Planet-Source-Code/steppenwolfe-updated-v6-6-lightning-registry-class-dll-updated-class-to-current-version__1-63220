VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReg"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'// Ok, rewrote entire dll to include many more functions.
'// Rewrote project from top to bottom and optimized for speed, efficiency and added error handling
'// Library supports string, dword, binary, multi_sz, expanded string,little endian, big endian, resource list,
'// resource descriptor, resource requirements, qword(64bit dword), and link values.
'// Rewrote the collection controls and changed from public variables to properties.
'// Now includes optional error handling, (returns a long interpreted by class wrapper).
'// Wrapper class includes endian conversion, user rights check, and added a mod with reg ntfs api, enjoy..

'//time structure
Private Type FILETIME
    dwLowDateTime                              As Long
    dwHighDateTime                             As Long
End Type

'//security structure
Private Type SECURITY_ATTRIBUTES
    nLength                                    As Long
    lpSecurityDescriptor                       As Long
    bInheritHandle                             As Boolean
End Type

'//collections
Private mRegKey                            As New cRegKeys
Private mRegValues                         As New cRegValues

'key constants
Public Enum HKEY_Type
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

'//value types
Public Enum Reg_Type
    REG_NONE = 0                            '//No value type
    REG_SZ = 1                              '//Unicode NULL terminated string
    REG_EXPAND_SZ = 2                       '//Unicode NULL terminated string
    REG_BINARY = 3                          '//Binary data
    REG_DWORD = 4                           '//32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4             '//32-bit number
    REG_DWORD_BIG_ENDIAN = 5                '//32-bit number, high byte first
    REG_LINK = 6                            '//Unicode symbolic link
    REG_MULTI_SZ = 7                        '//Array of Unicode strings
    REG_RESOURCE_LIST = 8                   '//Hardware resource description
    REG_FULL_RESOURCE_DESCRIPTOR = 9        '//Hardware resource description
    REG_RESOURCE_REQUIREMENTS_LIST = 10     '//Resource requirements
    REG_QWORD_LITTLE_ENDIAN = 11            '//64bit dword value
End Enum

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

'//error handling
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

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      lpcbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                                                                    ByVal lpValueName As String) As Long

Private Declare Function RegCreatekey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                                                                    ByVal lpSubKey As String, _
                                                                                    ByVal Reserved As Long, _
                                                                                    ByVal lpClass As String, _
                                                                                    ByVal dwOptions As Long, _
                                                                                    ByVal samDesired As Long, _
                                                                                    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                                                    phkResult As Long, _
                                                                                    lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  lpData As Any, _
                                                                                  ByVal cbData As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                                                                                ByVal dwIndex As Long, _
                                                                                ByVal lpName As String, _
                                                                                lpcbName As Long, _
                                                                                lpReserved As Long, _
                                                                                ByVal lpClass As String, _
                                                                                lpcbClass As Long, _
                                                                                lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, _
                                                                                ByVal dwIndex As Long, _
                                                                                ByVal lpValueName As String, _
                                                                                lpcbValueName As Long, _
                                                                                ByVal lpReserved As Long, _
                                                                                lpType As Long, _
                                                                                lpData As Byte, _
                                                                                lpcbData As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)

Public Function Access_Test(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            Optional ByVal lErrNum As Long) As Boolean
'//test user access
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '//security attributes
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//open key
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If RetVal = ERROR_NONE Then
        Access_Test = True
    End If
    
Handler:
lErrNum = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_BEndian(ByVal RootKey As HKEY_Type, _
                              ByVal SubKey As String, _
                              ByVal Value As Variant, _
                              Optional ByVal lErrdesc As Long) As String

'//read an big_endian value
Dim hKey     As Long
Dim RetVal   As Long
Dim sBuffer  As String
Dim slength  As Long

On Error GoTo Handler

    '//open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Exit Function
    End If
    '//create a buffer
    sBuffer = Space$(255)
    slength = 255
    '//query key for string value
    RetVal = RegQueryValueEx(hKey, Value, 0, REG_DWORD_BIG_ENDIAN, ByVal sBuffer, slength)
    '//remove null terminator and add value
    If RetVal = ERROR_NONE Then
        sBuffer = Left$(sBuffer, slength - 1)
        cRead_BEndian = sBuffer
    End If

Handler:
'//close key and set result
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_BEndian(ByVal RootKey As HKEY_Type, _
                               ByVal SubKey As String, _
                               ByVal Value As String, _
                               ByVal vData As Variant) As Long

'//write a big endian value (32b number)
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler
    
    '//required security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//open key and test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_BEndian = RetVal
        Exit Function
    End If
    '//write value, set error, and close key
    RetVal = RegSetValueEx(hKey, Value, 0&, REG_DWORD_BIG_ENDIAN, vData, 4)

Handler:
cWrite_BEndian = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_Binary(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal Value As Variant, _
                             Optional ByVal lErrdesc As Long) As Variant

'//read a binary value
Dim hKey            As Long
Dim RetVal          As Long
Dim byBuffer()      As Byte
Dim lbuffersize     As Long
Dim sTemp           As String
Dim i               As Long

On Error GoTo Handler

    '//open root key and test for value type
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Exit Function
    End If

    '//get buffer size
    RetVal = RegQueryValueEx(hKey, Value, 0&, REG_BINARY, ByVal 0&, lbuffersize)
    '//read into buffer
    If RetVal = ERROR_NONE Then
        ReDim byBuffer(lbuffersize - 1) As Byte
        RetVal = RegQueryValueEx(hKey, Value, 0&, REG_BINARY, byBuffer(0), lbuffersize)
        '//format value
        For i = 0 To UBound(byBuffer)
            sTemp = sTemp & Format(Trim(Hex(byBuffer(i))), "0#")
        Next i
        '//set value
        cRead_Binary = LTrim(sTemp)
    Else
        lErrdesc = RetVal
        Exit Function
    End If

Handler:
'//set error and close key
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_Binary(ByVal RootKey As HKEY_Type, _
                              ByVal SubKey As String, _
                              ByVal Value As Variant, _
                              ByRef bData() As Byte) As Long


'//write a binary value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '//security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_Binary = RetVal
        Exit Function
    End If
    '//write binary, set error, and close
    RetVal = RegSetValueEx(hKey, Value, 0&, REG_BINARY, bData(0), UBound(bData) + 1)

Handler:
cWrite_Binary = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_DWord(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal Value As Variant, _
                            Optional ByVal lErrdesc As Long) As Long

'//read a dword value
Dim hKey     As Long
Dim RetVal   As Long
Dim lBuffer  As Long

On Error GoTo Handler

    '//open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
    End If
    '//query value
    RetVal = RegQueryValueEx(hKey, Value, 0, REG_DWORD, lBuffer, 4)
    '//set value
    If RetVal = ERROR_NONE Then
        cRead_DWord = lBuffer
    Else
        cRead_DWord = 0
    End If

Handler:
'//close key
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_DWord(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal Value As String, _
                             ByVal lData As Long) As Long

'//write a DWORD value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler
    
    '//required security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//open key and test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_DWord = RetVal
        Exit Function
    End If
    '//write value, set error, and close key
    RetVal = RegSetValueEx(hKey, Value, 0&, REG_DWORD, lData, 4)

Handler:
cWrite_DWord = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_Link(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String, _
                           ByVal Value As Variant, _
                           Optional ByVal lErrdesc As Long) As Variant

'//read a link value
Dim hKey          As Long
Dim RetVal        As Long
Dim byBuffer()    As Byte
Dim lbuffersize   As Long
Dim byDefArray(0) As Byte

On Error GoTo Handler

    '//dimension byte array
    byDefArray(0) = 0
    '//open root key and test for value type
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Exit Function
    End If
    '//get buffer size
    RetVal = RegQueryValueEx(hKey, Value, 0&, REG_LINK, ByVal 0&, lbuffersize)
    '//read into buffer
    If RetVal = ERROR_NONE Then
        ReDim byBuffer(lbuffersize - 1) As Byte
        RetVal = RegQueryValueEx(hKey, Value, 0&, REG_LINK, byBuffer(0), lbuffersize)
        '//set value
        cRead_Link = byBuffer
    Else
        lErrdesc = RetVal
        Exit Function
    End If

Handler:
'//set error and close key
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_Link(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal Value As Variant, _
                            ByRef bData() As Byte) As Long

'//write a link value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '//security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_Link = RetVal
        Exit Function
    End If
    '//write link, set error, and close
    RetVal = RegSetValueEx(hKey, Value, 0&, REG_LINK, bData(0), UBound(bData) + 1)

Handler:
cWrite_Link = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_LEndian(ByVal RootKey As HKEY_Type, _
                              ByVal SubKey As String, _
                              ByVal Value As Variant, _
                              Optional ByVal lErrdesc As Long) As Long

'//read a dword value
Dim hKey     As Long
Dim RetVal   As Long
Dim lBuffer  As Long

On Error GoTo Handler

    '//open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
    End If
    '//query value
    RetVal = RegQueryValueEx(hKey, Value, 0, REG_DWORD_LITTLE_ENDIAN, lBuffer, 4)
    '//set value
    If RetVal = ERROR_NONE Then
        cRead_LEndian = lBuffer
    Else
        cRead_LEndian = 0
    End If

Handler:
'//close key
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_LEndian(ByVal RootKey As HKEY_Type, _
                               ByVal SubKey As String, _
                               ByVal Value As String, _
                               ByVal vData As Variant) As Long

'//write a little endian value (same as dword)
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler
    
    '//required security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//open key and test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_LEndian = RetVal
        Exit Function
    End If
    '//write value, set error, and close key
    RetVal = RegSetValueEx(hKey, Value, 0&, REG_DWORD_LITTLE_ENDIAN, vData, 4)

Handler:
cWrite_LEndian = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_Multi(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal Value As String, _
                            Optional ByVal lErrdesc As Long) As String

'//read a multi_sz value
Dim hKey        As Long
Dim RetVal      As Long
Dim sBuffer     As String
Dim Length      As Long
Dim resBinary() As Byte
Dim resString   As String

On Error GoTo Handler
    
    '//open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Exit Function
    End If
    '//get buffer size
    Length = 1024
    ReDim resBinary(0 To Length - 1) As Byte
    RetVal = RegQueryValueEx(hKey, Value, 0, REG_MULTI_SZ, resBinary(0), Length)
    '//if more data, resize our byte array
    If RetVal = ERROR_MORE_DATA Then
        ReDim resBinary(0 To Length - 1) As Byte
        RetVal = RegQueryValueEx(hKey, Value, 0, REG_MULTI_SZ, resBinary(0), Length)
    End If
    '//read into buffer
    If RetVal = ERROR_NONE Then
        resString = Space$(Length - 2)
        CopyMemory ByVal resString, resBinary(0), Length - 2
        sBuffer = resString
        If Len(TrimNull(sBuffer)) > 0 Then
            cRead_Multi = resString
        End If
    Else
        lErrdesc = RetVal
    End If

Handler:
'//set error and close key
lErrdesc = 0
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_Multi(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal Value As String, _
                             ByVal sData As String) As Long

'//write a multi+sz value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '//security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_Multi = RetVal
        Exit Function
    End If
    '//write multi, set error, and close key
    RetVal = RegSetValueEx(hKey, Value, 0, REG_MULTI_SZ, ByVal sData, Len(sData))

Handler:
cWrite_Multi = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_QWord(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal Value As Variant, _
                            Optional ByVal lErrdesc As Long) As Currency

'//read a qword value
Dim hKey     As Long
Dim RetVal   As Long
Dim lBuffer  As Long

On Error GoTo Handler

    '//open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
    End If
    '//query value
    RetVal = RegQueryValueEx(hKey, Value, 0, REG_QWORD_LITTLE_ENDIAN, lBuffer, 4)
    '//set value
    If RetVal = ERROR_NONE Then
        cRead_QWord = lBuffer
    Else
        cRead_QWord = 0
    End If

Handler:
'//close key
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_QWord(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal Value As String, _
                             ByVal cData As Currency) As Long

'//write a DWORD value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler
    
    '//required security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//open key and test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_QWord = RetVal
        Exit Function
    End If
    '//write value, set error, and close key
    RetVal = RegSetValueEx(hKey, Value, 0&, REG_QWORD_LITTLE_ENDIAN, cData, 4)

Handler:
cWrite_QWord = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_ResourceList(ByVal RootKey As HKEY_Type, _
                                   ByVal SubKey As String, _
                                   ByVal Value As Variant, _
                                   Optional ByVal lErrdesc As Long) As Variant

'//read a resource list value(binary)
Dim hKey          As Long
Dim RetVal        As Long
Dim byBuffer()    As Byte
Dim lbuffersize   As Long
Dim byDefArray(0) As Byte

On Error GoTo Handler

    '//dimension byte array
    byDefArray(0) = 0
    '//open root key and test for value type
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Exit Function
    End If
    '//get buffer size
    RetVal = RegQueryValueEx(hKey, Value, 0&, REG_RESOURCE_LIST, ByVal 0&, lbuffersize)
    '//read into buffer
    If RetVal = ERROR_NONE Then
        ReDim byBuffer(lbuffersize - 1) As Byte
        RetVal = RegQueryValueEx(hKey, Value, 0&, REG_RESOURCE_LIST, byBuffer(0), lbuffersize)
        '//set value
        cRead_ResourceList = byBuffer
    Else
        lErrdesc = RetVal
        Exit Function
    End If

Handler:
'//set error and close key
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_ResourceList(ByVal RootKey As HKEY_Type, _
                                    ByVal SubKey As String, _
                                    ByVal Value As Variant, _
                                    ByRef bData() As Byte) As Long


'//write a resource list value(binary)
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '//security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_ResourceList = RetVal
        Exit Function
    End If
    '//write resource, set error, and close
    RetVal = RegSetValueEx(hKey, Value, 0&, REG_RESOURCE_LIST, bData(0), UBound(bData) + 1)

Handler:
cWrite_ResourceList = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_ResDescriptor(ByVal RootKey As HKEY_Type, _
                                    ByVal SubKey As String, _
                                    ByVal Value As Variant, _
                                    Optional ByVal lErrdesc As Long) As Variant

'//read a resource list value(binary)
Dim hKey          As Long
Dim RetVal        As Long
Dim byBuffer()    As Byte
Dim lbuffersize   As Long
Dim byDefArray(0) As Byte

On Error GoTo Handler

    '//dimension byte array
    byDefArray(0) = 0
    '//open root key and test for value type
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Exit Function
    End If
    '//get buffer size
    RetVal = RegQueryValueEx(hKey, Value, 0&, REG_FULL_RESOURCE_DESCRIPTOR, ByVal 0&, lbuffersize)
    '//read into buffer
    If RetVal = ERROR_NONE Then
        ReDim byBuffer(lbuffersize - 1) As Byte
        RetVal = RegQueryValueEx(hKey, Value, 0&, REG_FULL_RESOURCE_DESCRIPTOR, byBuffer(0), lbuffersize)
        '//set value
        cRead_ResDescriptor = byBuffer
    Else
        lErrdesc = RetVal
        Exit Function
    End If

Handler:
'//set error and close key
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_ResDescriptor(ByVal RootKey As HKEY_Type, _
                                     ByVal SubKey As String, _
                                     ByVal Value As Variant, _
                                     ByRef bData() As Byte) As Long




'//write a resource list value(binary)
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '//security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_ResDescriptor = RetVal
        Exit Function
    End If
    '//write resource, set error, and close
    RetVal = RegSetValueEx(hKey, Value, 0&, REG_FULL_RESOURCE_DESCRIPTOR, bData(0), UBound(bData) + 1)

Handler:
cWrite_ResDescriptor = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_ResRequired(ByVal RootKey As HKEY_Type, _
                                  ByVal SubKey As String, _
                                  ByVal Value As Variant, _
                                  Optional ByVal lErrdesc As Long) As Variant

'//read a resource requirements value(binary)
Dim hKey          As Long
Dim RetVal        As Long
Dim byBuffer()    As Byte
Dim lbuffersize   As Long
Dim byDefArray(0) As Byte

On Error GoTo Handler

    '//dimension byte array
    byDefArray(0) = 0
    '//open root key and test for value type
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Exit Function
    End If
    '//get buffer size
    RetVal = RegQueryValueEx(hKey, Value, 0&, REG_RESOURCE_REQUIREMENTS_LIST, ByVal 0&, lbuffersize)
    '//read into buffer
    If RetVal = ERROR_NONE Then
        ReDim byBuffer(lbuffersize - 1) As Byte
        RetVal = RegQueryValueEx(hKey, Value, 0&, REG_RESOURCE_REQUIREMENTS_LIST, byBuffer(0), lbuffersize)
        '//set value
        cRead_ResRequired = byBuffer
    Else
        lErrdesc = RetVal
        Exit Function
    End If

Handler:
'//set error and close key
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_ResRequired(ByVal RootKey As HKEY_Type, _
                                   ByVal SubKey As String, _
                                   ByVal Value As Variant, _
                                   ByRef bData() As Byte) As Long


'//write a resource requirements value(binary)
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '//security structure
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//test access
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_ResRequired = RetVal
        Exit Function
    End If
    '//write resource, set error, and close
    RetVal = RegSetValueEx(hKey, Value, 0&, REG_RESOURCE_REQUIREMENTS_LIST, bData(0), UBound(bData) + 1)

Handler:
cWrite_ResRequired = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cRead_String(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal Value As String, _
                             Optional ByVal lErrdesc As Long) As String

'//read an SZ value
Dim hKey     As Long
Dim RetVal   As Long
Dim sBuffer  As String
Dim slength  As Long

On Error GoTo Handler

    '//open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Exit Function
    End If
    '//create a buffer
    sBuffer = Space$(255)
    slength = 255
    '//query key for string value
    RetVal = RegQueryValueEx(hKey, Value, 0, REG_SZ, ByVal sBuffer, slength)
    '//read string into buffer
    If RetVal = ERROR_NONE Then
        sBuffer = Left$(sBuffer, slength - 1)
        cRead_String = sBuffer
    End If

Handler:
'//close key and set result
lErrdesc = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0
    
End Function

Public Function cWrite_String(ByVal RootKey As HKEY_Type, _
                              ByVal SubKey As String, _
                              ByVal Value As String, _
                              ByVal Data As String) As Long

'//write an SZ value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '//security attributes
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//open key
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_String = RetVal
        Exit Function
    End If
    '//write string, set error, and close
    RetVal = RegSetValueEx(hKey, Value, 0, REG_SZ, ByVal Data, Len(Data))

Handler:
cWrite_String = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cWrite_String_Expand(ByVal RootKey As HKEY_Type, _
                                     ByVal SubKey As String, _
                                     ByVal Value As String, _
                                     ByVal sData As String) As Long

'//write an SZ value
Dim hKey    As Long
Dim RetVal  As Long
Dim deposit As Long
Dim secattr As SECURITY_ATTRIBUTES

On Error GoTo Handler

    '//security attributes
    With secattr
        .nLength = Len(secattr)
        .lpSecurityDescriptor = 0
        .bInheritHandle = 1
    End With
    '//open key
    RetVal = RegCreateKeyEx(RootKey, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If Not RetVal = ERROR_NONE Then
        cWrite_String_Expand = RetVal
        Exit Function
    End If
    '//write string, set error, and close
    RetVal = RegSetValueEx(hKey, Value, 0, REG_EXPAND_SZ, ByVal sData, Len(sData))

Handler:
cWrite_String_Expand = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cKey_Exist(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String) As Boolean

'//check for key
Dim hKey   As Long
Dim RetVal As Long

On Error GoTo Handler
    '//if 0 returned, key is valid
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_QUERY_VALUE, hKey)
    If RetVal = ERROR_NONE Then
        cKey_Exist = True
    End If

Handler:
RegCloseKey hKey
On Error GoTo 0

End Function

Public Function cValue_Exist(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal Value As String) As Boolean


'//read an SZ value
Dim hKey     As Long
Dim RetVal   As Long
Dim sBuffer  As String
Dim slength  As Long
Dim DataType As Long

On Error GoTo Handler

    '//open root key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If Not RetVal = ERROR_NONE Then
        Exit Function
    End If
    '//create a buffer
    sBuffer = Space$(255)
    slength = 255
    '//query key for string value
    RetVal = RegQueryValueEx(hKey, Value, 0, DataType, ByVal sBuffer, slength)
    '//read string into buffer
    If RetVal = ERROR_NONE Then
        cValue_Exist = True
    End If

Handler:
'//close key and set result
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cCreate_Key(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String) As Long

'//create a new key
Dim hKey   As Long
Dim RetVal As Long

On Error GoTo Handler
    '//create the key
    RetVal = RegCreatekey(RootKey, SubKey, hKey)

Handler:
'//set error and close
cCreate_Key = RetVal
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cDelete_Key(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String) As Long

'//delete a key
Dim RetVal  As Long

On Error GoTo Handler
    
    '//delete key
    RetVal = RegDeleteKey(RootKey, SubKey)
    '//set error
    cDelete_Key = RetVal

Handler:
On Error GoTo 0

End Function

Public Function cDelete_Value(ByVal RootKey As HKEY_Type, _
                              ByVal SubKey As String, _
                              ByVal Value As Variant) As Long

'//delete a value
Dim RetVal As Long
Dim handle As Long

On Error GoTo Handler

    '//open key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ALL_ACCESS, handle)
    '//exit on error
    If Not RetVal = ERROR_NONE Then
        cDelete_Value = RetVal
        Exit Function
    End If
    '//delete value
    RetVal = RegDeleteValue(handle, Value)

Handler:
'//set error and close key
cDelete_Value = RetVal
RetVal = RegCloseKey(handle)
On Error GoTo 0

End Function

Public Function cList_Keys(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String, _
                           Optional ByVal lErrdesc As Long) As cRegKeys

'//list all keys and add to collection
Dim KeyName   As String
Dim keylen    As Long
Dim classname As String
Dim classlen  As Long
Dim lastwrite As FILETIME
Dim hKey      As Long
Dim RetVal    As Long
Dim Index     As Long

On Error GoTo Handler

    '//open key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
    If Not RetVal = ERROR_NONE Then
        Set cList_Keys = Nothing
        lErrdesc = RetVal
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
            mRegKey.Add KeyName
        End If
        Index = Index + 1
    Loop Until Not RetVal = 0
    lErrdesc = 0

Handler:
'//set collection and close
Set cList_Keys = mRegKey
RetVal = RegCloseKey(hKey)
On Error GoTo 0

End Function

Public Function cColList_Keys(ByVal RootKey As HKEY_Type, _
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
Dim cTemp     As New Collection

On Error GoTo Handler

    Set cColList_Keys = New Collection
    '//open key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
    If Not RetVal = ERROR_NONE Then
        Set cColList_Keys = Nothing
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
            cColList_Keys.Add KeyName
        End If
        Index = Index + 1
    Loop Until RetVal <> 0

Handler:
'//set error, collection and close
Set cTemp = Nothing
RetVal = RegCloseKey(hKey)
On Error GoTo 0
    
End Function

Public Function cList_Values(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             Optional ByVal lErrdesc As Long) As cRegValues

'//list all values and add to a collection
Dim BinaryString   As Integer
Dim Value          As String
Dim ValueLen       As Long
Dim DataType       As Long
Dim Data(0 To 254) As Byte
Dim DataLen        As Long
Dim hKey           As Long
Dim Index          As Long
Dim RetVal         As Long

On Error GoTo Handler

    '//open key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_QUERY_VALUE, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Set cList_Values = Nothing
        Exit Function
    End If
    Index = 0
    '//loop through values and add to collection
    Do
        DataLen = 255
        Value = Space$(255)
        ValueLen = 255
        BinaryString = 0
        RetVal = RegEnumValue(hKey, Index, Value, ValueLen, 0, DataType, Data(0), DataLen)
        If RetVal = ERROR_NONE Then
            Value = Left$(Value, ValueLen)
            Select Case DataType
                Case REG_SZ, REG_EXPAND_SZ
                    mRegValues.Add Value, DataType, cRead_String(RootKey, SubKey, Value)
                Case REG_MULTI_SZ
                    mRegValues.Add Value, DataType, cRead_Multi(RootKey, SubKey, Value)
                Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
                    mRegValues.Add Value, DataType, cRead_DWord(RootKey, SubKey, Value)
                Case REG_BINARY
                    mRegValues.Add Value, DataType, cRead_Binary(RootKey, SubKey, Value)
                Case Else
                    mRegValues.Add Value, DataType, Nothing
            End Select
        End If
        Index = Index + 1
    Loop Until Not RetVal = 0
    
    '//set error, collection, and close key
    lErrdesc = 0
    Set cList_Values = mRegValues
    RetVal = RegCloseKey(hKey)

Exit Function
Handler:
lErrdesc = RetVal
On Error GoTo 0

End Function

Public Function cSearch_Values(ByVal RootKey As HKEY_Type, _
                               ByVal SubKey As String, _
                               Optional ByVal lErrdesc As Long) As Collection

'//list all values and add to a collection
Dim Value          As String
Dim ValueLen       As Long
Dim DataType       As Long
Dim Data(0 To 254) As Byte
Dim DataLen        As Long
Dim hKey           As Long
Dim Index          As Long
Dim RetVal         As Long
Dim cTemp          As New Collection

On Error GoTo Handler

    Set cTemp = New Collection
    '//open key
    RetVal = RegOpenKeyEx(RootKey, SubKey, 0, KEY_QUERY_VALUE, hKey)
    If Not RetVal = ERROR_NONE Then
        lErrdesc = RetVal
        Set cTemp = Nothing
        Exit Function
    End If
    Index = 0
    '//loop through values and add to collection
    Do
        DataLen = 255
        Value = Space$(255)
        ValueLen = 255
        RetVal = RegEnumValue(hKey, Index, Value, ValueLen, 0, DataType, Data(0), DataLen)
        If RetVal = ERROR_NONE Then
            Value = Left$(Value, ValueLen)
            cTemp.Add Value
        End If
        Index = Index + 1
    Loop Until RetVal <> 0

    '//set error, collection, and close key
    lErrdesc = 0
    Set cSearch_Values = cTemp
    Set cTemp = Nothing
    RetVal = RegCloseKey(hKey)


Exit Function

Handler:
    lErrdesc = RetVal
    On Error GoTo 0

End Function

Public Function TrimNull(Item As String) As String
'//trim nulls
Dim pos As Integer

On Error GoTo Handler
    '//trim nulls for sz_multi
    pos = InStr(Item, Chr$(0))
    If pos Then
        Item = Left$(Item, pos - 1)
    End If
    TrimNull = Item

Handler:
On Error GoTo 0

End Function
