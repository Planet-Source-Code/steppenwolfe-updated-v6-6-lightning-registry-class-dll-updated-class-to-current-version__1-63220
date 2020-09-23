Attribute VB_Name = "mBenchMark"
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

Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_RESOURCE_LIST = 8

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
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      lpcbData As Long) As Long
                                                                                      
Public lModKeyCount     As Long
Private cMKeyList As New Collection

Private lLibKeyCount As Long
Private cLKeyList As New Collection

Public Function Mod_List_Keys(ByVal Key As HKEY_Type, _
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
    RetVal = RegOpenKeyEx(Key, SubKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
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
    Set Mod_List_Keys = cKeyList
    Set cKeyList = Nothing
    RetVal = RegCloseKey(hKey)

Handler:

End Function

Public Sub Mod_Recurse_Keys(ByVal lHKey As Long, _
                            ByVal sSubKey As String)

Dim cKey        As Variant

On Error GoTo Handler

    For Each cKey In Mod_List_Keys(lHKey, sSubKey)
        If Not IsEmpty(cKey) Then
            cMKeyList.Add sSubKey & Chr(92) & cKey
            Mod_Recurse_Keys lHKey, sSubKey & Chr(92) & cKey
        End If
    Next cKey
    
lModKeyCount = cMKeyList.Count

Handler:
'Set cMKeyList = Nothing
On Error GoTo 0

End Sub

Public Function ReadString(lHKey As Long, _
                           SubKey As String, _
                           sName As String) As String
    
Dim hKey As Long
Dim RetVal As Long
Dim sBuffer As String
Dim slength As Long
Dim DataType As Long

On Error Resume Next

    RetVal = RegOpenKeyEx(lHKey, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If RetVal <> ERROR_NONE Then
        Exit Function
    End If
    
    sBuffer = Space(255)
    slength = 255
    RetVal = RegQueryValueEx(hKey, sName, 0, DataType, ByVal sBuffer, slength)
    
    If RetVal = ERROR_NONE Then
        If DataType = REG_SZ Or DataType = REG_EXPAND_SZ Then
            sBuffer = Left(sBuffer, slength - 1)
            ReadString = sBuffer
        End If
    End If
    
    RetVal = RegCloseKey(hKey)
    
On Error GoTo 0

End Function

