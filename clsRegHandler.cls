VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'***************************************************************************************
'*  CONTROL WRAPPER clsRegHandler ->> (nsprcl.dll)                                     *
'*                                                                                     *
'*  Created:     November 10, 2005                                                     *
'*  Purpose:     Interface to registry control library (nsprcl.dll)                    *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.4                                                                   *
'*  Compile:     PCode                                                                 *
'*  Referenced:  clsMruSearch, frmSearch, frmRegDemo                                   *
'*  Author:      John Underhill                                                        *
'*                                                                                     *
'***************************************************************************************

'** List of exposed functions **

'** Value Types **
'// Access_Check - Test user access rights
'// Read_BEndian - read a big endian value
'// Write_BEndian - write a big_endian value
'// Read_Binary - read a binary value
'// Write_Binary - write a binary value
'// Read_Dword - read a dword value
'// Write_Dword - write a dword value
'// Read_Link - read a binary link value
'// Write_Link - write a binary link value
'// List_Values - puts all of a keys values into a collection
'// Read_LEndian - read a little endian value
'// Write_LEndian - write a little_endian value
'// Read_Multi - read a multi_sz value
'// Write_Multi - write a multi_sz value
'// Read_MultiCN - reads a multi value and splits into a collection
'// Write_MultiCN - converts a collection into a multi_sz value
'// Write_Qword - write a 64bit dword number
'// Read_Qword - read a 64bit dword number
'// Read_ResDesc - read hardware resource description (hex)binary
'// Write_ResDesc = write hardware resource description (hex)binary
'// Read_ResList - read a hardware resource list (hex)binary
'// Write_ResList - write to a hardware resource list (hex)binary
'// Read_ResRequired - read a hardware resource requirements list (hex)binary
'// Write_ResRequired - write to a hardware resource requirements list (hex)binary
'// Read_String - read a string(sz) or expanded string(expand_sz)
'// Write_String - write a string value
'// Write_StringEX - write an expanded string value

'** Key Related Functions **
'// List_Keys - puts all subkeys under specified branch into a collection
'// Key_Exists - test if key exists
'// Create_Key - create a new key
'// Delete_Key - delete a key

'** Value Related Functions **
'// Write_Value - write value types: 1)sz 2)expand_sz 3)multi_sz 4)binary 5)dword 6)little_endian 7)big_endian
'// Delete_Value - delete a value
'// Search_Value - search for a value under the key

'** Conversion Routines **
'// Make_LEndian16 - convert integer to 16bit little_endian
'// Make_LEndian32 - convert long to 32bit little_endian
'// Make_BEndian32 - convert long to big endian format

'** Error Logging **
'// Get_Error - interprets errors passed from dll
'// Log_Error - sends errors to a log file

'** Security Functions **
'// Test_Access - test users level of access permissions
'// Set_Key_Permissions - change access permissions to a key

'** Backup and Restore **
'// Save_Key - save a binary image of a registry key
'// Restore_Key - restore key from binary image




'//root keys
Public Enum HKEY_Type
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

'//Generic Access Rights
Private Const GENERIC_ALL                         As Long = &H10000000
Private Const GENERIC_READ                        As Long = &H80000000
Private Const GENERIC_EXECUTE                     As Long = &H20000000
Private Const GENERIC_WRITE                       As Long = &H40000000

'//Standard Access Rights
Private Const DELETE                              As Long = &H10000
Private Const READ_CONTROL                        As Long = &H20000
Private Const WRITE_DAC                           As Long = &H40000
Private Const WRITE_OWNER                         As Long = &H80000
Private Const SYNCHRONIZE                         As Long = &H100000
Private Const STANDARD_RIGHTS_REQUIRED            As Long = &HF0000
Private Const STANDARD_RIGHTS_READ                As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE               As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE             As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL                 As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL                 As Long = &HFFFF
Private Const ACCESS_SYSTEM_SECURITY              As Long = &H1000000
Private Const MAXIMUM_ALLOWED                     As Long = &H2000000

'//inherit flags of an Ace header.
Private Const OBJECT_INHERIT_ACE                  As Long = &H1
Private Const CONTAINER_INHERIT_ACE               As Long = &H2
Private Const NO_PROPAGATE_INHERIT_ACE            As Long = &H4
Private Const INHERIT_ONLY_ACE                    As Long = &H8
Private Const INHERITED_ACE                       As Long = &H10
Private Const VALID_INHERIT_FLAGS                 As Long = &H1F

'//security descriptor flags.
Private Const SE_DACL_AUTO_INHERIT_REQ            As Long = &H100
Private Const SE_SACL_AUTO_INHERIT_REQ            As Long = &H200
Private Const SE_DACL_AUTO_INHERITED              As Long = &H400
Private Const SE_SACL_AUTO_INHERITED              As Long = &H800
Private Const SE_DACL_PROTECTED                   As Long = &H1000
Private Const SE_SACL_PROTECTED                   As Long = &H2000

'//ACE being added.
Private Const ACCESS_ALLOWED_ACE_TYPE             As Integer = 0
Private Const ACCESS_DENIED_ACE_TYPE              As Integer = 1

'//well-known SIDs, users and groups
Private Const SECURITY_WORLD_SID_AUTHORITY        As Long = &H1
Private Const SECURITY_NT_AUTHORITY               As Long = &H5
Private Const SECURITY_BUILTIN_DOMAIN_RID         As Long = &H20
Private Const DOMAIN_ALIAS_RID_ADMINS             As Long = &H220
Private Const DOMAIN_ALIAS_RID_USERS              As Long = &H221
Private Const SECURITY_LOCAL_SYSTEM_RID           As Long = &H12
Private Const SECURITY_WORLD_RID                  As Long = &H0
Private Const DOMAIN_USER_RID_ADMIN               As Long = &H1F4
Private Const DOMAIN_USER_RID_GUEST               As Long = &H1F5
Private Const DOMAIN_GROUP_RID_ADMINS             As Long = &H200
Private Const INVALID_HANDLE_VALUE                As Integer = -1
Private Const OPEN_EXISTING                       As Integer = 3
Private Const FILE_FLAG_BACKUP_SEMANTICS          As Long = &H2000000

'//API Constants
Private Const DACL_SECURITY_INFORMATION           As Long = &H4
Private Const SECURITY_DESCRIPTOR_REVISION        As Integer = 1
Private Const SECURITY_DESCRIPTOR_MIN_LENGTH      As Integer = 20
Private Const SD_SIZE                             As Double = (65536 + SECURITY_DESCRIPTOR_MIN_LENGTH)
Private Const ACL_REVISION2                       As Integer = 2
Private Const ACL_REVISION                        As Integer = 2
Private Const MAXDWORD                            As Long = &HFFFFFFFF
Private Const SidTypeUser                         As Integer = 1
Private Const AclSizeInformation                  As Integer = 2

'//Registry access masks
Private Const KEY_QUERY_VALUE                     As Long = &H1
Private Const KEY_SET_VALUE                       As Long = &H2
Private Const KEY_CREATE_SUB_KEY                  As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS              As Long = &H8
Private Const KEY_NOTIFY                          As Long = &H10
Private Const KEY_CREATE_LINK                     As Long = &H20
Private Const KEY_WOW64_32KEY                     As Long = &H200
Private Const KEY_WOW64_64KEY                     As Long = &H100

Private Const KEY_READ                            As Double = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE                           As Double = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE                         As Double = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS                      As Double = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Enum REGISTRY_ACCESS
    '//unique rights
    Registry_Read = KEY_READ
    Registry_Write = KEY_WRITE
    Registry_Execute = KEY_EXECUTE
    Registry_Full_Control = KEY_ALL_ACCESS
    '//compound rights
    Registry_Read_Write = KEY_READ Or KEY_WRITE
    Registry_Read_Execute = KEY_READ Or KEY_EXECUTE
    Registry_Read_Write_Execute = KEY_READ Or KEY_WRITE Or KEY_EXECUTE
End Enum

Public Enum ACCESS_TYPE
    '//permissive DACL
    Access_Allowed
    '//deny DACL
    Access_Denied
End Enum

'<<<inheritence properties>>>
Public Enum INHERITENCE_FLAGS
    Container_Inherit = CONTAINER_INHERIT_ACE
    Object_Inherit = OBJECT_INHERIT_ACE
    Non_Propogate = NO_PROPAGATE_INHERIT_ACE
    Inherit_Only = INHERIT_ONLY_ACE
    Inherit_Ace = INHERITED_ACE
    Object_Container_Inherit = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
    Child_Inherit_Level = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE Or NO_PROPAGATE_INHERIT_ACE
    Child_Container_Inherit = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE Or INHERIT_ONLY_ACE
    Family_Container_Inherit = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE Or INHERITED_ACE
End Enum

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
    eLockMemory = 10
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

Private cReg As clsReg


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                           REGISTRY INTERFACE
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function Access_Check(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String) As Boolean

'//test user write access rights
'//to a specified key
On Error GoTo Handler
    '//test write access to a key
    With New clsReg
        If .Access_Test(RootKey, SubKey) Then
            Access_Check = True
        End If
    End With
    
Handler:
On Error GoTo 0

End Function

Public Function Read_BEndian(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String) As Variant

'//list a big_endian value
Dim lReturn     As Long

On Error GoTo Handler

    '//read value
    With New clsReg
        Read_BEndian = .cRead_BEndian(RootKey, SubKey, KeyVal, lReturn)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

Public Function Write_BEndian(ByVal RootKey As HKEY_Type, _
                              ByVal SubKey As String, _
                              ByVal KeyVal As String, _
                              ByVal vData As Variant) As Boolean

'//write a big_endian value
Dim lReturn     As Long

On Error GoTo Handler
    '//write value
    With New clsReg
        lReturn = .cWrite_BEndian(RootKey, SubKey, KeyVal, vData)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_BEndian = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

Public Function Read_Binary(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal KeyVal As String) As Variant

'//get binary data
Dim lReturn     As Long
'On Error GoTo Handler
    
    '//read value
    With New clsReg
        Read_Binary = .cRead_Binary(RootKey, SubKey, KeyVal, lReturn)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

Public Function Write_Binary(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String, _
                             ByVal sData As String) As Boolean

'//write a binary value
Dim lReturn     As Long
Dim j           As Long
Dim MByte()     As Byte
Dim Btemp()     As String

On Error GoTo Handler

    Btemp() = Split(sData, " ")
    For j = 0 To UBound(Btemp) - 1
        ReDim Preserve MByte(j)
        MByte(j) = CByte(Btemp(j))
    Next j
    
    With New clsReg
        lReturn = .cWrite_Binary(RootKey, SubKey, KeyVal, MByte)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_Binary = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

Public Function Read_Dword(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String, _
                           ByVal KeyVal As String) As String

'//get dword data
'//will also read little endian
Dim lReturn     As Long

On Error GoTo Handler
    '//read value
    With New clsReg
        Read_Dword = .cRead_DWord(RootKey, SubKey, KeyVal, lReturn)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

Public Function Write_Dword(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal KeyVal As String, _
                            ByVal lData As Long) As Boolean

'//write a dword value
Dim lReturn     As Long

On Error GoTo Handler
    '//write value
    With New clsReg
        lReturn = .cWrite_DWord(RootKey, SubKey, KeyVal, lData)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_Dword = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

Public Function Read_LEndian(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String) As Variant

'//list a little_endian value
Dim lReturn     As Long

On Error GoTo Handler

    '//read value
    With New clsReg
        Read_LEndian = .cRead_LEndian(RootKey, SubKey, KeyVal, lReturn)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

Public Function Write_LEndian(ByVal RootKey As HKEY_Type, _
                              ByVal SubKey As String, _
                              ByVal KeyVal As String, _
                              ByVal vData As Variant) As Boolean

'//write a little_endian value
Dim lReturn     As Long

On Error GoTo Handler
    '//write value
    With New clsReg
        lReturn = .cWrite_LEndian(RootKey, SubKey, KeyVal, vData)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_LEndian = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

'//WIN32 ONLY
Public Function Read_Link(ByVal RootKey As HKEY_Type, _
                          ByVal SubKey As String, _
                          ByVal KeyVal As String) As Variant

'//get link data
Dim lReturn     As Long
'On Error GoTo Handler
    
    '//read value
    With New clsReg
        Read_Link = .cRead_Link(RootKey, SubKey, KeyVal, 0)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

'//WIN32 ONLY
Public Function Write_Link(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String, _
                             ByVal sData As String) As Boolean

'//write a link value
Dim lReturn     As Long
Dim j           As Long
Dim MByte()     As Byte
Dim Btemp()     As String

On Error GoTo Handler

    Btemp() = Split(sData, " ")
    For j = 0 To UBound(Btemp) - 1
        ReDim Preserve MByte(j)
        MByte(j) = CByte(Btemp(j))
    Next j
    
    With New clsReg
        lReturn = .cWrite_Link(RootKey, SubKey, KeyVal, MByte)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_Link = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

Public Function Read_Multi(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String, _
                           ByVal KeyVal As String) As String

'//read multi value into string
Dim sTemp       As String
Dim l           As Long
Dim lReturn     As Long
Dim sMVal       As String
Dim aMVals()    As String

On Error GoTo Handler
    '//read value
    With New clsReg
        sMVal = .cRead_Multi(RootKey, SubKey, KeyVal)
    End With
    
    '//format results for string
    aMVals = Split(sMVal, vbNullChar)
    For l = 0 To UBound(aMVals)
        sTemp = sTemp & aMVals(l) & " "
    Next l
    '//trim terminating null char
    Read_Multi = Left$(sTemp, Len(sTemp) - 1)
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

Public Function Write_Multi(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal KeyVal As String, _
                            ByVal sData As String) As Boolean

'//write a multi_sz value
Dim lReturn     As Long

On Error GoTo Handler
    '//write value
    With New clsReg
        lReturn = .cWrite_Multi(RootKey, SubKey, KeyVal, sData)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_Multi = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

Public Function Read_MultiCN(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String) As Collection

'//process multi_sz value
'//and return collection
Dim aTemp()     As String
Dim cTemp       As New Collection
Dim sTemp       As String
Dim l           As Long
Dim lReturn     As Long

On Error Resume Next

    Set cTemp = New Collection
    '//get multi string
    With New clsReg
        sTemp = .cRead_Multi(RootKey, SubKey, KeyVal, lReturn)
    End With

    '//get values and add to collection
    aTemp = Split(sTemp, vbNullChar)
    For l = 0 To UBound(aTemp)
        cTemp.Add aTemp(l)
    Next l
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    '//set collection and release library
    Set Read_MultiCN = cTemp
    Set cTemp = Nothing
    
On Error GoTo 0

End Function

Public Sub Write_MultiCN(ByVal RootKey As HKEY_Type, _
                         ByVal SubKey As String, _
                         ByVal KeyVal As String, _
                         ByVal DataSet As Collection)

'//process a collection and
'//write to registry as a
'//multi_sz entry
Dim sTemp       As String
Dim sValue      As String
Dim cTemp       As New Collection
Dim l           As Long
Dim lReturn     As Long
Dim Item        As Variant

On Error Resume Next

    For Each Item In DataSet
        sTemp = sTemp & Item & vbNullChar
    Next
    '//add terminating null char
    sTemp = sTemp & vbNullChar
    '//write to registry
    With New clsReg
        lReturn = .cWrite_Multi(RootKey, SubKey, KeyVal, sTemp)
    End With
    
    '//test for error and release library
    If Not lReturn = 0 Then Get_Error lReturn
    
On Error GoTo 0

End Sub

'//WIN32 ONLY
Public Function Read_Qword(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String, _
                           ByVal KeyVal As String) As Currency

'//get qword data
'//qwords are 64 bit integers
'//used currency as it is 64
'//but to display, you must
'//multiply by 10000
'//to move decimal point

Dim lReturn     As Long

On Error GoTo Handler
    '//read value
    With New clsReg
        Read_Qword = .cRead_QWord(RootKey, SubKey, KeyVal, lReturn)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

'//WIN32 ONLY
Public Function Write_Qword(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal KeyVal As String, _
                            ByVal cData As Currency) As Boolean

'//write a qword (64bit dword) value
'//wasn't sure how to use 64 bit in vb
'//currency was option I chose..
Dim lReturn     As Long

On Error GoTo Handler
    '//write value
    With New clsReg
        lReturn = .cWrite_QWord(RootKey, SubKey, KeyVal, cData)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_Qword = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

'//WIN32 ONLY
Public Function Read_ResDesc(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String) As Variant

'//get resource descriptor
Dim lReturn     As Long
'On Error GoTo Handler
    
    '//read value
    With New clsReg
        Read_ResDesc = .cRead_ResDescriptor(RootKey, SubKey, KeyVal, lReturn)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

'//WIN32 ONLY
Public Function Write_ResDesc(ByVal RootKey As HKEY_Type, _
                              ByVal SubKey As String, _
                              ByVal KeyVal As String, _
                              ByVal sData As String) As Boolean

'//write a resource descriptor value
Dim lReturn     As Long
Dim j           As Long
Dim MByte()     As Byte
Dim Btemp()     As String

On Error GoTo Handler

    Btemp() = Split(sData, " ")
    For j = 0 To UBound(Btemp) - 1
        ReDim Preserve MByte(j)
        MByte(j) = CByte(Btemp(j))
    Next j
    
    With New clsReg
        lReturn = .cWrite_ResDescriptor(RootKey, SubKey, KeyVal, MByte)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_ResDesc = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

'//WIN32 ONLY
Public Function Read_ResList(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String) As Variant

'//get resource list
Dim lReturn     As Long
'On Error GoTo Handler
    
    '//read value
    With New clsReg
        Read_ResList = .cRead_ResourceList(RootKey, SubKey, KeyVal, lReturn)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

'//WIN32 ONLY
Public Function Write_ResList(ByVal RootKey As HKEY_Type, _
                              ByVal SubKey As String, _
                              ByVal KeyVal As String, _
                              ByVal sData As String) As Boolean

'//write a resource list value
Dim lReturn     As Long
Dim j           As Long
Dim MByte()     As Byte
Dim Btemp()     As String

On Error GoTo Handler

    Btemp() = Split(sData, " ")
    For j = 0 To UBound(Btemp) - 1
        ReDim Preserve MByte(j)
        MByte(j) = CByte(Btemp(j))
    Next j
    
    With New clsReg
        lReturn = .cWrite_ResourceList(RootKey, SubKey, KeyVal, MByte)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_ResList = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

'//WIN32 ONLY
Public Function Read_ResRequired(ByVal RootKey As HKEY_Type, _
                                 ByVal SubKey As String, _
                                 ByVal KeyVal As String) As Variant

'//get resource list
Dim lReturn     As Long
'On Error GoTo Handler
    
    '//read value
    With New clsReg
        Read_ResRequired = .cRead_ResRequired(RootKey, SubKey, KeyVal, lReturn)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

'//WIN32 ONLY
Public Function Write_ResRequired(ByVal RootKey As HKEY_Type, _
                                  ByVal SubKey As String, _
                                  ByVal KeyVal As String, _
                                  ByVal sData As String) As Boolean

'//write a resource list value
Dim lReturn     As Long
Dim j           As Long
Dim MByte()     As Byte
Dim Btemp()     As String

On Error GoTo Handler

    Btemp() = Split(sData, " ")
    For j = 0 To UBound(Btemp) - 1
        ReDim Preserve MByte(j)
        MByte(j) = CByte(Btemp(j))
    Next j
    
    With New clsReg
        lReturn = .cWrite_ResRequired(RootKey, SubKey, KeyVal, MByte)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_ResRequired = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

Public Function Read_String(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal KeyVal As String) As String

'//list any string value
'//sz, expand_sz, link, list
Dim lReturn     As Long

On Error Resume Next

    '//read value
    With New clsReg
        Read_String = .cRead_String(RootKey, SubKey, KeyVal, lReturn)
    End With
    '//test for error
    If Not lReturn = 0 Then Get_Error lReturn
    
Handler:
On Error GoTo 0

End Function

Public Function Write_String(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String, _
                             ByVal sData As String) As Boolean

'//write string data
Dim lReturn     As Long

On Error GoTo Handler
    '//write value
    With New clsReg
        lReturn = .cWrite_String(RootKey, SubKey, KeyVal, sData)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_String = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

Public Function Write_StringEX(ByVal RootKey As HKEY_Type, _
                               ByVal SubKey As String, _
                               ByVal KeyVal As String, _
                               ByVal sData As String) As Boolean

'//write an expanded string value
Dim lReturn     As Long

On Error GoTo Handler
    '//write value
    With New clsReg
        lReturn = .cWrite_String_Expand(RootKey, SubKey, KeyVal, sData)
    End With
    
    '//test for error
    If lReturn = 0 Then
        Write_StringEX = True
    Else
        Get_Error lReturn
    End If
    
Handler:
On Error GoTo 0

End Function

Public Function Create_Key(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String) As Boolean

'//create a key
Dim lReturn     As Long

On Error GoTo Handler

    With New clsReg
        lReturn = .cCreate_Key(RootKey, SubKey)
        If lReturn = 0 Then
            Create_Key = True
        Else
            Get_Error lReturn
        End If
    End With
    
Handler:
On Error GoTo 0

End Function

Public Function Delete_Key(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String) As Boolean

'//delete a key
Dim lReturn     As Long

On Error GoTo Handler

    '//delete the specified key
    With New clsReg
        lReturn = .cDelete_Key(RootKey, SubKey)
        If lReturn = 0 Then
            Delete_Key = True
        Else
            Get_Error lReturn
        End If
    End With
    
Handler:
On Error GoTo 0

End Function

Public Function Key_Exists(ByVal RootKey As HKEY_Type, _
                           ByVal SubKey As String) As Boolean

'//test for key
On Error GoTo Handler

    '//if no error then key exists
    With New clsReg
        If .cKey_Exist(RootKey, SubKey) Then
            Key_Exists = True
        End If
    End With

Handler:
On Error GoTo 0

End Function

Public Function Value_Exists(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String) As Boolean

'//test for key
On Error GoTo Handler

    '//if no error then key exists
    With New clsReg
        If .cValue_Exist(RootKey, SubKey, KeyVal) Then
            Value_Exists = True
        End If
    End With

Handler:
On Error GoTo 0

End Function

Public Function Write_Value(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal KeyVal As String, _
                            ByVal ValData As String, _
                            ByVal DataType As Integer) As Boolean

'//create a new value with data
'//supports sz, multi_sz, expand_sz
'//binary, dword, little_endian
'//big_endian and link
Dim j           As Integer
Dim lData       As Long
Dim MByte()     As Byte
Dim Btemp()     As String
Dim lReturn     As Long

On Error GoTo Handler

    With New clsReg
        Select Case DataType
            Case 1
                '//sz
                lReturn = .cWrite_String(RootKey, SubKey, KeyVal, ValData)
        
            Case 2
                '//expand_sz
                lReturn = .cWrite_String_Expand(RootKey, SubKey, KeyVal, ValData)
            
            Case 3
                '//multi_sz
                lReturn = .cWrite_Multi(RootKey, SubKey, KeyVal, ValData)
            
            Case 4
                '//binary
                Btemp() = Split(ValData, " ")
                For j = 0 To UBound(Btemp) - 1
                    ReDim Preserve MByte(j)
                    MByte(j) = CByte(Btemp(j))
                Next j
                lReturn = .cWrite_Binary(RootKey, SubKey, KeyVal, MByte)
        
            Case 5
                '//dword
                lData = CLng(ValData)
                lReturn = .cWrite_DWord(RootKey, SubKey, KeyVal, lData)
    
            Case 6
                '//little_endian
                lReturn = .cWrite_LEndian(RootKey, SubKey, KeyVal, lData)
            
            Case 7
                '//big_endian
                lReturn = .cWrite_BEndian(RootKey, SubKey, KeyVal, lData)
        End Select
    End With
    
    If lReturn = 0 Then
        Write_Value = True
    Else
        Get_Error lReturn
    End If
    
    Erase Btemp
    Erase MByte

Handler:
On Error GoTo 0

End Function

Public Function Delete_Value(ByVal RootKey As HKEY_Type, _
                             ByVal SubKey As String, _
                             ByVal KeyVal As String) As Boolean

'//delete a value
Dim lReturn     As Long

On Error GoTo Handler

    '//deletes the specified value
    With New clsReg
        lReturn = .cDelete_Value(RootKey, SubKey, KeyVal)
        If lReturn = 0 Then
            Delete_Value = True
        Else
            Get_Error lReturn
        End If
    End With

    
Handler:
On Error GoTo 0

End Function

Public Function Search_Value(ByVal RootKey As HKEY_Type, _
                             ByVal KeyVal As String, _
                             ByVal SearchVal As String) As Boolean

'//search for a value
Dim Item        As Variant
Dim lReturn     As Long

On Error GoTo Handler
    
    '//enumerate values and compare
    '//to search item
    With New clsReg
        For Each Item In .cSearch_Values(RootKey, KeyVal, lReturn)
            If LCase$(Item) = SearchVal Then
                Search_Value = True
                Exit For
            End If
        Next Item
    End With
    
    '//log errors
    If Not lReturn = 0 Then Get_Error lReturn
        
Handler:
On Error GoTo 0
    
End Function

Public Function List_Keys(ByVal RootKey As HKEY_Type, _
                          ByVal SubKey As String) As Collection

'//list all subkeys in path
Dim Item        As cRegKey
Dim cTemp       As New Collection
Dim lReturn     As Long

On Error GoTo Handler

    Set cTemp = New Collection
    '//add each key returned from library
    '//to a new collection
    With New clsReg
        For Each Item In .cList_Keys(RootKey, SubKey, lReturn)
            cTemp.Add Item.Key
        Next Item
    End With
    
    '//test for error, and release library
    If Not lReturn = 0 Then Get_Error lReturn
    Set List_Keys = cTemp
    Set cTemp = Nothing
    
Handler:
On Error GoTo 0

End Function

Public Function List_Values(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String) As Collection

'//add values under key to collection
'//lists values of type string, string_exp
'//dword, little_endian, and binary
Dim Item        As cRegValue
Dim sCurrVal    As String
Dim cTemp       As New Collection
Dim lReturn     As Long

On Error GoTo Handler

    Set cTemp = New Collection
    '//add each value to collection
    With New clsReg
        For Each Item In .cList_Values(RootKey, SubKey, lReturn)
            cTemp.Add Item.Key
        Next Item
    End With
    '//test for error and release
    '//collection and library
    If Not lReturn = 0 Then Get_Error lReturn
    Set List_Values = cTemp
    Set cTemp = Nothing
    
Handler:
On Error GoTo 0

End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                         ERROR LOGGING
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Function Get_Error(ByVal lErrNum As Long) As String
'//interpret registry errors passed from library
'//including custom errors from clsRegSecurity

On Error GoTo Handler

    Select Case lErrNum
    Case 1
        Get_Error = "Error: " & lErrNum & " Bad Database. The database is invalid"
    Case 2
        Get_Error = "Error: " & lErrNum & " Bad Key Name. Key is corrupt or does not exist"
    Case 3
        Get_Error = "Error: " & lErrNum & " Key Locked. The key requested does not exist"
    Case 4
        Get_Error = "Error: " & lErrNum & " Failed Read. The key requested can not be read accessed"
    Case 5
        Get_Error = "Error: " & lErrNum & " Failed Write. The key requested can not be write accessed"
    Case 6
        Get_Error = "Error: " & lErrNum & " Out Of Memory. The process has exceeded its memory allocation"
    Case 7
        Get_Error = "Error: " & lErrNum & " Work Area Invalid. The process work area appears to be corrupt"
    Case 8
        Get_Error = "Error: " & lErrNum & " Access Is Denied. Access to this key has been denied"
    Case 87
        Get_Error = "Error: " & lErrNum & " Invalid Parameters. The call parameters passed are invalid"
    Case 234
        Get_Error = "Error: " & lErrNum & " More Data. The container has more data"
    Case 259
        Get_Error = "Error: " & lErrNum & " No More Items. No more items for this query"
    '//begin clsRegSecurity custom error messages
    Case 901
        Get_Error = "Error: " & lErrNum & " Could not intialize SWSA SID"
    Case 902
        Get_Error = "Error: " & lErrNum & " Could not open key"
    Case 903
        Get_Error = "Error: " & lErrNum & " Failed to create Key Descriptor"
    Case 904
        Get_Error = "Error: " & lErrNum & " Could not retieve SID with LocalAlloc"
    Case 905
        Get_Error = "Error: " & lErrNum & " Get Key Security Failure"
    Case 906
        Get_Error = "Error: " & lErrNum & " Buffer Creation Failed"
    Case 907
        Get_Error = "Error: " & lErrNum & " Could not Initialize Security Descriptor"
    Case 908
        Get_Error = "Error: " & lErrNum & " Empty DACL or Request Failure"
    Case 909
        Get_Error = "Error: " & lErrNum & " No DACL return data"
    Case 910
        Get_Error = "Error: " & lErrNum & " Account Name is Invalid"
    Case 911
        Get_Error = "Error: " & lErrNum & " Requested SID does not exist"
    Case 912
        Get_Error = "Error: " & lErrNum & " Memory allocation failure for new DACL"
    Case 913
        Get_Error = "Error: " & lErrNum & " New DACL could not be initialized"
    Case 914
        Get_Error = "Error: " & lErrNum & " LocalAlloc Ace returned zero"
    Case 915
        Get_Error = "Error: " & lErrNum & " DACL type NON_INHERITENCE was not added to DACL"
    Case 916
        Get_Error = "Error: " & lErrNum & " Failed to add ACE to ACL Structure"
    Case 917
        Get_Error = "Error: " & lErrNum & " DACL type ACCESS_ALLOWED_ACE_TYPE was not added to DACL"
    Case 918
        Get_Error = "Error: " & lErrNum & " ACE Entry could Not be Allocated"
    Case 919
        Get_Error = "Error: " & lErrNum & " ACL Build Failure"
    Case 920
        Get_Error = "Error: " & lErrNum & " Failed to set Descriptor control entry"
    Case 921
        Get_Error = "Error: " & lErrNum & " Failed to Allocate Key Descriptor"
    Case Else
        Get_Error = "Error: An Unknown Error has occured"
    End Select
    
    Log_Error Get_Error

Handler:
On Error GoTo 0

End Function

Private Sub Log_Error(ByVal ErrDesc As String)
'//log errors

On Error Resume Next

    '//write to log
    Open App.Path & "\err.log" For Append As #1
        Print #1, CStr(Now) & vbTab & ErrDesc
    Close #1

On Error GoTo 0

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                       ENDIAN CONVERSION ROUTINES
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'// Thanks go out to actorics.de PSC submission

Private Function Low_Byte(ByVal iNum As Integer) As Byte

On Error GoTo Handler

    Low_Byte = iNum And &HFF

Handler:
On Error GoTo 0

End Function

Private Function High_Byte(ByVal iNum As Integer) As Byte

On Error GoTo Handler

    High_Byte = (iNum And &HFF00&) \ 256

Handler:
On Error GoTo 0

End Function

Private Function Make_Word(ByVal bLow As Byte, ByVal bHigh As Byte) As Integer

On Error GoTo Handler

    If bHigh And &H80 Then
        Make_Word = ((bHigh * 256&) + bLow) Or &HFFFF0000
    Else
        Make_Word = (bHigh * 256) + bLow
    End If

Handler:
On Error GoTo 0

End Function

Public Function Make_LEndian16(ByVal iVal As Integer) As Integer
'//create a 16bit little_endian
On Error GoTo Handler

Dim b(1) As Byte

    b(0) = Low_Byte(iVal)
    b(1) = High_Byte(iVal)
    Make_LEndian16 = Make_Word(b(1), b(0))

Handler:
On Error GoTo 0

End Function

Public Function Make_LEndian32(lVal As Long) As String
'//create 32bit little_endian
Dim b0 As Byte
Dim b1 As Byte
Dim b2 As Byte
Dim b3 As Byte

On Error GoTo Handler

    b0 = (lVal And &HFF&)
    b1 = (lVal And &HFF00&) \ &H100&
    b2 = (lVal And &HFF0000) \ &H10000
    b3 = (lVal And &H7F000000) \ &H1000000 - 128 * (lVal < 0)
    Make_LEndian32 = Chr(b0) & Chr(b1) & Chr(b2) & Chr(b3)

Handler:
On Error GoTo 0

End Function

Public Function Make_BEndian32(lVal As Long) As String
'//create 32bit big_endian
Dim b0 As Byte
Dim b1 As Byte
Dim b2 As Byte
Dim b3 As Byte

On Error GoTo Handler

    b0 = (lVal And &HFF&)
    b1 = (lVal And &HFF00&) \ &H100&
    b2 = (lVal And &HFF0000) \ &H10000
    b3 = (lVal And &H7F000000) \ &H1000000 - 128 * (lVal < 0)
    Make_BEndian32 = Chr(b3) & Chr(b2) & Chr(b1) & Chr(b0)

Handler:
On Error GoTo 0

End Function

Public Function Convert_Unicode(ByVal sWord As String) As Variant
'//convert ascii to unicode
On Error GoTo Handler

    Convert_Unicode = StrConv(sWord, vbUnicode)
    
Handler:
On Error GoTo 0

End Function

Public Function Convert_Byte(ByVal sVal As String) As Byte
'//convert string to byte array
Dim bVal()  As Byte

Dim i As Long

On Error GoTo Handler

    ReDim bVal(0 To Len(sVal))
    For i = 0 To UBound(bVal())
        bVal(i) = Mid$(sVal, i, 1)
    Next i
    
Handler:
On Error GoTo 0

End Function

Public Function Convert_Curr(ByVal sValue As String) As Currency

Dim l           As Long
Dim Negative    As Boolean

    sValue = Trim$(sValue)
    If Left$(sValue, 1) = "-" Then
        Negative = True
        sValue = Mid$(sValue, 2)
    End If
    
    l = Len(sValue)
    If l < 4 Then
        Convert_Curr = CCur(IIf(Negative, "-0.", "0.") & Right$("0000" & sValue, 4))
    Else
        Convert_Curr = CCur(IIf(Negative, "-", "") & Left$(sValue, l - 4) & "." & Right$(sValue, 4))
    End If
    
End Function

Public Function Convert_Text(ByVal cValue As Currency) As String

Dim Temp As String, l As Long

    Temp = Format$(cValue, "#.0000")
    l = Len(Temp)
    Temp = Left$(Temp, l - 5) & Right$(Temp, 4)
    
    Do While Len(Temp) > 1 And Left$(Temp, 1) = "0"
        Temp = Mid$(Temp, 2)
    Loop
    
    Do While Len(Temp) > 2 And Left$(Temp, 2) = "-0"
        Temp = "-" & Mid$(Temp, 3)
    Loop
    
    Temp = Temp / 10000
    Convert_Text = Temp
        
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                            BACKUP AND RESTORE FUNCTIONS (clsRegBackup)
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function Test_Access(eAccess As eSEUserRights) As Boolean

On Error GoTo Handler

    '//test if the user has a specified
    '//level of access
    '//used to test user rights
    '//before commiting to an action
    With New clsRegBackup
        If .cAccess_Level(eAccess) Then
            Test_Access = True
        End If
    End With

Handler:
On Error GoTo 0

End Function

Public Function Save_Key(ByVal RootKey As HKEY_Type, _
                         ByVal SubKey As String, _
                         ByVal sFile As String) As Boolean

On Error GoTo Handler

    '//saves the specified key
    '//to a binary image file
    With New clsRegBackup
        If .cSave_Key(RootKey, SubKey, sFile) Then
            Save_Key = True
        End If
    End With
    
Handler:
On Error GoTo 0

End Function
 
Public Function Restore_Key(ByVal RootKey As HKEY_Type, _
                            ByVal SubKey As String, _
                            ByVal sFile As String) As Boolean

On Error GoTo Handler

    '//restores key from a
    '//binary image
    With New clsRegBackup
        If .cRestore_Key(RootKey, SubKey, sFile) Then
            Restore_Key = True
        End If
    End With
    
Handler:
On Error GoTo 0

End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                               REGISTRY ACCESS PERMISSIONS (clsRegSecurity)
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function Set_Key_Permissions(ByVal RootKey As HKEY_Type, _
                                    ByVal sKey As String, _
                                    ByVal sAccount As String, _
                                    ByVal eMask As REGISTRY_ACCESS, _
                                    ByVal eType As ACCESS_TYPE, _
                                    ByVal eFlags As INHERITENCE_FLAGS, _
                                    Optional ByVal lErrNum As Long) As Boolean

On Error GoTo Handler

    With New clsRegSecurity
        If .Update_Key_Permissions(RootKey, sKey, sAccount, eMask, eType, eFlags, lErrNum) Then
            Set_Key_Permissions = True
        End If
    End With

Handler:
On Error GoTo 0

End Function
