VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegSecurity"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'//this stuff was hard to find, and even harder to get working.
'//hope you like it..

'##List of Custom Errors##
'//Could not intialize SWSA SID. Err# 901
'//Could not open key. Err# 902
'//Failed to create Key Descriptor.  Err# 903
'//Could not retieve SID with LocalAlloc. Err# 904
'//Get Key Security Failure. Err# 905
'//Buffer creation failed: Err#  906
'//Could not Initialize Descriptor: " & lPsd & ". Err# 907
'//Empty DACL or request failure. Err# 908
'//No DACL return data. Err# 909
'//Account.AccountName & " Err# 910
'//Requested SID does not exist: " & Account.AccountName & " Err# 911
'//Memory allocation failure for new DACL. Err# 912
'//New DACL could not be initialized. Err# 913
'//LocalAlloc lAce returned zero. Err# 914
'//DACL type NON_INHERITENCE was not added to DACL.  Err# 915
'//Failed to add ACE to structure. Err# 916
'//DACL type ACCESS_ALLOWED_ACE_TYPE was not added to DACL: " & Account.pSid & " Err# 917
'//ACE entry could not be allocated.  Err# 918
'//ACL build failure. Err# 919
'//Failed to set Descriptor control entry. Err# 920
'//Failed to add DACL to Descriptor. Err# 921
'//Failed to Allocate Key Descriptor. Err# 922

Public Type udtAccountPerm
    AccountName                                       As String
    AccessMask                                        As Long
    AceFlags                                          As Byte
    AceType                                           As Byte
    pSid                                              As Long
    SidPassedByCaller                                 As Boolean
End Type

Private Type udtSDMInfo
    pSD                                               As Long
    pAcl                                              As Long
End Type

'//ACL manipulation
Private Type udtACL
    AclRevision                                       As Byte
    Sbz1                                              As Byte
    AclSize                                           As Integer
    AceCount                                          As Integer
    Sbz2                                              As Integer
End Type

Private Type udtAclSize
    AceCount                                          As Long
    AclBytesInUse                                     As Long
    AclBytesFree                                      As Long
End Type

Private Type udtAceHeader
    AceType                                           As Byte
    AceFlags                                          As Byte
    AceSize                                           As Integer
End Type

Private Type udtACE
    Header                                            As udtAceHeader
    Mask                                              As Long
    SidStart                                          As Long
End Type

Private Type udtSecAttrib
    Length                                            As Long
    SecurityDescriptor                                As Long
    InheritHandle                                     As Long
End Type

Private Type udtSidID
    Value(6)                                          As Byte
End Type

'//Memory constants
Private Const LMEM_FIXED                          As Long = &H0
Private Const LMEM_ZEROINIT                       As Long = &H40
Private Const lPtr                                As Double = (LMEM_FIXED + LMEM_ZEROINIT)

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
Private Const KEY_QUERY_VALUE                      As Long = &H1
Private Const KEY_SET_VALUE                       As Long = &H2
Private Const KEY_CREATE_SUB_KEY                  As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS               As Long = &H8
Private Const KEY_NOTIFY                          As Long = &H10
Private Const KEY_CREATE_LINK                     As Long = &H20
'Private Const KEY_WOW64_32KEY                     As Long = &H200
'Private Const KEY_WOW64_64KEY                     As Long = &H100

Private Const KEY_READ                            As Double = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE                            As Double = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE                         As Double = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS                       As Double = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

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

'<<<DACL type>>>
Public Enum ACCESS_TYPE
    '//permissive DACL
    Access_Allowed
    '//deny DACL
    Access_Denied
End Enum

'<<<ACE access mask>>>
Public Enum ACCESS_MASK
    Delete_Only = DELETE Or READ_CONTROL
    Execute_Only = STANDARD_RIGHTS_EXECUTE Or STANDARD_RIGHTS_READ
    Read_Only = STANDARD_RIGHTS_READ
    Read_Write = READ_CONTROL Or STANDARD_RIGHTS_WRITE
    Read_Write_Delete = STANDARD_RIGHTS_READ Or STANDARD_RIGHTS_WRITE Or DELETE
    Read_Write_Execute = STANDARD_RIGHTS_READ Or STANDARD_RIGHTS_WRITE Or STANDARD_RIGHTS_EXECUTE
    Read_Write_Execute_Delete = STANDARD_RIGHTS_READ Or STANDARD_RIGHTS_WRITE Or STANDARD_RIGHTS_EXECUTE Or DELETE
    Standard_Rights = STANDARD_RIGHTS_ALL
    Full_Control = MAXIMUM_ALLOWED
    System_Control = MAXIMUM_ALLOWED Or ACCESS_SYSTEM_SECURITY
End Enum

'<<<inheritence properties>>>
Public Enum INHERITENCE_FLAGS
    '<<<singular inheritence attributes first>>>
    '//folder and and future subfolders inherit attributes
    Container_Inherit = CONTAINER_INHERIT_ACE
    '//folder and future files inherit attributes
    Object_Inherit = OBJECT_INHERIT_ACE
    '//just the direct children
    Non_Propogate = NO_PROPAGATE_INHERIT_ACE
    '//ACE applies not to this object, but to child objects
    Inherit_Only = INHERIT_ONLY_ACE
    '//apply to parent and child
    Inherit_Ace = INHERITED_ACE
    '<<<compound inheritence structures>>>
    '//folder, and future subfolders and files inherit attributes
    Object_Container_Inherit = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE
    '//subfolders, files, children only, one level
    Child_Inherit_Level = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE Or NO_PROPAGATE_INHERIT_ACE
    '//subfolders, files, children only
    Child_Container_Inherit = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE Or INHERIT_ONLY_ACE
    '//subfolders, files, parent and children
    Family_Container_Inherit = OBJECT_INHERIT_ACE Or CONTAINER_INHERIT_ACE Or INHERITED_ACE
End Enum

'//registry api
Private Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, _
                                                               ByVal SecurityInformation As Long, _
                                                               ByVal pSecurityDescriptor As Long) As Long
                                                               
Private Declare Function LocalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, _
                                                        ByVal wBytes As Long) As Long
                                                        
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, _
                                                                         ByVal hpvSource As Long, _
                                                                         ByVal cbCopy As Long)
                                                                         
Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                          ByVal dwRevision As Long) As Long
                                                                          
Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As Long, _
                                                                                          ByVal lpAccountName As String, _
                                                                                          ByVal Sid As Long, _
                                                                                          cbSid As Long, _
                                                                                          ByVal ReferencedDomainName As String, _
                                                                                          cbReferencedDomainName As Long, _
                                                                                          peUse As Long) As Long
                                                                                          
Private Declare Function GetLengthSid Lib "advapi32.dll" (ByVal pSid As Long) As Long

Private Declare Function InitializeAcl Lib "advapi32.dll" (ByVal pAcl As Long, _
                                                           ByVal nAclLength As Long, _
                                                           ByVal dwAclRevision As Long) As Long
                                                           
Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                       ByVal bDaclPresent As Long, _
                                                                       ByVal pDacl As Long, _
                                                                       ByVal bDaclDefaulted As Long) As Long
                                                                       
Private Declare Function GetAce Lib "advapi32.dll" (ByVal pAcl As Long, _
                                                    ByVal dwAceIndex As Long, _
                                                    pACE As Long) As Long
                                                    
Private Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                       lpbDaclPresent As Long, _
                                                                       pDacl As Long, _
                                                                       lpbDaclDefaulted As Long) As Long
                                                                       
Private Declare Function GetAclInformation Lib "advapi32.dll" (ByVal pAcl As Long, _
                                                               pAclInformation As Any, _
                                                               ByVal nAclInformationLength As Long, _
                                                               ByVal dwAclInformationClass As Long) As Long
                                                               
Private Declare Function GetSecurityDescriptorControl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                          pControl As Long, _
                                                                          lpdwRevision As Long) As Long
                                                                          
Private Declare Function SetSecurityDescriptorControl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, _
                                                                          ByVal controlBitsOfInterest As Long, _
                                                                          ByVal controlBitsToSet As Long) As Long
                                                                          
Private Declare Function EqualSid Lib "advapi32.dll" (ByVal pSid1 As Long, _
                                                      ByVal pSid2 As Long) As Long
                                                      
Private Declare Function AddAce Lib "advapi32.dll" (ByVal pAcl As Long, _
                                                    ByVal dwAceRevision As Long, _
                                                    ByVal dwStartingAceIndex As Long, _
                                                    ByVal pAceList As Long, _
                                                    ByVal nAceListLength As Long) As Long
                                                    
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As udtSidID, _
                                                                      ByVal nSubAuthorityCount As Byte, _
                                                                      ByVal nSubAuthority0 As Long, _
                                                                      ByVal nSubAuthority1 As Long, _
                                                                      ByVal nSubAuthority2 As Long, _
                                                                      ByVal nSubAuthority3 As Long, _
                                                                      ByVal nSubAuthority4 As Long, _
                                                                      ByVal nSubAuthority5 As Long, _
                                                                      ByVal nSubAuthority6 As Long, _
                                                                      ByVal nSubAuthority7 As Long, _
                                                                      lpPSid As Long) As Long
                                                                      
Private Declare Sub FreeSid Lib "advapi32.dll" (ByVal pSid As Long)

Private Declare Function GetKernelObjectSecurity Lib "advapi32.dll" (ByVal hObject As Long, _
                                                                     ByVal RequestedInformation As Long, _
                                                                     ByVal pSecurityDescriptor As Long, _
                                                                     ByVal nLength As Long, _
                                                                     lpnLengthNeeded As Long) As Long
                                                                     
Private Declare Function SetKernelObjectSecurity Lib "advapi32.dll" (ByVal hObject As Long, _
                                                                     ByVal SecurityInformation As Long, _
                                                                     ByVal pSecurityDescriptor As Long) As Long
                                                                     
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, _
                                                                            ByVal dwDesiredAccess As Long, _
                                                                            ByVal dwShareMode As Long, _
                                                                            ByVal lpSecurityAttributes As Long, _
                                                                            ByVal dwCreationDisposition As Long, _
                                                                            ByVal dwFlagsAndAttributes As Long, _
                                                                            ByVal hTemplateFile As Long) As Long
                                                                            
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
                                                                                
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, _
                                                               ByVal SecurityInformation As Long, _
                                                               ByVal pSecurityDescriptor As Long, _
                                                               lpcbSecurityDescriptor As Long) As Long

Private Function Set_Key_Permissions(ByVal hKey As Long, _
                                     Account As udtAccountPerm, _
                                     ByVal lErrNum As Long) As Boolean

'//apply permissions to the key
Dim lReturn     As Long
Dim uInfo       As udtSDMInfo
Dim oldSD       As Long
Dim lLength     As Long
Dim bStatus     As Boolean
Dim lErrCond    As Long

On Error GoTo Handler

    bStatus = False
    uInfo.pAcl = 0
    uInfo.pSD = 0
    lLength = 0
    
    '//get descriptor size
    lReturn = RegGetKeySecurity(hKey, DACL_SECURITY_INFORMATION, 0, lLength)

    If lLength = 0 Then
        '//Failed to create Key Descriptor.  Err# 903
        lErrNum = 903
        GoTo Handler
    End If
    '//get SID
    oldSD = LocalAlloc(lPtr, lLength)
    If oldSD = 0 Then
        '//Could not retieve SID with LocalAlloc. Err# 904
        lErrNum = 904
        GoTo Handler
    End If
    '//get key security
    If Not RegGetKeySecurity(hKey, DACL_SECURITY_INFORMATION, oldSD, lLength) = 0 Then
        '//Get Key Security Failure. Err# 905
        lErrNum = 905
        GoTo Handler
    End If
    '//set descriptor
    If Set_Descriptor(oldSD, Account, uInfo, lErrCond) = 0 Then
        '//error in set_descriptor routine
        '//get error val and exit
        lErrNum = lErrCond
        GoTo Handler
    End If
    '//allocate descriptor
    If Not RegSetKeySecurity(hKey, DACL_SECURITY_INFORMATION, uInfo.pSD) = 0 Then
        '//Failed to Allocate Key Descriptor. Err# 922
        lErrNum = 922
        GoTo Handler
    End If

    Set_Key_Permissions = True

Handler:
    '//cleanup
    If oldSD <> 0 Then
        LocalFree oldSD
    End If
    oldSD = 0

    If uInfo.pSD <> 0 Then
        LocalFree uInfo.pSD
    End If
    uInfo.pSD = 0

    If uInfo.pAcl <> 0 Then
        LocalFree uInfo.pAcl
    End If
    uInfo.pAcl = 0

End Function

Public Function Update_Key_Permissions(ByVal lKey As HKEY_Type, _
                                       ByVal sKey As String, _
                                       ByVal sAccount As String, _
                                       ByVal eMask As REGISTRY_ACCESS, _
                                       ByVal eType As ACCESS_TYPE, _
                                       ByVal eFlags As INHERITENCE_FLAGS, _
                                       Optional ByVal lErrNum As Long) As Boolean

'//modify a keys security permissions
Dim bSuccess            As Boolean
Dim hKey                As Long
Dim Account             As udtAccountPerm
Dim lResult             As Long
Dim siaNtAuthority      As udtSidID
Dim lErrCond            As Long

On Error GoTo Handler

    '//set up default account
    With Account
        .AccountName = vbNullString
        .AccessMask = GENERIC_READ
        .AceFlags = CONTAINER_INHERIT_ACE
        .AceType = ACCESS_ALLOWED_ACE_TYPE
    End With

    '//build universal sid
    siaNtAuthority.Value(5) = SECURITY_WORLD_SID_AUTHORITY
    If AllocateAndInitializeSid(siaNtAuthority, 1, SECURITY_WORLD_RID, 0, 0, 0, 0, 0, 0, 0, Account.pSid) = 0 Then
        '//Could not intialize SWSA SID. Err# 901
        lErrNum = 901
        Exit Function
    End If

    '//set up new account
    With Account
        .AccountName = sAccount
        .AccessMask = eMask
        .AceFlags = eFlags
        .AceType = eType
        .pSid = 0
        .SidPassedByCaller = False
    End With

    '//open the key
    If Not RegOpenKeyEx(lKey, sKey, 0, READ_CONTROL Or WRITE_DAC, hKey) = 0 Then
        '//Could not open key. Err# 902
        lErrNum = 902
        Exit Function
    Else
        If Not Set_Key_Permissions(hKey, Account, lErrCond) Then
            lErrNum = lErrCond
            GoTo Handler:
        End If
        RegCloseKey hKey
        Update_Key_Permissions = True
    End If

Handler:
    '//cleanup
    With Account
        If .pSid <> 0 And .SidPassedByCaller Then
            FreeSid (.pSid)
            .pSid = 0
        End If
    End With

End Function

Private Function Set_Descriptor(ByVal lOldSD As Long, _
                                Account As udtAccountPerm, _
                                uInfo As udtSDMInfo, _
                                Optional ByVal lErrNum As Long) As Long

Dim lNewACL         As Long
Dim lNewACLLen      As Long
Dim lTotalDACLLen   As Long
Dim sDomain         As String
Dim lDomain         As Long
Dim lSidLen         As Long
Dim lCount          As Long
Dim lUse            As Long
Dim lReturn         As Long
Dim lResult         As Long
Dim tTempACL        As udtACL
Dim tempAce         As udtACE
Dim lPtr            As Long
Dim lPsd            As Long
Dim lAceIndex       As Long
Dim lDaclPresent    As Long
Dim lDaclDefaulted  As Long
Dim sACLInfo        As udtAclSize
Dim lPAcl           As Long
Dim lFlag           As Long
Dim lRevision       As Long
Dim lControlBits    As Long
Dim lControlSet     As Long
Dim lErrCond        As Long

On Error GoTo Handler

    '//init structure
    uInfo.pAcl = 0
    uInfo.pSD = 0

    '//buffer for new descriptor
    lPsd = LocalAlloc(lPtr, SECURITY_DESCRIPTOR_MIN_LENGTH)
    If lPsd = 0 Then
        '//Buffer creation failed: Err#  906
        lErrNum = 906
        Exit Function
    End If
    uInfo.pSD = lPsd

    '//init descriptor
    lResult = InitializeSecurityDescriptor(lPsd, SECURITY_DESCRIPTOR_REVISION)
    If lResult = 0 Then
        '//Could not Initialize Descriptor: " & lPsd & ". Err# 907
        lErrNum = 907
        GoTo Handler
    End If

    '//check for existing SD
    If lOldSD Then
        lResult = GetSecurityDescriptorDacl(lOldSD, lDaclPresent, lPAcl, lDaclDefaulted)
        If lResult = 0 Then
            '//Empty DACL or request failure. Err# 908
            lErrNum = 908
            GoTo Handler
        End If
        '//extract DACL
        If (lDaclPresent <> 0 And lPAcl <> 0) Then
            lResult = GetAclInformation(lPAcl, sACLInfo, Len(sACLInfo), 2&)
            If lResult = 0 Then
                '//No DACL return data. Err# 909
                lErrNum = 909
                GoTo Handler
            End If
            lTotalDACLLen = sACLInfo.AclBytesInUse
        Else
            lTotalDACLLen = Len(tTempACL)
        End If
    Else
        lTotalDACLLen = Len(tTempACL)
    End If

    '//create a buffer
    sDomain = Space(256)
    '//if SID is not supplied in call, look it up
    If Account.pSid = 0 Then
        lDomain = 256
        '//get size of SID and allocate memory
        lResult = LookupAccountName(0, Account.AccountName, 0, lSidLen, sDomain, lDomain, lUse)
        Account.pSid = LocalAlloc(lPtr, lSidLen)
        If Account.pSid = 0 Then
            '//Account.AccountName & " Err# 910
            lErrNum = 910
        End If

        '//get the SID
        lResult = LookupAccountName(0, Account.AccountName, Account.pSid, lSidLen, sDomain, lDomain, lUse)
        If lResult = 0 Then
            '//Requested SID does not exist: " & Account.AccountName & " Err# 911
            lErrNum = 911
            GoTo Handler
        End If
    End If

    '//get the length for the new ACL
    lNewACLLen = Len(tempAce) + GetLengthSid(Account.pSid) - 4
    lTotalDACLLen = lTotalDACLLen + lNewACLLen

    ' Allocate memory for the new ACL
    lNewACL = LocalAlloc(lPtr, lTotalDACLLen)
    If lNewACL = 0 Then
        '//Memory allocation failure for new DACL. Err# 912
        lErrNum = 912
        GoTo Handler
    End If

    uInfo.pAcl = lNewACL

    '//initialize the new ACL
    lResult = InitializeAcl(lNewACL, lTotalDACLLen, ACL_REVISION)
    If lResult = 0 Then
        '//New DACL could not be initialized. Err# 913
        lErrNum = 913
        GoTo Handler
    End If

    '//build the DACL in sequence, explicit deny is first
    If Account.AceType = ACCESS_DENIED_ACE_TYPE Then
        lResult = Build_New_ACE(lNewACL, Account.AceType, Account.AceFlags, Account.AccessMask, Account.pSid, lErrCond)
        If lResult = 0 Then
            '//Error in Build_New_ACE
            '//get error and exit
            lErrNum = lErrCond
            GoTo Handler
        End If
        lAceIndex = lAceIndex + 1
    End If

    '//copy non-inherited ACEs
    If (lDaclPresent <> 0 And lPAcl <> 0 And sACLInfo.AceCount > 0) Then
        '//combine old and new ACE entries
        For lCount = 0 To (sACLInfo.AceCount - 1)
            '//next ACE
            lResult = GetAce(lPAcl, lCount, lPtr)
            If lResult = 0 Then
                '//DACL type NON_INHERITENCE was not added to DACL.  Err# 915
                lErrNum = 915
                GoTo Handler
            End If

            CopyMemory tempAce, lPtr, LenB(tempAce)
            '//exit once inherited ACE found
            If ((tempAce.Header.AceFlags And INHERITED_ACE) = INHERITED_ACE) Then
                Exit For
            End If

            '//check ACE value
            If Not (Is_Equal(Account, lPtr + 8)) Then
                '//add the ACE to new DACL
                lResult = AddAce(lNewACL, ACL_REVISION, MAXDWORD, lPtr, tempAce.Header.AceSize)
                If lResult = 0 Then
                    '//Failed to add ACE to structure. Err# 916
                    lErrNum = 916
                    GoTo Handler
                End If
                lAceIndex = lAceIndex + 1
            End If
        Next lCount
    End If

    '//add explicit permit entries to DACL
    If Account.AceType = ACCESS_ALLOWED_ACE_TYPE Then
        lResult = Build_New_ACE(lNewACL, Account.AceType, Account.AceFlags, Account.AccessMask, Account.pSid)
        If lResult = 0 Then
            '//DACL type ACCESS_ALLOWED_ACE_TYPE was not added to DACL: " & Account.pSid & " Err# 917
            lErrNum = 917
            GoTo Handler
        End If
        lAceIndex = lAceIndex + 1
    End If

    '//copy inherited ACEs to DACL
    If (lDaclPresent <> 0 And lPAcl <> 0 And sACLInfo.AceCount > 0) Then
        '//get each ACE with inheritence flag and add to DACL
        For lCount = lCount To (sACLInfo.AceCount - 1)
            '//get next ACE
            lResult = GetAce(lPAcl, lCount, lPtr)
            If lResult = 0 Then
                '//ACE entry could not be allocated.  Err# 918
                lErrNum = 918
                GoTo Handler
            End If

            CopyMemory tempAce, lPtr, LenB(tempAce)
            '//add to ACL
            lResult = AddAce(lNewACL, ACL_REVISION, MAXDWORD, lPtr, tempAce.Header.AceSize)
            If lResult = 0 Then
                '//ACL build failure. Err# 919
                lErrNum = 919
                GoTo Handler
            End If
            lAceIndex = lAceIndex + 1
        Next lCount
    End If

    '//get descriptor attributes
    If lOldSD <> 0 Then
        lResult = GetSecurityDescriptorControl(lOldSD, lFlag, lRevision)
        If lResult <> 0 Then
            If ((lFlag And SE_DACL_AUTO_INHERITED) = SE_DACL_AUTO_INHERITED) Then
                lControlBits = SE_DACL_AUTO_INHERIT_REQ Or SE_DACL_AUTO_INHERITED
                lControlSet = lControlBits
            ElseIf ((lFlag And SE_DACL_PROTECTED) = SE_DACL_PROTECTED) Then
                lControlBits = SE_DACL_PROTECTED
                lControlSet = lControlBits
            End If
            If lControlSet <> 0 Then
                lResult = SetSecurityDescriptorControl(lPsd, lControlBits, lControlSet)
                If lResult = 0 Then
                    '//Failed to set Descriptor control entry. Err# 920
                    lErrNum = 920
                    GoTo Handler
                End If
            End If
        End If
    End If

    '//add DACL to security descriptor
    lResult = SetSecurityDescriptorDacl(lPsd, 1, lNewACL, 0)
    If lResult = 0 Then
        '//Failed to add DACL to Descriptor. Err# 921
        lErrNum = 921
        GoTo Handler
    End If

    lReturn = 1

Handler:
    '//free SIDs
    With Account
        If .pSid <> 0 And Not (.SidPassedByCaller) Then
            LocalFree (.pSid)
            .pSid = 0
        End If
    End With

    '//free the structures
    If lReturn = 0 Then
        If uInfo.pSD <> 0 Then
            LocalFree uInfo.pSD
        End If
        uInfo.pSD = 0
        If uInfo.pAcl <> 0 Then
            LocalFree uInfo.pAcl
        End If
        uInfo.pAcl = 0
    End If

    Set_Descriptor = lReturn

End Function

Private Function Build_New_ACE(ByVal pNewACL As Long, _
                               ByVal AceType As Byte, _
                               ByVal AceFlags As Byte, _
                               ByVal AccessMask As Long, _
                               ByVal pSid As Long, _
                               Optional ByVal lErrNum As Long) As Long

Dim lResult  As Long
Dim lAceLen  As Long
Dim lSidLen  As Long
Dim tTempAce As udtACE
Dim lAce     As Long

On Error GoTo Handler

    '//get len of ACL and SID
    lSidLen = GetLengthSid(pSid)
    lAceLen = Len(tTempAce) + lSidLen - 4

    '//create the buffer
    lAce = LocalAlloc(lPtr, lAceLen)
    If lAce = 0 Then
        '//LocalAlloc lAce returned zero. Err# 914
        lErrNum = 914
        GoTo Handler
    End If

    '//add ACE to structure
    With tTempAce
        .Header.AceType = AceType
        .Header.AceFlags = AceFlags
        .Header.AceSize = lAceLen
        .Mask = AccessMask
    End With
    '//copy SID and ACE info
    CopyMemory ByVal lAce, VarPtr(tTempAce), LenB(tTempAce)
    CopyMemory ByVal lAce + 8, pSid, lSidLen

    '//add ACE to ACL
    lResult = AddAce(pNewACL, ACL_REVISION, MAXDWORD, lAce, lAceLen)
    LocalFree lAce

Handler:
    Build_New_ACE = lResult

End Function

Private Function Is_Equal(Account As udtAccountPerm, _
                          pSid As Long) As Boolean

'//test for current sid
    Is_Equal = (EqualSid(Account.pSid, pSid))

End Function
