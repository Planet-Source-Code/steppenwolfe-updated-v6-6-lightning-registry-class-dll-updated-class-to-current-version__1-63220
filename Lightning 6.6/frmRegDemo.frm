VERSION 5.00
Begin VB.Form frmRegDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lightning 1.6.5  Demonstration Routines"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frType 
      Caption         =   "Value Controls"
      Height          =   3225
      Index           =   2
      Left            =   2220
      TabIndex        =   18
      Top             =   3690
      Width           =   8025
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read QWord"
         Height          =   375
         Index           =   37
         Left            =   6510
         TabIndex        =   48
         Top             =   810
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write Qword"
         Height          =   375
         Index           =   36
         Left            =   6510
         TabIndex        =   47
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read Res Req"
         Height          =   375
         Index           =   35
         Left            =   4980
         TabIndex        =   45
         Top             =   2610
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write Res Req"
         Height          =   375
         Index           =   34
         Left            =   4980
         TabIndex        =   44
         Top             =   2160
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read Res List"
         Height          =   375
         Index           =   33
         Left            =   4980
         TabIndex        =   43
         Top             =   1710
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write Res List"
         Height          =   375
         Index           =   32
         Left            =   4980
         TabIndex        =   42
         Top             =   1260
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read Res Desc"
         Height          =   375
         Index           =   31
         Left            =   4980
         TabIndex        =   41
         Top             =   810
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write Res Desc"
         Height          =   375
         Index           =   30
         Left            =   4980
         TabIndex        =   40
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read Link"
         Height          =   375
         Index           =   29
         Left            =   3450
         TabIndex        =   39
         Top             =   2610
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write Link"
         Height          =   375
         Index           =   28
         Left            =   3450
         TabIndex        =   38
         Top             =   2160
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read MultiCN"
         Height          =   375
         Index           =   27
         Left            =   3450
         TabIndex        =   37
         Top             =   1710
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write MultiCN"
         Height          =   375
         Index           =   26
         Left            =   3450
         TabIndex        =   36
         Top             =   1260
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read Multi"
         Height          =   375
         Index           =   25
         Left            =   3450
         TabIndex        =   35
         Top             =   810
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write Multi"
         Height          =   375
         Index           =   24
         Left            =   3450
         TabIndex        =   34
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Search Value"
         Height          =   375
         Index           =   17
         Left            =   6510
         TabIndex        =   33
         Top             =   1710
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "List Values"
         Height          =   375
         Index           =   15
         Left            =   6510
         TabIndex        =   32
         Top             =   1260
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Delete Value"
         Height          =   375
         Index           =   14
         Left            =   6510
         TabIndex        =   31
         Top             =   2160
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read String"
         Height          =   375
         Index           =   13
         Left            =   1860
         TabIndex        =   30
         Top             =   2610
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write String"
         Height          =   375
         Index           =   12
         Left            =   1860
         TabIndex        =   29
         Top             =   2160
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read BEndian"
         Height          =   375
         Index           =   11
         Left            =   1860
         TabIndex        =   28
         Top             =   1710
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write BEndian"
         Height          =   375
         Index           =   10
         Left            =   1860
         TabIndex        =   27
         Top             =   1260
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read LEndian"
         Height          =   375
         Index           =   9
         Left            =   1860
         TabIndex        =   26
         Top             =   810
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write LEndian"
         Height          =   375
         Index           =   8
         Left            =   1860
         TabIndex        =   25
         Top             =   360
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read Expanded"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   24
         Top             =   2610
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write Expanded"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read Dword"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Top             =   1710
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write Dword"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Top             =   1260
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Read Binary"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   810
         Width           =   1305
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Write Binary"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Frame frType 
      Caption         =   "Security Controls"
      Height          =   3225
      Index           =   3
      Left            =   10410
      TabIndex        =   13
      Top             =   3690
      Width           =   2055
      Begin VB.CommandButton cmdControl 
         Caption         =   "Test User Access"
         Height          =   375
         Index           =   22
         Left            =   210
         TabIndex        =   14
         Top             =   360
         Width           =   1635
      End
   End
   Begin VB.Frame frType 
      Caption         =   "Settings"
      Height          =   3405
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   12315
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   3630
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   480
         Width           =   8475
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Text            =   "test data"
         Top             =   2310
         Width           =   3015
      End
      Begin VB.TextBox txtResult 
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Text            =   "results"
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Text            =   "test value"
         Top             =   1710
         Width           =   3015
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Text            =   "Software\MRU Pro"
         Top             =   1110
         Width           =   2955
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Text            =   "HKEY_CURRENT_USER"
         Top             =   510
         Width           =   2955
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   17
         Top             =   2670
         Width           =   450
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   9
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   7
         Top             =   1500
         Width           =   405
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Sub Key"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   6
         Top             =   900
         Width           =   600
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Root Key"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   300
         Width           =   660
      End
   End
   Begin VB.Frame frType 
      Caption         =   "Key Controls"
      Height          =   3225
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   3690
      Width           =   1905
      Begin VB.CommandButton cmdControl 
         Caption         =   "Key Exist"
         Height          =   375
         Index           =   21
         Left            =   270
         TabIndex        =   15
         Top             =   1710
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "List Keys"
         Height          =   375
         Index           =   16
         Left            =   270
         TabIndex        =   12
         Top             =   1260
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Delete Key"
         Height          =   375
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   810
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         Caption         =   "Create Key"
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   360
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmRegDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cLightning As clsLightning


Private Sub cLightning_ErrorCond(ByVal sRoutine As String, _
                                 ByVal sKey As String, _
                                 ByVal sError As String)

    MsgBox "Error: " + sRoutine + " failed to access key: " + sKey + " Code# " + sError, vbExclamation, "Error"
    
End Sub

Private Sub cmdControl_Click(Index As Integer)

On Error Resume Next

    With cLightning
        Select Case Index
        
            '//create key
            Case 0
                If .Create_Key(HKEY_CURRENT_USER, "Software\MRU Pro\Test Key") Then
                    txtResult.Text = "Key Created"
                Else
                    txtResult.Text = "Key Was Not Created"
                End If
                
            '//delete key
            Case 1
                If .Delete_Key(HKEY_CURRENT_USER, "Software\MRU Pro\Test Key") Then
                    txtResult.Text = "Key Deleted"
                Else
                    txtResult.Text = "Key Was Not Deleted"
                End If
                
            '//write binary
            Case 2
                txtValue(3).Text = "00 01 00 01 10 11"
                .Write_Binary HKEY_CURRENT_USER, "Software\MRU Pro", "testbinary", txtValue(3).Text
                txtResult.Text = "Binary Value Written"
                
            '//read binary
            Case 3
                txtValue(3).Text = "Reading: testbinary"
                txtResult.Text = CStr(.Read_Binary(HKEY_CURRENT_USER, "Software\MRU Pro", "testbinary"))
            Case 4
            
            '//write dword
                txtValue(3).Text = "1171123"
                .Write_DWord HKEY_CURRENT_USER, "Software\MRU Pro", "testdword", CLng(txtValue(3).Text)
                txtResult.Text = "Dword Value Written"
                
            '//read dword
            Case 5
                txtValue(3).Text = "Reading: testdword"
                txtResult.Text = CStr(.Read_DWord(HKEY_CURRENT_USER, "Software\MRU Pro", "testdword"))
                
            '//write expanded string
            Case 6
                txtValue(3).Text = "an expanded string value"
                .Write_Expanded HKEY_CURRENT_USER, "Software\MRU Pro", "testexpand", txtValue(3).Text
                txtResult.Text = "Expanded String Value Written"
                
            '//read expand
            Case 7
                txtValue(3).Text = "Reading: testexpand"
                txtResult.Text = CStr(.Read_String(HKEY_CURRENT_USER, "Software\MRU Pro", "testexpand"))
                
            '//write little endian
            Case 8
                txtValue(3).Text = "171777"
                .Write_LEndian HKEY_CURRENT_USER, "Software\MRU Pro", "testlittleend", CLng(txtValue(3).Text)
                txtResult.Text = "Little Endian Value Written"
                
            '//read little endian
            Case 9
                txtValue(3).Text = "Reading: testlittleend"
                txtResult.Text = CStr(.Read_LEndian(HKEY_CURRENT_USER, "Software\MRU Pro", "testlittleend"))
                
            '//write big endian
            Case 10
                Dim vBEnd As Variant
                txtValue(3).Text = "171717"
                '//convert to big_endian value
                vBEnd = .Make_BEndian32(CLng("171777"))
                .Write_BEndian HKEY_CURRENT_USER, "Software\MRU Pro", "testbigendian", vBEnd
                txtResult.Text = "Big Endian Value Written"
                
            '//read big endian
            Case 11
                txtValue(3).Text = "Reading Value: testbigendian"
                txtResult.Text = CStr(.Read_BEndian(HKEY_CURRENT_USER, "Software\MRU Pro", "testbigendian"))
                
            '//write string
            Case 12
                txtValue(3).Text = "test string"
                .Write_String HKEY_CURRENT_USER, "Software\MRU Pro", "teststring", txtValue(3).Text
                txtResult.Text = "String Value Written"
                
            '//read string
            Case 13
                txtValue(3).Text = "Reading Value: teststring"
                txtResult.Text = CStr(.Read_String(HKEY_CURRENT_USER, "Software\MRU Pro", "teststring"))
                
            '//delete string
            Case 14
                txtValue(3).Text = "Deleting value: textexpand"
                If .Delete_Value(HKEY_CURRENT_USER, "Software\MRU Pro", "testexpand") Then
                    txtResult.Text = "String Value Was Deleted"
                Else
                    txtResult.Text = "The value Was Not Deleted"
                End If
                
            '//list values
            Case 15
                txtValue(3).Text = "Listing Values in Debug Window"
                Dim vVal As Variant
                Debug.Print "Listing Values in HKLM:"
                Debug.Print "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                For Each vVal In .List_Values(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
                    Debug.Print vVal
                Next
                
            '//list keys
            Case 16
                txtValue(3).Text = "Listing Keys in Debug Window"
                Dim vKey As Variant
                For Each vKey In .List_Keys(HKEY_CURRENT_USER, "Software")
                    Debug.Print vKey
                Next vKey
                
            '//search for a value
            '//test for presence of an application
            '//in this case MS Word
            Case 17
                Dim sVPath As String
                txtValue(3).Text = "Search For: Installed Version of MS Word"
                txtResult.Text = "MS Word was not found"
                .p_Intercept = False
                If .Search_Value(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\8.0\Word\InstallRoot", "path") Then
                    sVPath = .Read_String(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\8.0\Word\InstallRoot", "Path")
                    txtResult.Text = "Word Version is 8.0, location: " & sVPath
                End If
                If .Search_Value(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\9.0\Word\InstallRoot", "path") Then
                    sVPath = .Read_String(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\9.0\Word\InstallRoot", "Path")
                    txtResult.Text = "Word Version is 9.0, location: " & sVPath
                End If
                If .Search_Value(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\10.0\Word\InstallRoot", "path") Then
                    sVPath = .Read_String(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\10.0\Word\InstallRoot", "Path")
                    txtResult.Text = "Word Version is 10.0, location: " & sVPath
                End If
                If .Search_Value(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\11.0\Word\InstallRoot", "path") Then
                    sVPath = .Read_String(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\11.0\Word\InstallRoot", "Path")
                    txtResult.Text = "Word Version is 11.0, location: " & sVPath
                End If
                .p_Intercept = True
                
            '//key exists
            Case 21
                txtValue(3).Text = "Key Exists: Software\MRU Pro"
                If .Key_Exist(HKEY_CURRENT_USER, "Software\MRU Pro") Then
                    txtResult.Text = "The key exists"
                Else
                    txtResult.Text = "The key does not exist"
                End If

            '// key access
            Case 22
                If .Access_Test(HKEY_CURRENT_USER, "Software\MRU Pro") Then
                    txtResult.Text = "You have write access"
                Else
                    txtResult.Text = "You do not have write access"
                End If
            
            '//write multi
            Case 24
                Dim sMData As String
                sMData = "test1" & vbNullChar & "test2" & vbNullChar & "test3" & vbNullChar & vbNullChar
                txtValue(3).Text = "Write Multi Value: test1 test2 test3"
                .Write_Multi HKEY_CURRENT_USER, "Software\MRU Pro", "testmulti", sMData
                txtResult.Text = "Multi value has been added"
                
            '//read multi
            Case 25
                txtValue(3).Text = "Reading Value: testmulti"
                txtResult.Text = .Read_Multi(HKEY_CURRENT_USER, "Software\MRU Pro", "testmulti")
                
            'write multi from collection
            Case 26
                txtValue(3).Text = "Write Multi Collection Set: testcol1 testcol2 testcol3"
                Dim cMulti  As New Collection
                Set cMulti = New Collection
                cMulti.Add "testcol1"
                cMulti.Add "testcol2"
                cMulti.Add "testcol3"
                .Write_MultiCN HKEY_CURRENT_USER, "Software\MRU Pro", "testmulticn", cMulti
                Set cMulti = Nothing
                
            '//read multi with collection
            Case 27
                Dim vCol    As Variant
                Dim sColRes As String
                txtValue(3).Text = "Reading From Collection: testmulticn"
                 For Each vCol In .Read_MultiCN(HKEY_CURRENT_USER, "Software\MRU Pro", "testmulticn")
                    sColRes = sColRes & vCol & " "
                Next vCol
                txtResult.Text = sColRes
                
            '//write link value
            Case 28
                txtValue(3).Text = "00 01 00 01 10 11"
                .Write_Link HKEY_CURRENT_USER, "Software\MRU Pro", "testlink", txtValue(3).Text
                txtResult.Text = "Link Value Written"
                
            '//read link value
            Case 29
                txtValue(3).Text = "Reading: testlink"
                txtResult.Text = .Read_Link(HKEY_CURRENT_USER, "Software\MRU Pro", "testlink")
                
            '//write hardware resource description
            '//NOTE* Resource decriptors are interfaces
            '//between your hardware and system settings
            '//such as memory, IRQ and I/O allocations
            '//Adding a number of dummy values, could have
            '//serious consecquences..
            '//In other words, unless you are
            '//writing drivers, leave it alone..
            Case 30
                txtValue(3).Text = "00 01 00 01 10 11"
                .Write_ResDescriptor HKEY_CURRENT_USER, "Software\MRU Pro", "testresdesc", txtValue(3).Text
                txtResult.Text = "Resource Description Value Written"
                
            '//read hardware resource description
            Case 31
                txtValue(3).Text = "Reading: testresdesc"
                txtResult.Text = .Read_ResDescriptor(HKEY_CURRENT_USER, "Software\MRU Pro", "testresdesc")
                
            '//write hardware resource list
            '//NOTE* Same warning as above
            Case 32
                txtValue(3).Text = "00 01 00 01 15 11"
                .Write_ResourceList HKEY_CURRENT_USER, "Software\MRU Pro", "testreslist", txtValue(3).Text
                txtResult.Text = "Resource List Value Written"
                
            '//read hardware resource list
            Case 33
                txtValue(3).Text = "Reading: testreslist"
                txtResult.Text = .Read_ResourceList(HKEY_CURRENT_USER, "Software\MRU Pro", "testreslist")
                
            '//write hardware resource requirement
            '//NOTE* Same warning as above
            Case 34
                txtValue(3).Text = "00 01 00 01 10 11"
                .Write_ResRequired HKEY_CURRENT_USER, "Software\MRU Pro", "testresreq", txtValue(3).Text
                txtResult.Text = "Resource Requirement Value Written"
                
            '//read hardware resource requirement
            Case 35
                txtValue(3).Text = "Reading: testresreq"
                txtResult.Text = .Read_ResRequired(HKEY_CURRENT_USER, "Software\MRU Pro", "testresreq")
                
            '//write qword(64bit number)
            Case 36
                Dim cVal As Currency
                cVal = .Convert_Curr("999999999")
                .Write_QWord HKEY_CURRENT_USER, "Software\MRU Pro", "testqword", cVal
                txtResult.Text = "Qword Value Written"
                
            '//read qword
            Case 37
                Dim cSVal As Currency
                txtValue(3).Text = "Reading: testqword"
                cSVal = .Convert_Text(.Read_QWord(HKEY_CURRENT_USER, "Software\MRU Pro", "testqword"))
                txtResult.Text = cSVal
        End Select
    End With
    
On Error GoTo 0

End Sub

Private Sub Form_Load()

Dim sDesc   As String

On Error GoTo Handler

    Set cLightning = New clsLightning
    
    '/* enable error trapping
    With cLightning
        .p_Intercept = True
        .p_Logging = True
        .p_Notify = True
    End With
    
    sDesc = "** List of exposed functions **" & vbNewLine & vbNewLine & _
            "Access_Check - Test user access rights" & vbNewLine & "Read_BEndian - read a big endian value" & vbNewLine & _
            "Write_BEndian - write a big_endian value" & vbNewLine & "Read_Binary - read a binary value" & vbNewLine & _
            "Write_Binary - write a binary value" & vbNewLine & "Read_Dword - read a dword value" & vbNewLine & _
            "Write_Dword - write a dword value" & vbNewLine & "Read_Link - read a binary link value" & vbNewLine & _
            "Write_Link - write a binary link value" & vbNewLine & "List_Values - puts all of a keys values into a collection" & vbNewLine & _
            "Read_LEndian - read a little endian value" & vbNewLine & "Write_LEndian - write a little_endian value" & vbNewLine & _
            "Read_Multi - read a multi_sz value" & vbNewLine & "Write_Multi - write a multi_sz value" & vbNewLine & _
            "Write_MultiCN - converts a collection into a multi_sz value" & vbNewLine & "Read_ResDesc - read hardware resource description" & vbNewLine & _
            "Write_ResDesc = write hardware resource description" & vbNewLine & "Read_ResList - read a hardware resource list" & vbNewLine & _
            "Write_ResList - write to a hardware resource list" & vbNewLine & "Read_ResRequired - read a hardware resource requirements list" & vbNewLine & _
            "Write_ResRequired - write to a hardware resource requirements list" & vbNewLine & "Read_String - read a string(sz) or expanded string(expand_sz)" & vbNewLine & _
            "Write_String - write a string value" & vbNewLine & "Write_Expanded - write a expanded string" & vbNewLine & _
            "Read_Qword - read a 64bit number" & vbNewLine & "Write_Qword - write a 64bit number" & vbNewLine & _
            "List_Keys - puts all subkeys under specified branch into a collection" & vbNewLine & "Key_Exists - test if key exists" & vbNewLine & _
            "Create_Key - create a new key" & vbNewLine & "Delete_Key - delete a key" & vbNewLine & _
            "Write_Value - write value types: 1)sz 2)expand_sz 3)multi_sz 4)binary 5)dword 6)little_endian 7)big_endian" & vbNewLine & "Delete_Value - delete a value" & vbNewLine & _
            "Search_Value - search for a value under the key" & vbNewLine & "Make_LEndian16 - convert integer to 16bit little_endian" & vbNewLine & _
            "Make_LEndian32 - convert long to 32bit little_endian" & vbNewLine & "Make_BEndian32 - convert long to big endian format" & vbNewLine & _
            "Log_Error - sends errors to a log file" & vbNewLine & "Test_Access - test users level of access permissions" & vbNewLine & _
            "Set_Key_Permissions - change access permissions to a key" & vbNewLine & "Save_Key - save a binary image of a registry key" & vbNewLine & _
            "Restore_Key - restore key from binary image"

            txtDescription.Text = sDesc
Handler:

End Sub

Private Sub Form_Unload(Cancel As Integer)

'//Destroy app key on unload
'//you don't want bogus resource type
'//values in a key

On Error GoTo Handler

    With cLightning
        .Delete_Key HKEY_CURRENT_USER, "Software\MRU Pro"
    End With
    Set cLightning = Nothing

Handler:

End Sub
