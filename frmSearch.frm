VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lightning! V.2 - MRU Cleaner Pro"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      Caption         =   "Examples"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9420
      TabIndex        =   22
      Top             =   6930
      Width           =   915
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Recovery"
      Height          =   375
      Index           =   4
      Left            =   8490
      TabIndex        =   21
      Top             =   6930
      Width           =   885
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Clean MRUs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7440
      TabIndex        =   9
      Top             =   6930
      Width           =   1005
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Deselect"
      Height          =   375
      Index           =   2
      Left            =   6450
      TabIndex        =   8
      Top             =   6930
      Width           =   945
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Select All"
      Height          =   375
      Index           =   1
      Left            =   5430
      TabIndex        =   7
      Top             =   6930
      Width           =   975
   End
   Begin VB.Frame frmStatus 
      Caption         =   "Status"
      Height          =   1185
      Left            =   4350
      TabIndex        =   3
      Top             =   5520
      Width           =   6015
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Keys Enumerated:"
         Height          =   195
         Index           =   3
         Left            =   2550
         TabIndex        =   10
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Time to Scan:"
         Height          =   195
         Index           =   2
         Left            =   2550
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "MRUs Found:"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Lists Found:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   360
         Width           =   870
      End
   End
   Begin MSComctlLib.ListView lstResults 
      Height          =   5205
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   9181
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame frmSearch 
      Caption         =   "Search Options"
      Height          =   1815
      Left            =   210
      TabIndex        =   1
      Top             =   5520
      Width           =   3885
      Begin VB.PictureBox picBg 
         BorderStyle     =   0  'None
         Height          =   1605
         Left            =   60
         ScaleHeight     =   1605
         ScaleWidth      =   3795
         TabIndex        =   11
         Top             =   180
         Width           =   3795
         Begin VB.CheckBox chkOptions 
            Caption         =   "Run MRUs"
            Height          =   165
            Index           =   7
            Left            =   2040
            TabIndex        =   20
            Top             =   750
            Width           =   1395
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Media Player"
            Height          =   165
            Index           =   6
            Left            =   2040
            TabIndex        =   19
            Top             =   450
            Width           =   1395
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "WordPad MRUs"
            Height          =   165
            Index           =   5
            Left            =   2040
            TabIndex        =   18
            Top             =   150
            Width           =   1545
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Macromedia"
            Height          =   165
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   1395
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Adobe"
            Height          =   165
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   1050
            Width           =   1395
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Internet Explorer"
            Height          =   165
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   750
            Width           =   1635
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Recent Documents"
            Height          =   165
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   450
            Width           =   1755
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Office MRUs"
            Height          =   165
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   150
            Width           =   1395
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Search All MRUs"
            Height          =   225
            Left            =   2040
            TabIndex        =   12
            Top             =   1020
            Width           =   1665
         End
      End
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Start Scan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4350
      TabIndex        =   0
      Top             =   6930
      Width           =   1035
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'// Active-x Registry control and MRU Search Project - Revised November 10, 2005
'// November 10, 2005
'// After a short run, I am re-releasing this code with many major updates..

'// Posted this on the 8th, and pulled it a few hours lator, because I wanted to
'// write the active-x library it to its real potential..
'// So have a second look, it was completely rewritten..
'// also added a module that demostrates registry permissions, it was
'// taken from another of my projects, (unpuplished).
'// Author: John Underhill aka Steppenwolfe
'// For a comment (or a job.. ;o) email: steppenwolfe_2000@yahoo.com

'// First off, time to cover my ass..
'** This project and all the code it contains is licenced under a relaxed GNU.
'** This project is for educational purposes only
'** By using this code, the user agrees to the following:
'** -The author is in no way responsible for anything pertaining to the use
'** of this code.
'** -The user agrees to assume all responsibility for anything this project
'** might do, or might not do during its operation.
'** -The author makes absolutely no guarantees as to the fitness of this code, and no
'** warranties or responsibilities are expressed or implied.

'// And because I expect to see a dozen new 'MRU Assasin' type sharewares on download.com
'// in the following months, this bit is applicable too..

'** -You may use this code in your personal projects in any way you wish, including
'** use of the active-x control.
'** -If you intend on publishing this code, in whole or in part, as part of a commercial project
'** the author expects to be notified.
'** Publishing of this project in its current form, (as an MRU scanner with -this- interface),
'** is not permitted.
'** -If you want to use this code for a commercial project, you must give credit to the
'** author in the read me/about dialog/ and-or /help file of the application. (Example: Registry
'** controls courtesy of John Underhill)

'// Ok, with that piece of nastiness behind us, lets get on with the show.. `;-}

'// You might, (as most people certainly are..), be a little leary of programming as it
'// pertains to the Windows Registry. Who can blame you? The registry is a very complex
'// database, with many mysteries, and in all the years I have been working in the heart of it,
'// I still do not know what half of it does. But the registry, in its essence, is just a database.
'// Used to store everything from persistant settings to raw security data.
'// Many people are under the impression that if you delete or change the wrong value, it
'// will immediately make the system inoperative. I assure you, this is rarely the case. In fact,
'// many of the system settings that are most vulnerable, are permissions restricted to
'// system only access, and can not be changed by user invoked api, (unless you know about
'// token impersonation and NTFS api, but that project, is for another day..)
'// This is not to say that you shouldn't be careful about what you do, but working with
'// the registry is not the minefield that most people think it is..

'// About the MRU program:
'// MRUs (Most Recently Used) lists are used by programs to keep track of recent object access,
'// like the last several documents accessed, or recent images viewed.
'// MRUs are considered by some to be a privacy concern, because a computer accessible to other
'// users will list the usage of objects by the application, (ie the wife scutinizing your porn stash :-( .
'// I looked at a number of MRU cleaners, the best of which was MRU Blaster by JavaCool, but they all
'// had the same flaw, the reliance on a static list of MRU locations.
'// The MRU Pro uses no lists, unlike all the professional removal tools I have seen, but relies
'// on pattern matching to 'predict' if an entry is an MRU.
'// After looking at the registry entries, I saw several patterns amoung most developers in writing
'// these lists, the key path usually contained a signifier or list identifier, most common of which were:
'// mru, history, list, file, recent and 'last'. The MRU values themselves also had a loose series of
'// conventions, they are: Filex, urlx, a-z, 0-99, and '000-999'.
'// So what I did was used both of these patterns, the first to cull likely key candidates, the second match
'// to find MRU value entries under that key. I further filter word in word matches with maximum value length filters,
'// and added specific type filters for common applications.
'// If you want to develop this further, feel free to do so, but for now..
'// The MRU Pro is really just for demostration purposes, and is meant only to demonstrate the
'// speed and power of the accompanying active-x registry control.

'// About the Active-x control
'// Completely rewrote active-x library and wrapper, now every imaginable registry function is exposed
'// big_endian, little_endian, sz, expand_sz, mult_sz, list, binary, dword and much more..
'// added error handling and dll error interpretation to dll wrapper.
'// Expanded and improved filter functions in mru class, and fixed small bugs
'// I pulled the original a couple hours after I posted it, because I knew it could be made into an impressive
'// piece of code with a little more work, so take another look, and let me know what you think..

'// What to do..
'// 1) Compile the active-x control into the release folder
'// 2) Run the scan

'// I have tested this thoroughly on my own box, (xp sp2), and all is well, so enjoy..

'// If you want to use this, or think it is good code, all I ask is that you take the time to vote..
'// Cheers
'// John

Option Explicit
'//events
Private WithEvents cMruSearch As clsMrusearch
Attribute cMruSearch.VB_VarHelpID = -1
Public Event Process(percent As Long)
Public Event Maxcount(maxnum As Long)
Public Event Listcount(listnum As Long)
Public Event Valcount(valnum As Long)
Public Event Labelchange(sCaption As String)
Public Event Keycount(keynum As Long)

'//counters
Private lGCount         As Long
Private lListcount      As Long
Private lValCount       As Long
Private lKeyNum         As Long
Private tSearchTime     As Double

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                       FORM CONTROLS
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Private Sub cmdControl_Click(Index As Integer)

Dim chk     As CheckBox
Dim bTest   As Boolean
Dim i       As Long
Dim msg     As Integer

    Select Case Index
        '//test for options
        Case 0
            For Each chk In chkOptions
                If chk.Value = 1 Then
                    bTest = True
                    Exit For
                End If
            Next chk
            
            If Not bTest Then
                MsgBox "Please select a search option to proceed!", vbExclamation, "No Options!"
                Exit Sub
            End If
            
            '//start scan
            Start_Scan
        '//select all
        Case 1
            With lstResults
                For i = 1 To .ListItems.Count Step 1
                    .ListItems(i).Checked = True
                Next i
                i = 0
            End With
        '//deselect all
        Case 2
            With lstResults
                For i = 1 To .ListItems.Count Step 1
                    .ListItems(i).Checked = False
                Next i
                i = 0
            End With
        '//remove selected
        Case 3
            '//verify choice
            msg = MsgBox("Do you want to Delete the Selected MRU entries from the Registry?" & vbNewLine _
            & "Click Yes to Proceed, or No to Cancel.", vbYesNo, "Delete MRU Entries!")
            If Not msg = 6 Then
                Exit Sub
            End If
            
            '//check restore type
            '//use system restore
            If bRestoreType Then
                frmNotify.Show vbModeless, Me
                '//if restore fails, abort
                If Not Start_Restore("MRUPro") Then
                    MsgBox "The System Restore operation has failed! Aborting Removal!", _
                    vbCritical, "System Restore Failure!"
                    Unload frmNotify
                    Exit Sub
                End If
                SaveSetting App.EXEName, "MRUPro", "lblsta1", "Last Restore: " & Format(Now, "yyyy mm dd")
        
            Else
                '//use snapshot
                With frmNotify
                    .Show vbModeless, Me
                    .lblNotice(1).Caption = "Creating a Registry Snapshot.."
                    '//if backup fails, abort
                    With New clsRegHandler
                        If Not .Save_Key(HKEY_CURRENT_USER, "Software", App.Path & "\regback.kbs") Then
                            MsgBox "The key backup operation has failed! Aborting Removal!", _
                            vbCritical, "Key Backup Failure!"
                            Unload frmNotify
                            Exit Sub
                        End If
                    End With
                End With
                SaveSetting App.EXEName, "MRUPro", "lblsta2", "Last Snapshot: " & Format(Now, "yyyy mm dd")
            End If
            
            Unload frmNotify
            Remove_MRU
            
        Case 4
            frmRecovery.Show vbModeless, Me
            
    End Select
    
End Sub

Private Sub cmdShow_Click()
    frmRegDemo.Show
End Sub

Private Sub Form_Load()

    Get_Scanner_Options
    '//disable removal option
    cmdControl(3).Enabled = False
    '//set up listview
    With lstResults
        .View = lvwReport
        .Checkboxes = True
        .ListItems.Clear
        .AllowColumnReorder = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, Text:="Key", Width:=(.Width / 8) * 6
        .ColumnHeaders.Add 2, Text:="Application", Width:=(.Width / 8)
        .ColumnHeaders.Add 3, Text:="Value", Width:=(.Width / 8) - 100
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    '//set persistent options
    Set_Scanner_Options
    '//unload all forms
    Application_Terminate
End Sub

Private Sub chkAll_Click()

Dim chk As CheckBox

    For Each chk In chkOptions
        If chkAll.Value = 1 Then
            chk.Value = 1
        Else
            chk.Value = 0
        End If
    Next chk
    
End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                       CORE PROCESSORS
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Private Sub Remove_MRU()
'//removal routine

Dim l           As Long
Dim lCounter    As Long

On Error Resume Next

    With New clsRegHandler
        With frmProgress
            .Show vbModeless, Me
            .lblTitle.Caption = "MRU Deletion Progress"
            .lblStatus(3).Caption = "Deleting selected MRUs.."
            .prgScan.Max = lstResults.ListItems.Count
        End With
        For l = lstResults.ListItems.Count To 1 Step -1
            If lstResults.ListItems(l).Checked = True Then
                Debug.Print lstResults.ListItems(l).Text, Trim$(lstResults.ListItems(l).SubItems(2))
                '** I remmed the deletion calls because..
                '** if you are not stepping through the code
                '** and spot this, then you should not be deleting objects
                '** from the registry! Pay attention people!!
                '** ran it on my machine, and had no problems..
                If Not Len(Trim$(lstResults.ListItems(l).SubItems(2))) = 1 Then
                '** to enable val deletions, unrem this bit
                    '.Delete_Value HKEY_CURRENT_USER, Trim$(lstResults.ListItems(l).Text), Trim$(lstResults.ListItems(l).SubItems(2))
                Else
                    '//test for mru index entry, if present, reset to nothing
                    If .Search_Value(HKEY_CURRENT_USER, Trim$(lstResults.ListItems(l).Text), "MRUList") Then
                        '** and unrem this..
                        '.Write_Value HKEY_CURRENT_USER, Trim$(lstResults.ListItems(l).Text), "MRUList", "", 1
                    End If
                End If
                    '** and this..
                    '.Delete_Value HKEY_CURRENT_USER, Trim$(lstResults.ListItems(l).Text), Trim$(lstResults.ListItems(l).SubItems(2))
                lstResults.ListItems.Remove (l)
            End If
            lCounter = lCounter + 1
            With frmProgress
                .prgScan.Value = lCounter
                .lblStatus(0).Caption = "Deleting: " & lCounter
            End With
            DoEvents
        Next l
    End With
    
    Unload frmProgress
    If bRestoreType Then
        End_Restore
    End If

On Error GoTo 0

End Sub

Private Sub Start_Scan()
'//core for scan routines
Dim cVal        As Variant
Dim l           As Long
Dim lVal        As Long
Dim sFormat     As String
Dim tpeChoice   As tChoice
Dim lItem       As ListItem

    lstResults.ListItems.Clear
    cmdControl(3).Enabled = False
    lblStatus(0).Caption = "Lists Found:"
    lblStatus(1).Caption = "MRUs Found:"
    lblStatus(2).Caption = "Time to Scan:"
    lblStatus(3).Caption = "Keys Enumerated:"
    
    Set cMruSearch = New clsMrusearch
    With cMruSearch
        '//get the options settings for filter routine
        '//and fill structure, pass structure to class
        '//with friend sub
        With tpeChoice
            If chkOptions(0).Value = 1 Then .tChoice0 = True
            If chkOptions(1).Value = 1 Then .tChoice1 = True
            If chkOptions(2).Value = 1 Then .tChoice2 = True
            If chkOptions(3).Value = 1 Then .tChoice3 = True
            If chkOptions(4).Value = 1 Then .tChoice4 = True
            If chkOptions(5).Value = 1 Then .tChoice5 = True
            If chkOptions(6).Value = 1 Then .tChoice6 = True
            If chkOptions(7).Value = 1 Then .tChoice7 = True
            If chkAll.Value = 1 Then .tChoice8 = True
        End With
        '//start timer
        tSearchTime = Timer
        '//pass the structure
        .Pass_UDT tpeChoice
        With frmProgress
            .Show vbModeless, Me
            .lblStatus(3).Caption = "Enumerating Registry.."
        End With
        For Each cVal In .Mru_Control(HKEY_CURRENT_USER, "Software")
            '//put results in list
            With lstResults
                Set lItem = .ListItems.Add(Text:=Left$(cVal, InStr(cVal, Chr(31)) - 1))
                lItem.SubItems(1) = Mid(lItem.Text, (InStrRev(lItem.Text, Chr(92)) + 1))
                lItem.SubItems(2) = Mid(cVal, InStrRev(cVal, Chr(31)) + 1)
            End With
        Next
        '//timer status
        lblStatus(2).Caption = "Time to Scan: " & Format(Timer - tSearchTime, "#0.0000") & " Seconds"
        Unload frmProgress
    End With
    
    cmdControl(3).Enabled = True
    Set cMruSearch = Nothing

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                       STATUS EVENTS
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Public Sub cMruSearch_Process(ByVal percent As Long)
'//progress bar event

On Error Resume Next

    RaiseEvent Process(percent)
    With frmProgress
        .prgScan.Value = percent
        .lblStatus(0).Caption = Format((percent / lGCount) * 100, "###") & " %"
    End With
    
On Error GoTo 0

End Sub

Public Sub cMruSearch_Maxcount(ByVal maxnum As Long)
'//max progress bar value

On Error Resume Next

    RaiseEvent Maxcount(maxnum)
    With frmProgress
        .prgScan.Max = maxnum
        lGCount = maxnum
    End With

On Error GoTo 0

End Sub

Public Sub cMruSearch_Listcount(ByVal listnum As Long)
'//mru list count

On Error Resume Next

    RaiseEvent Listcount(listnum)
    lListcount = lListcount + listnum
    With Me
        .lblStatus(0).Caption = "Number of Lists Found: " & lListcount
    End With
    
On Error GoTo 0

End Sub

Public Sub cMruSearch_ValCount(ByVal valnum As Long)
'//mru val count

On Error Resume Next

    RaiseEvent Valcount(valnum)
    With Me
    lValCount = valnum
        .lblStatus(1).Caption = "Number of MRUs Found: " & lValCount
    End With
    
On Error GoTo 0

End Sub

Public Sub cMruSearch_Labelchange(ByVal sCaption As String)
'//scanning status update

On Error Resume Next

    RaiseEvent Labelchange(sCaption)
    With frmProgress
        .lblStatus(3).Caption = "Status: " & sCaption
    End With
    
On Error GoTo 0

End Sub

Public Sub cMruSearch_Keycount(ByVal keynum As Long)
'//number of keys enumerated

On Error Resume Next

    RaiseEvent Keycount(keynum)
    lKeyNum = lKeyNum + keynum
    With Me
        .lblStatus(3).Caption = "Keys Enumerated: " & lKeyNum
    End With
    With frmProgress
        .lblStatus(0).Caption = "Key Enumeration Count: " & lKeyNum
    End With
    
On Error GoTo 0

End Sub

