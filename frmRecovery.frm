VERSION 5.00
Begin VB.Form frmRecovery 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Recovery Controls"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
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
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBg 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   90
      ScaleHeight     =   1605
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   90
      Width           =   4515
      Begin VB.Frame frmXp 
         Caption         =   "Recovery Method"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   90
         TabIndex        =   4
         Top             =   60
         Width           =   2025
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   90
            ScaleHeight     =   585
            ScaleWidth      =   1875
            TabIndex        =   5
            Top             =   210
            Width           =   1875
            Begin VB.OptionButton optRestore 
               Caption         =   "Use Snapshot"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   60
               TabIndex        =   7
               Top             =   360
               Width           =   1395
            End
            Begin VB.OptionButton optRestore 
               Caption         =   "Use System Restore"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   6
               Top             =   60
               Value           =   -1  'True
               Width           =   1815
            End
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Restore Registry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2940
         TabIndex        =   2
         Top             =   1230
         Width           =   1455
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Create Snapshot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1410
         TabIndex        =   1
         Top             =   1230
         Width           =   1455
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Last Snapshot: Pending"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   2250
         TabIndex        =   9
         Top             =   750
         Width           =   1710
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Last Restore: Not Available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   2250
         TabIndex        =   8
         Top             =   450
         Width           =   1965
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "System Restore: Not Available"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   2250
         TabIndex        =   3
         Top             =   150
         Width           =   2190
      End
   End
End
Attribute VB_Name = "frmRecovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRestore_Click(Index As Integer)
'//user dialogs for backup and restore of reg keys
Dim msg As Integer

    Select Case Index
        Case 0
            msg = MsgBox("Do you want to Create a backup of the HKCU\Software Key?" & vbNewLine _
            & "Click Yes to Proceed, or No to Cancel.", vbYesNo, "Create Key Backup")
            If msg = 6 Then
                With New clsRegHandler
                    If Not .Save_Key(HKEY_CURRENT_USER, "Software", App.Path & "\regback.kbs") Then
                        MsgBox "The key backup operation has failed! Aborting backup!", _
                        vbCritical, "Key Backup Failure!"
                    End If
                End With
            End If
        
        Case 1
            msg = MsgBox("Do you want to Restore to the previous HKCU\Software Key?" & vbNewLine _
            & "Click Yes to Proceed, or No to Cancel.", vbYesNo, "Create Key Backup")
            If msg = 6 Then
                With New clsRegHandler
                    If Not .Restore_Key(HKEY_CURRENT_USER, "Software", App.Path & "\regback.kbs") Then
                        MsgBox "The key Restoration operation has failed! Aborting Restore!", _
                        vbCritical, "Key Restore Failure!"
                    End If
                End With
            End If
    End Select
    
End Sub

Private Sub Form_Load()
    Get_Recovery_Options
    Restore_Status
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set_Recovery_Options
End Sub
