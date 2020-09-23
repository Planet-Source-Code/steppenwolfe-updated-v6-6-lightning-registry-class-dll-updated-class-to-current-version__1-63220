VERSION 5.00
Begin VB.Form frmNotify 
   BorderStyle     =   0  'None
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   3585
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Label lblNotice 
         AutoSize        =   -1  'True
         Caption         =   "Creating a System Restore Point.."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblNotice 
         AutoSize        =   -1  'True
         Caption         =   "MRU Pro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   60
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
