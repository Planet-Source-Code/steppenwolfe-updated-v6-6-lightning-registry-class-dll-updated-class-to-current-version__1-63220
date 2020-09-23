VERSION 5.00
Begin VB.Form frmBenchmark 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registry Access Bench Mark"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   435
      Left            =   4920
      TabIndex        =   2
      Top             =   2340
      Width           =   1125
   End
   Begin VB.OptionButton optBenchMark 
      Caption         =   "Using Async MultiThread"
      Height          =   255
      Index           =   1
      Left            =   630
      TabIndex        =   1
      Top             =   1590
      Width           =   2265
   End
   Begin VB.OptionButton optBenchMark 
      Caption         =   "Using a Standard Module"
      Height          =   255
      Index           =   0
      Left            =   630
      TabIndex        =   0
      Top             =   1200
      Value           =   -1  'True
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   $"frmBenchmark.frx":0000
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   150
      Width           =   5295
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   4
      Top             =   2520
      Width           =   390
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      Caption         =   "Keys Enumerated:"
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   3
      Top             =   2190
      Width           =   1290
   End
End
Attribute VB_Name = "frmBenchmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tBenchTime As Double
'Private WithEvents mMTR As clsMTReg


Private Sub cmdStart_Click()
Dim lCount As Long

    lblStat(0).Caption = "Keys Enumerated:"
    lblStat(1).Caption = "Time:"
    tBenchTime = 0
    tBenchTime = Timer
    lModKeyCount = 0
    
    Select Case True
        Case optBenchMark(0)
            '//in a module
            Mod_Recurse_Keys HKEY_LOCAL_MACHINE, "SOFTWARE"
            lblStat(0).Caption = "Keys Enumerated: " & lModKeyCount
            lblStat(1).Caption = "Time: " & _
            Format(Timer - tBenchTime, "#0.0000") & " Seconds"
            
        Case optBenchMark(1)
            '//multithread
            With mMTR
                .Root = HKEY_LOCAL_MACHINE
                .SubKey = "SOFTWARE"
                .Start
            End With
    End Select
    
End Sub

Private Sub cmdReg_Click()
    
    lblStat(0).Caption = "Keys Enumerated:"
    lblStat(1).Caption = "Time:"
    tBenchTime = 0
    tBenchTime = Timer
    
    '//set variables and start
   With mMTR
        .Root = HKEY_LOCAL_MACHINE
        .SubKey = "SOFTWARE"
        .Start
    End With

Handler:

End Sub

Private Sub Form_Load()
'// if you get an error, then you have
'// not compiled the prjMTReg project
'// and added the exe as a reference!

   Set mMTR = New clsMTReg
End Sub


Private Sub mMTR_Complete()

'//raised event at completion
lblStat(0).Caption = "Keys Enumerated: " & mMTR.cGKeyList.Count
lblStat(1).Caption = "Time: " & Format(Timer - tBenchTime, "#0.0000") & " Seconds"

End Sub


