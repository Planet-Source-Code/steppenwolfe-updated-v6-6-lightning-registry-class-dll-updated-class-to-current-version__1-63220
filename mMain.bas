Attribute VB_Name = "mMain"
Option Explicit

Public bOSVersion   As Boolean
Public bRestoreType As Boolean

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Sub Main()

Dim iccex As tagInitCommonControlsEx

On Error Resume Next

   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex

On Error GoTo 0
   frmSearch.Show
   
End Sub

Public Sub Application_Terminate()
'//close all windows and end

On Error Resume Next

Dim frm As Form
    
    For Each frm In Forms
        Unload frm
    Next frm
    
End Sub

Public Sub Restore_Status()
'//get system restore availability

    bOSVersion = Restore_Available
    If Not bOSVersion Then
        With frmRecovery
            .optRestore(1).Value = True
            .optRestore(0).Enabled = False
            .lblStatus(0).Caption = "System Restore: Not Available"
            .lblStatus(1).Caption = "Last Restore: Not Available"
        End With
    Else
        With frmRecovery
            '.lblStatus(0).Caption = "System Restore: Available"
            '.lblStatus(1).Caption = "Last Restore: Pending"
        End With
    End If
    
End Sub

Public Sub Get_Scanner_Options()
'//set persistent options

Dim i As Integer

    With frmSearch
        '//scanner controls
        For i = 0 To 7
            .chkOptions(i).Value = GetSetting(App.EXEName, "MRUPro", "chkopt" & i, "1")
        Next i
        .chkAll.Value = GetSetting(App.EXEName, "MRUPro", "chkall", "1")
    End With
    '//restore type status
    bRestoreType = GetSetting(App.EXEName, "MRUPro", "optres0", "False")
        
End Sub

Public Sub Set_Scanner_Options()
'//fetch persistent options

Dim i As Integer

    With frmSearch
        '//scanner controls
        For i = 0 To 7
            SaveSetting App.EXEName, "MRUPro", "chkopt" & i, .chkOptions(i).Value
        Next i
        SaveSetting App.EXEName, "MRUPro", "chkall", .chkAll.Value
    End With
    
End Sub

Public Sub Set_Recovery_Options()
'//set persistent options

    With frmRecovery
        '//recovery controls
        SaveSetting App.EXEName, "MRUPro", "optres0", .optRestore(0).Value
        SaveSetting App.EXEName, "MRUPro", "optres1", .optRestore(1).Value
        SaveSetting App.EXEName, "MRUPro", "lblsta0", .lblStatus(0).Caption
        SaveSetting App.EXEName, "MRUPro", "lblsta1", .lblStatus(1).Caption
        SaveSetting App.EXEName, "MRUPro", "lblsta2", .lblStatus(2).Caption
    End With
    
End Sub

Public Sub Get_Recovery_Options()
'//fetch persistent options

    With frmRecovery
        '//recovery controls
        .optRestore(0).Value = GetSetting(App.EXEName, "MRUPro", "optres0", "False")
        .optRestore(1).Value = GetSetting(App.EXEName, "MRUPro", "optres1", "True")
        .lblStatus(0).Caption = GetSetting(App.EXEName, "MRUPro", "lblsta0", "System Restore: Not Available")
        .lblStatus(1).Caption = GetSetting(App.EXEName, "MRUPro", "lblsta1", "Last Restore: Not Available")
        .lblStatus(2).Caption = GetSetting(App.EXEName, "MRUPro", "lblsta2", "Last Snapshot: Pending")
    End With
    
End Sub
