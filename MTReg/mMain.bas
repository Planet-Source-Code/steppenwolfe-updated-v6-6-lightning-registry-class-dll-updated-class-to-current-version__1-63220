Attribute VB_Name = "mMain"
Option Explicit
'//original loader module as written by Big Steve M.
'//runnable tlb is also by Steve, comments left intact

'//FileName:    mStart.bas
'//Author:      SP McMahon
'//Date:        2 February 2000

'//Collection of Runnable items to start:
Private m_colRunnables   As Collection
'//The ID of our API Timer:
Private m_lTimerID       As Long



Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal nIDEvent As Long) As Long
                                                 
Private Declare Function CoLockObjectExternal Lib "ole32" (ByVal pUnk As IUnknown, _
                                                           ByVal fLock As Long, _
                                                           ByVal fLastUnlockReleases As Long) As Long
                                                           
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long


Private Sub TimerProc(ByVal lHwnd As Long, _
                      ByVal lMsg As Long, _
                      ByVal lTimerID As Long, _
                      ByVal lTime As Long)


Dim this As Runnable

    '//Enumerate through the collection, firing the
    '//Runnable_Start method for each item in it and
    '//releasing our extra lock on the object:
    With m_colRunnables
        Do While .Count > 0
            Set this = .Item(1)
            .Remove 1
            this.Start
            '//Ask the system to release its lock on the object
            CoLockObjectExternal this, 0, 1
        Loop
    End With
    '//Remove the timer:
    KillTimer 0, lTimerID
    m_lTimerID = 0

End Sub

Public Sub Start(this As Runnable)

'//Ask the system to lock the object so that
'//it will still perform its work even if it
'//is released

    CoLockObjectExternal this, 1, 1
    '//Add this to runnables:
    If m_colRunnables Is Nothing Then
        Set m_colRunnables = New Collection
    End If
    m_colRunnables.Add this
    '//Create a timer to start running the object:
    If Not m_lTimerID Then
        m_lTimerID = SetTimer(0, 0, 1, AddressOf TimerProc)
    End If

End Sub


