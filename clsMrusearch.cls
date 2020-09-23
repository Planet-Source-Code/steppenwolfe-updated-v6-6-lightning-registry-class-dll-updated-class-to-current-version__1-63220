VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMrusearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//could - and possibly should - make reg calls inline with this class
'//rather then passing collections between classes.. But I wanted
'//to leave you with a reusable class, and this way, can also call
'//reghandler from registry cleaner routines
'//also note use of syntax of 'With New' rather then dim and set
'//to invoke a class, ex in MRU_Recurse_Keys - better method I think..

Private cKeyList        As New Collection
Private sExplList()     As String
Private bExpList        As Boolean
Private cSubKey         As Collection
Private cMatch          As New Collection
Private lProgress       As Long
Private bBuild0         As Boolean
Private bBuild1         As Boolean
Private bBuild2         As Boolean
Private bBuild3         As Boolean
Private bBuild4         As Boolean
Private bBuild5         As Boolean
Private bBuild6         As Boolean
Private bBuild7         As Boolean
Private bBuild8         As Boolean

Dim cReg As New clsRegHandler

'//progress and statistics
Public Event Process(ByVal percent As Long)
Public Event Maxcount(ByVal maxnum As Long)
Public Event Keycount(ByVal keynum As Long)
Public Event Listcount(ByVal listnum As Long)
Public Event Valcount(ByVal valnum As Long)
Public Event Labelchange(ByVal sCaption As String)

Private Sub Class_Initialize()
    '//initialize collections
    Set cKeyList = New Collection
    Set cMatch = New Collection
    Set cReg = New clsRegHandler
End Sub

Private Sub Class_Terminate()
    '//terminate collections
    Set cKeyList = Nothing
    Set cMatch = Nothing
    Set cReg = Nothing
End Sub

Public Function Mru_Control(Optional ByVal lHKey As Long, _
                            Optional ByVal sSubKey As String) As Collection

Dim Item        As Variant
Dim cFilter     As New Collection
Dim cTemp       As New Collection
Dim lVCount     As Long
Dim lIncr       As Long

    '//check for all options
    If bBuild8 Then
        Set Mru_Control = Predictive_MRU_Search(lHKey, sSubKey)
        RaiseEvent Valcount(Mru_Control.Count)
        Exit Function
    End If
    
    '//filter results to user selected
    Set cFilter = New Collection
    Set cTemp = New Collection
    '//transfer our collection for event calls
    Set cTemp = Predictive_MRU_Search(lHKey, sSubKey)
    For Each Item In cTemp
        RaiseEvent Labelchange("Filtering MRU Entries..")
        RaiseEvent Maxcount(cTemp.Count)
        If Build_Control(Item) Then
            lVCount = lVCount + 1
            RaiseEvent Valcount(lVCount)
            cFilter.Add Item
        End If
        lIncr = lIncr + 1
        RaiseEvent Process(lIncr)
        DoEvents
    Next Item
    
    Set cTemp = Nothing
    Set Mru_Control = cFilter
    Set cFilter = Nothing
    
End Function

Private Function Build_Control(ByVal vMatch As Variant) As Boolean
'//filter function

    '//office mru
    If bBuild0 Then
        If InStr(vMatch, "\Office\") > 0 Then
            Build_Control = True
            Exit Function
        End If
    End If
    '//recent documents
    If bBuild1 Then
        If InStr(vMatch, "\RecentDocs") > 0 Then
            Build_Control = True
            Exit Function
        End If
    End If
    '//internet explorer
    If bBuild2 Then
        If InStr(vMatch, "\Internet Explorer\") > 0 Then
            Build_Control = True
            Exit Function
        End If
    End If
    '//adobe
    If bBuild3 Then
        If InStr(vMatch, "\Adobe\") > 0 Then
            Build_Control = True
            Exit Function
        End If
    End If
    '//macromedia
    If bBuild4 Then
        If InStr(vMatch, "\Macromedia\") > 0 Then
            Build_Control = True
            Exit Function
        End If
    End If
    '//wordpad
    If bBuild5 Then
        If InStr(vMatch, "\Wordpad\") > 0 Then
            Build_Control = True
            Exit Function
        End If
    End If
    '//mediaplayer
    If bBuild6 Then
        If InStr(vMatch, "\MediaPlayer") > 0 Then
            Build_Control = True
            Exit Function
        End If
    End If
    '//run mru
    If bBuild7 Then
        If InStr(vMatch, "\RunMRU") > 0 Then
            Build_Control = True
            Exit Function
        End If
    End If

End Function

Friend Sub Pass_UDT(ByRef tpeChoice As tChoice)

    '//pass user choices into class
    '//with friend sub
    With tpeChoice
        If .tChoice0 Then bBuild0 = True
        If .tChoice1 Then bBuild1 = True
        If .tChoice2 Then bBuild2 = True
        If .tChoice3 Then bBuild3 = True
        If .tChoice4 Then bBuild4 = True
        If .tChoice5 Then bBuild5 = True
        If .tChoice6 Then bBuild6 = True
        If .tChoice7 Then bBuild7 = True
        If .tChoice8 Then bBuild8 = True
    End With

End Sub

Private Function Predictive_MRU_Search(ByVal lHKey As Long, _
                                      ByVal sSubKey As String) As Collection
'//The most common value for mru's are File0, url0, a-z, 0-99, 000-999
'//that with a second scoring of common path references mru\history\list\file\recent\last
'//and it is most certainly an mru entry

Dim cVal        As Variant
Dim cSearch     As Variant
Dim cMatchVal   As Variant
Dim sMWord      As String
Dim x           As Long
Dim lCcount     As Long
Dim cKey        As Variant
Dim lIncr       As Long

On Error Resume Next

    Set cMatch = New Collection
    '//get key candidates with reg key
    '//enumeration recursion loop
    RaiseEvent Labelchange("Scanning Registry Keys..")
    MRU_Recurse_Keys lHKey, sSubKey
    RaiseEvent Labelchange("Searching Registry Values..")
    '//Search each candidate keys values
    '//for potential mru entries
    With cReg
    For Each cSearch In cKeyList
        RaiseEvent Maxcount(cKeyList.Count)
        '//first lookup filters test for mru word,
        '//on a match, lists all possible words of
        '//that type under the match key
        Select Case True
            '//first test if desired value type is present
            Case .Search_Value(lHKey, cSearch, "file0")
                '//match positive, is valid mru key
                '//so increment list number value
                RaiseEvent Listcount(1)
                '//get values
                For Each cMatchVal In .List_Values(lHKey, cSearch)
                    '//vals containing 'file'
                    If InStr(LCase(cMatchVal), "file") > 0 Then
                        '//last filter check on correct length
                        If Not Len(cMatchVal) > 5 Then
                            '//add to temporary collection
                            cMatch.Add cSearch & Chr(31) & cMatchVal
                        End If
                    End If
                Next cMatchVal
                
            Case .Search_Value(lHKey, cSearch, "url1")
                RaiseEvent Listcount(1)
                For Each cMatchVal In .List_Values(lHKey, cSearch)
                    '//vals containing 'url'
                    If InStr(LCase(cMatchVal), "url") > 0 Then
                        If Not Len(cMatchVal) > 5 Then
                            cMatch.Add cSearch & Chr(31) & cMatchVal
                        End If
                    End If
                Next cMatchVal
                
            Case .Search_Value(lHKey, cSearch, "a")
                RaiseEvent Listcount(1)
                For Each cMatchVal In .List_Values(lHKey, cSearch)
                    '//lowercase a - z
                    For x = 97 To 122
                        If Chr(x) = LCase(cMatchVal) Then
                            cMatch.Add cSearch & Chr(31) & cMatchVal
                            Exit For
                        End If
                    Next x
                Next cMatchVal
                
            Case .Search_Value(lHKey, cSearch, "0")
                RaiseEvent Listcount(1)
                For Each cMatchVal In .List_Values(lHKey, cSearch)
                    '0 to 99
                    For x = 0 To 99
                        If cMatchVal = x Then
                            cMatch.Add cSearch & Chr(31) & cMatchVal
                            Exit For
                        End If
                    Next x
                Next cMatchVal
                
            Case .Search_Value(lHKey, cSearch, "000")
                RaiseEvent Listcount(1)
                For Each cMatchVal In .List_Values(lHKey, cSearch)
                    '000 To 999
                    For x = 0 To 999
                        sMWord = Format$(x, "000")
                        If cMatchVal = sMWord Then
                            cMatch.Add cSearch & Chr(31) & cMatchVal
                            Exit For
                        End If
                    Next x
                Next cMatchVal
                
            Case Else
            '//sanity check: if no match is found get
            '//next level of key and add to collection -
            '//as many mru entries are form ex. mru\5603 - mru values
                For Each cKey In .List_Keys(lHKey, cSearch)
                    cKeyList.Add cSearch & Chr(92) & cKey
                Next cKey
        End Select
        
        lIncr = lIncr + 1
        RaiseEvent Process(lIncr)
        DoEvents
    Next
    End With
    Set Predictive_MRU_Search = cMatch
    
On Error GoTo 0

End Function

Private Sub MRU_Recurse_Keys(ByVal lHKey As Long, _
                             ByVal sSubKey As String)

Dim cKey        As Variant
Dim x           As Integer

On Error Resume Next

    '//start by recursing software keys
    '//and adding them to a collection
    '//this is the fastest recursion routine
    '//I have ever seen!! 12,000 subkeys
    '//enumerated in .3 seconds on my machine!!!
    For Each cKey In cReg.List_Keys(lHKey, sSubKey)
        If Not IsEmpty(cKey) Then
            '//first filter to likely subkeys
            '//and add to our collection
            Select Case x <> 0
                '//test key for likely match word
                Case x = InStr(LCase$(cKey), "mru")
                    '//positives added to collection
                    cKeyList.Add sSubKey & Chr(92) & cKey
                Case x = InStr(LCase$(cKey), "history")
                    cKeyList.Add sSubKey & Chr(92) & cKey
                Case x = InStr(LCase$(cKey), "list")
                    cKeyList.Add sSubKey & Chr(92) & cKey
                Case x = InStr(LCase$(cKey), "file")
                    cKeyList.Add sSubKey & Chr(92) & cKey
                Case x = InStr(LCase$(cKey), "recent")
                    cKeyList.Add sSubKey & Chr(92) & cKey
                Case x = InStr(LCase$(cKey), "last")
                    cKeyList.Add sSubKey & Chr(92) & cKey
                Case x = InStr(LCase$(cKey), "url")
                    cKeyList.Add sSubKey & Chr(92) & cKey
            End Select
            RaiseEvent Keycount(1)
            '//recurse each key
            MRU_Recurse_Keys lHKey, sSubKey & Chr(92) & cKey
        End If
        DoEvents
    Next

On Error GoTo 0

End Sub
