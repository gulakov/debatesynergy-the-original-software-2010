VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Config 
   Caption         =   "Debate Synergy Flow Options"
   ClientHeight    =   5710
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   3910
   OleObjectBlob   =   "Config.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wd As New Word.Application
  

Sub UserForm_initialize()
    floc = D8x("FPath")
    skiprows = D8x("SkipRows")
    abc = D8x("ABC")
    voters = D8x("Voters")
    authors = D8x("Authors")
    flowtitle = D8x("FlowTitle")
End Sub

Private Function D8x(Setting As String) As String
    On Error Resume Next
    Dim Data
    Data = wd.System.PrivateProfileString(Application.TemplatesPath & "\D8.ini", "Flow", Setting)
    Select Case Data
        Case "True": D8x = True
        Case "False": D8x = False
        Case Else: D8x = Data
    End Select
End Function
Private Sub D8s(Setting As String, Data As Variant)
    On Error Resume Next
    wd.System.PrivateProfileString(Application.TemplatesPath & "\D8.ini", "Flow", Setting) = Data
End Sub

Private Sub ok_click()
    
    D8s "FPath", floc
    D8s "SkipRows", skiprows
    D8s "ABC", abc
    D8s "Voters", voters
    D8s "Authors", authors
    D8s "FlowTitle", flowtitle
    
    Set wd = Nothing
    ActiveWorkbook.Activate
    End
End Sub

Private Sub reset_click()

    If MsgBox("Are you sure you want to reset all of your settings?", vbYesNo + vbQuestion, "Reset Settings") <> vbYes Then Exit Sub
    
    Set wd = New Word.Application
    
    D8s "FPath", Replace(Application.DefaultFilePath & "\Flows\", "\\", "\")
    D8s "SkipRows", True
    D8s "ABC", True
    D8s "Voters", True
    D8s "Authors", True
    D8s "FlowTitle", True
    
    Set wd = Nothing
    ActiveWorkbook.Activate
    
    
    UserForm_initialize
End Sub

'other keys


Private Sub modshort_Click()
    On Error Resume Next
    SPath = Application.TemplatesPath & "\FlowShorthand.txt"
    Dir (SPath)
        If Err Then SPath = "C:\FlowShorthand.txt"
   
    If Dir(SPath) = "" Then Range("A3").ClearContents
    ShellOpen hWndAccessApp, vbNullString, SPath, vbNullString, vbNullString, 1
End Sub

Private Sub shortreset_Click()
    If MsgBox("Are you sure you want to reset all of your shorthand values?", vbYesNo + vbQuestion, "Reset Settings") <> vbYes Then Exit Sub
    
    On Error Resume Next
    SPath = Application.TemplatesPath & "\FlowShorthand.txt"
    Dir (SPath)
        If Err Then SPath = "C:\FlowShorthand.txt"
   
   Kill SPath
   Range("A3").ClearContents
End Sub
Private Sub cancel_Click()
    End
End Sub


'change paths
Private Sub browsef_Click()
    Dim sh, dPick, SPath

    Set sh = New Shell32.Shell
    Set dPick = sh.BrowseForFolder(0, "Select a new folder.", 0, "")
        
    If dPick Is Nothing Then Exit Sub
    
    SPath = dPick.Self.Path
    If InStr(SPath, "{") Then Exit Sub
    
    If Right(SPath, 1) <> "\" Then SPath = SPath & "\"
    
    If SPath > "" Then floc = SPath
End Sub


'hand cursors
Sub cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single): SetCursor (LoadCurson(0&, 32649&)): End Sub
Sub ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single): SetCursor (LoadCurson(0&, 32649&)): End Sub
Sub reset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single): SetCursor (LoadCurson(0&, 32649&)): End Sub
Sub browsef_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single): SetCursor (LoadCurson(0&, 32649&)): End Sub
Sub modshort_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single): SetCursor (LoadCurson(0&, 32649&)): End Sub
Sub shortreset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single): SetCursor (LoadCurson(0&, 32649&)): End Sub





