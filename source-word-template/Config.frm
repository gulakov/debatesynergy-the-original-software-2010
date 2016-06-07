VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Config 
   Caption         =   "Debate Synergy Options"
   ClientHeight    =   5650
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   3880
   OleObjectBlob   =   "Config.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim HeaderCh As Boolean, BrowseCh As Boolean



Private Sub keyreset_Click()
    If MsgBox("Are you sure you want to reset all of your keyboard shortcuts?", vbYesNo + vbQuestion, "Reset Key Shortcuts") <> vbYes Then Exit Sub

    CustomizationContext = NormalTemplate
    With KeyBindings
        .ClearAll
        .Add 2, "D8_BlockDown", BuildKeyCode(wdKeyControl, wdKeyAlt, 40)
        .Add 2, "D8_BlockEnd", BuildKeyCode(wdKeyControl, wdKeyAlt, 39)
        .Add 2, "D8_BlockUp", BuildKeyCode(wdKeyControl, wdKeyAlt, 38)
        .Add 2, "D8_BlockStart", BuildKeyCode(wdKeyControl, wdKeyAlt, 37)
        .Add 2, "D8_CiteMagic", BuildKeyCode(wdKeyControl, wdKeyT)
        .Add 2, "D8_CiteMagic", BuildKeyCode(wdKeyAlt, wdKeyT)
        .Add 2, "D8_FixCaps", BuildKeyCode(wdKeyControl, wdKeyK)
        .Add 2, "D8_FixCiteRequest", BuildKeyCode(wdKeyControl, wdKeyQ)
        .Add 2, "D8_FixPageContinued", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyReturn)
        .Add 2, "D8_FormatNotHeading", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF3)
        .Add 2, "D8_FormatHeading", BuildKeyCode(wdKeyF3)
        .Add 2, "D8_FormatHat", BuildKeyCode(wdKeyF3, wdKeyControl)
        .Add 2, "D8_FormatToggle", BuildKeyCode(wdKeyF1)
        .Add 2, "D8_FormatSmallAll", BuildKeyCode(wdKeyF1, wdKeyControl)
        .Add 2, "D8_FormatSmallAllMore", BuildKeyCode(wdKeyF1, wdKeyControl, wdKeyAlt)
        .Add 2, "D8_FormatNormal", BuildKeyCode(wdKeyF2)
        .Add 2, "D8_FormatSimilar", BuildKeyCode(wdKeyF2, wdKeyControl)
        .Add 2, "D8_FormatHighlight", BuildKeyCode(wdKeyF4)
        .Add 2, "D8_FormatBox", BuildKeyCode(wdKeyF4, wdKeyControl)
        .Add 2, "D8_PageHeader", BuildKeyCode(wdKeyF11)
        .Add 2, "D8_PasteText", BuildKeyCode(wdKeyG, wdKeyControl)
        .Add 2, "D8_PasteText", BuildKeyCode(wdKeyG, wdKeyAlt)
        .Add 2, "D8_PasteReturns", BuildKeyCode(wdKeyG, wdKeyControl, wdKeyAlt)
        .Add 2, "D8_RemoveReturns", BuildKeyCode(wdKeyControl, wdKeyR)
        .Add 2, "D8_RemoveReturns", BuildKeyCode(wdKeyAlt, wdKeyR)
        .Add 2, "D8_ScrollDown", BuildKeyCode(wdKeyPageDown)
        .Add 2, "D8_SaveUSB", BuildKeyCode(wdKeyS, wdKeyControl, wdKeyShift)
        
        .Add 2, "D8_Speech", BuildKeyCode(wdKeyBackSingleQuote)
        .Add 2, "D8_SpeechMarker", BuildKeyCode(wdKeyControl, wdKeyBackSingleQuote)
        .Add 2, "D8_SpeechResize", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyBackSingleQuote)
        
        
      '  .Add 2, "D8_Timer", BuildKeyCode(wdKeyF10)
        .Add 2, "D8_TOC", BuildKeyCode(wdKeyF12)
        .Add 2, "D8_WindowCycle", BuildKeyCode(wdKeyControl, wdKeyTab)
        
        .Add 2, "D8_ZoomFull", BuildKeyCode(wdKeyControl, wdKey0)
        .Add 2, "D8_WarrantAdd", BuildKeyCode(wdKeyF9)
        .Add 2, "D8_WarrantToggle", BuildKeyCode(wdKeyControl, wdKeyF9)
    End With
End Sub




Sub UserForm_Initialize()
    HeaderCh = False
    
    FirstRun
    
    'main
    Dim SN: SN = Application.UserName
    UName.Text = Left(SN, InStr(SN, ",") - 1)
    teamname.Text = LTrim(Mid(SN, InStr(SN, ",") + 1))
    header.Value = D8x("Header")
    pagecount.Value = D8x("PageCount")
    toolbar.Value = D8x("Toolbar")
    lastedit.Value = D8x("LastEdit")
    recover.Text = Options.SaveInterval
    startview.Value = D8x("startview")
    
    'folders
    exploc.Text = D8x("VTub")
    sploc.Text = D8x("SpeechFolder")
    everypath.Text = D8x("EveryPath")
    everyprog.Text = D8x("EveryProg")

    'appearance
    cite.Text = D8x("Cite")
    cwords.Text = D8x("CiteWords")
    small.Text = D8x("Small")
    continues.Text = D8x("Continues")
    toc.Value = D8x("RemoveTOC")
    pastecol.Value = D8x("Paste")
    
    'show
    x1.Value = D8x("x1")
    x2.Value = D8x("x2")
    x3.Value = D8x("x3")
    x4.Value = D8x("x4")
    x5.Value = D8x("x5")
    x6.Value = D8x("x6")
    x7.Value = D8x("x7")
    x8.Value = D8x("x8")
    x9.Value = D8x("x9")
    x10.Value = D8x("x10")
    x11.Value = D8x("x11")
    x12.Value = D8x("x12")
    x13.Value = D8x("x13")
    x14.Value = D8x("x14")
    x15.Value = D8x("x15")
    x16.Value = D8x("x16")
    x17.Value = D8x("x17")
    x18.Value = D8x("x18")
    x19.Value = D8x("x19")
    x20.Value = D8x("x20")
    x21.Value = D8x("x21")
    x22.Value = D8x("x22")
    x23.Value = D8x("x23")
    x24.Value = D8x("x24")
    x25.Value = D8x("x25")
    
    Me.show
    Do
        DoEvents
        If GetKeyState(&H1B) < 0 Then End
    Loop
    
End Sub

Private Sub ok_click()

    'make sure folders exist
    Dim fso As New FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(everypath.Text) Then MkDir everypath.Text
    If Not fso.FolderExists(exploc.Text) Then MkDir exploc.Text
    If Not fso.FolderExists(sploc.Text) Then MkDir sploc.Text
    
    'main
    Application.UserName = UName.Text & ", " & teamname.Text
    D8s "Header", header.Value
    D8s "PageCount", pagecount.Value
    D8s "Toolbar", toolbar.Value
    D8s "LastEdit", lastedit.Value
    Options.SaveInterval = val(recover.Text)
    D8s "startview", startview.Value
    
    'folders
    D8s "VTub", exploc.Text
    D8s "SpeechFolder", sploc.Text
    D8s "EveryPath", everypath.Text
    D8s "EveryProg", everyprog.Text

    'appearance
    D8s "Cite", cite.Text
    D8s "CiteWords", cwords.Text
    D8s "Small", small.Text
    D8s "Continues", continues.Text
    D8s "RemoveTOC", toc.Value
    D8s "Paste", pastecol.Value
    
    'show
    D8s "x1", x1.Value
    D8s "x2", x2.Value
    D8s "x3", x3.Value
    D8s "x4", x4.Value
    D8s "x5", x5.Value
    D8s "x6", x6.Value
    D8s "x7", x7.Value
    D8s "x8", x8.Value
    D8s "x9", x9.Value
    D8s "x10", x10.Value
    D8s "x11", x11.Value
    D8s "x12", x12.Value
    D8s "x13", x13.Value
    D8s "x14", x14.Value
    D8s "x15", x15.Value
    D8s "x16", x16.Value
    D8s "x17", x17.Value
    D8s "x18", x18.Value
    D8s "x19", x19.Value
    D8s "x20", x20.Value
    D8s "x21", x21.Value
    D8s "x22", x22.Value
    D8s "x23", x23.Value
    D8s "x24", x24.Value
    D8s "x25", x25.Value
    
    
    
    If HeaderCh Then
        NormalTemplate.OpenAsDocument
        If header.Value Then
            D8_PageHeader
        Else
            With ActiveWindow.ActivePane.View
                .Type = wdPrintView
                .SeekView = wdSeekCurrentPageHeader
                Selection.WholeStory
                Selection.Delete
                .SeekView = wdSeekMainDocument
            End With
        End If
        Documents(NormalTemplate.Name).Save
        Documents(NormalTemplate.Name).Close 0
    End If
    
    Fresh
    End
End Sub

Private Sub resetall_click()

    If MsgBox("Are you sure you want to reset all of your settings?", vbYesNo + vbQuestion, "Reset Settings") <> vbYes Then Exit Sub
    
    
    D8s "SpeechFolder", Replace(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\", "\\", "\")
    D8s "everyProg", Replace("C:\Program Files\Everything\Everything.exe", "\\", "\")
    D8s "everyPath", Replace(Options.DefaultFilePath(wdDocumentsPath) & "\", "\\", "\")

    Dim VTub As String, fso As New FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    VTub = Options.DefaultFilePath(wdDocumentsPath) & "\Virtual Tub\"
    If Not fso.FolderExists(VTub) Then MkDir VTub
    D8s "VTub", Replace(VTub, "\\", "\")
    Set fso = Nothing
        
    D8s "Cite", "AuthorLast Year – Quals (AuthorFirst, Date, Title, URL, Initials)"
    D8s "CiteWords", 5
    
    D8s "Small", 8
    D8s "Continues", "[CONTINUED]"
    D8s "RemoveTOC", True
    D8s "Header", False
    D8s "PageCount", True
    D8s "Toolbar", False
    D8s "Paste", False
    D8s "LastEdit", True
    
    D8s "x1", True
    D8s "x2", True
    D8s "x3", False
    D8s "x4", False
    D8s "x5", True
    D8s "x6", False
    D8s "x7", False
    D8s "x8", True
    D8s "x9", True
    D8s "x10", True
    D8s "x11", True
    D8s "x12", True
    D8s "x13", True
    D8s "x14", True
    D8s "x15", True
    D8s "x16", True
    D8s "x17", True
    D8s "x18", True
    D8s "x19", True
    D8s "x20", True
    D8s "x21", True
    D8s "x22", True
    D8s "x23", True
    D8s "x24", True
    D8s "x25", False
    

    
    UserForm_Initialize
    
    
End Sub

'other keys
Private Sub keys_Click()
    With Dialogs(432)
        .Category = 2
        .show
    End With
End Sub

Private Sub header_Change()
    HeaderCh = True
End Sub
Private Sub cancel_Click()
    End
End Sub


'change paths

Private Sub browseeprog_Click()
    With Application.FileDialog(3)
        .AllowMultiSelect = False
        .InitialFileName = "C:\Program Files\"
        .InitialView = 2
        .Filters.Clear
        .Filters.Add "Programs", "*.exe"
        If .show Then If InStr(.SelectedItems(1), "\Everything.exe") Then _
            everyprog.Text = .SelectedItems(1)
    End With
        
End Sub

Private Sub browseepath_Click()
    Dim locPick
    locPick = Browse
    If locPick > "" Then everypath.Text = locPick
End Sub
Private Sub browseexp_Click()
    Dim locPick
    locPick = Browse
    If locPick > "" Then exploc.Text = locPick
End Sub
Private Sub browsesp_Click()
    Dim locPick
    locPick = Browse
    If locPick > "" Then sploc.Text = locPick
End Sub
Function Browse() As String
    
    Dim sh, dPick, sPath

    Set sh = New Shell32.Shell
    Set dPick = sh.BrowseForFolder(0, "Select a new folder.", 0, "")
        
    If dPick Is Nothing Then Exit Function
    
    sPath = dPick.Self.Path
    If InStr(sPath, "{") Then Exit Function
    
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    BrowseCh = True
    Browse = sPath
    
    Set sh = Nothing
    Set dPick = Nothing
End Function

'block specific inputs
Private Sub cwords_Change()
    Dim X
    For X = 1 To 255
        If (X < 48 Or X > 57) Then cwords.Text = Replace(cwords.Text, Chr(X), "")
    Next X
End Sub
Private Sub small_Change()
    Dim X
    For X = 1 To 255
        If (X < 48 Or X > 57) Then small.Text = Replace(small.Text, Chr(X), "")
    Next X
End Sub
Private Sub recover_Change()
    Dim X
    For X = 1 To 255
        If (X < 48 Or X > 57) Then recover.Text = Replace(recover.Text, Chr(X), "")
    Next X
End Sub
Private Sub cite_change()
    cite.Text = Replace(cite.Text, "'", "‘")
    cite.Text = Replace(cite.Text, vbCrLf, "")
    
End Sub

'hand cursors
Sub browsetub_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SetCursor 65567: End Sub
Sub browseexp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SetCursor 65567: End Sub
Sub browsesp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SetCursor 65567: End Sub
Sub ok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SetCursor 65567: End Sub
Sub resetall_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SetCursor 65567: End Sub
Sub cancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SetCursor 65567: End Sub
Sub keys_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SetCursor 65567: End Sub
Sub keyreset_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): SetCursor 65567: End Sub

Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    End
End Sub
