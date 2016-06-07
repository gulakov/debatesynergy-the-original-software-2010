Attribute VB_Name = "d8paperless"
Sub D8_SpeechResize()
    w_flowresize
End Sub

Sub D8_FolderAutoOpen()
'Runs in the background to automatically open all documents in the speech folder.

    On Error Resume Next
    Dim X, d, sPath, fso As New FileSystemObject, IsOpen, AllFiles, ListenOn As Boolean, FPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    sPath = D8x("SPath")
    
    
    If MsgBox("This feature will start a listener that automatically opens all documents in " & _
        "the root of your Speech folder (" & sPath & "), but ignoring anything in subfolders. " & _
        "Press Escape to stop the listener. It is suggested that the Speech folder be set to " & _
        "a DropBox folder (getdropbox.com).", _
            vbOKCancel + vbQuestion + vbDefaultButton2, "AutoOpen Folder") <> vbOK Then Exit Sub
            
    Do
        DoEvents
        Application.Caption = "Debate Synergy [LISTENER ON]"
        Set AllFiles = fso.GetFolder(sPath).Files
        
        For Each X In AllFiles
                FPath = sPath & X.Name
                
                
                'is the file open?
                IsOpen = False
                For Each d In Documents
                    If d.FullName = FPath Then IsOpen = True
                Next d
                
                'if not open, open it
                If Not IsOpen And Left(X.Name, 1) <> "~" And _
                    (Right(FPath, 3) = "doc" Or Right(FPath, 4) = "docx" Or Right(FPath, 3) = "rtf") _
                    Then Documents.Open FPath
        Next X
    Loop Until GetKeyState(&H1B) < 0
    Application.Caption = "Debate Synergy"
    MsgBox "Stopped listening to the Speech folder.", vbInformation
    
    Set fso = Nothing
End Sub


Sub D8_Speech()
' Sends selection to Speech.
' If no selection, toggles reading view.
   
    With Selection
        If .Type < 2 Then 'if no selection
            If (.Paragraphs(1).Range.Start = .Start Or _
                .Paragraphs(1).Range.Start + 1 = .Start) And _
                .Range.ParagraphFormat.OutlineLevel = 1 Then 'if on block header
                    Call D8_BlockSelect
                    Call D8_SpeechSend
            Else
                With ActiveWindow
                    If .View.FullScreen Then ' if in fullscreen mode, insert marker
                        .View.FullScreen = False
                        .DocumentMap = True
                        .View.Zoom = 100
                    Else 'if not in reading mode, go to reading mode
                        Call D8_SpeechRead
                    End If
                End With
            End If
        Else
              Call D8_SpeechSend
        End If
        
    End With
    
    
    
    
Handler:
End Sub
Public Function D8_SpeechRead()
    On Error Resume Next
           
    ActiveDocument.Styles("Document Map").Font.Size = 16
    With ActiveWindow
        .View.Type = wdWebView
        .View.FullScreen = True
        .DisplayVerticalScrollBar = True
        .DisplayHorizontalScrollBar = False
        .DisplayRulers = False
        .DocumentMap = True
        .View.Zoom = 200
        
    End With
End Function
Public Function D8_SpeechSend()
'this macro is inspired by Aaron Hardy and the Whitman template

    On Error Resume Next
    
    Dim d, Sp As Document
    
    If Selection.Type < 2 Then Exit Function
    
    For Each d In Documents
        If InStr(LCase(d.Name), "speech") Then
            Set Sp = d
            Exit For
        End If
    Next d
    If Sp Is Nothing Then Exit Function
    
    Selection.Copy

    With Sp.ActiveWindow.Selection
        If Len(.Paragraphs(1).Range) > 1 Then .TypeParagraph
        .Paste
        If Len(.Paragraphs(1).Range) > 1 Then .TypeParagraph
        .EndKey
    End With
    
    Set Sp = Nothing
    
End Function
Public Sub D8_SpeechMarker()
    
    If ActiveWindow.View.ReadingLayout Then ActiveWindow.View.ReadingLayoutAllowEditing = True
    With Selection
        .Collapse 0
        If .Words(1).End <> .End Then .MoveRight wdWord
        .Font.Color = wdColorRed
        .Font.Size = .Font.Size + 5
        .TypeText ChrW(8362) & " stopped here at " & FormatDateTime(Time, 4) & " " & ChrW(8362) & " "
    End With
End Sub
Public Function D8_SpeechNew(iName As String)
    
    Dim cTime, SpName, i
    
    If Hour(Now) > 12 Then cTime = Hour(Now) - 12 & "PM"
    If Hour(Now) <= 12 Then cTime = Hour(Now) & "AM"

    SpName = D8x("SPath") & "Speech " & iName & " " & _
        Month(Now) & "-" & Day(Now) & " " & cTime
        
    'check filename uniqueness
    If Dir(SpName & ".doc") > "" Then
        For i = 2 To 100
            If Dir(SpName & " " & i & ".doc") = "" Then Exit For
        Next i
        SpName = SpName & " " & i
    End If
    
    Documents.Add.SaveAs SpName & ".doc", 0
    ActiveWindow.View.Type = wdWebView
    ActiveWindow.View.Zoom = 100
    
End Function

Public Sub D8_SaveUSB()
'Saves the active document to USB drive and to the speech save folder.
    
    Dim Buff As String, drList() As String, SaveMess As String, ActPath, u
    
    ActPath = ActiveDocument.Path
    SaveMess = "Saved document to: " & vbCr

    
    'get USB list
    Buff = Space(105)
    GetLogicalDriveStringsA 105, Buff
    drList = Split(Buff, vbNullChar)
    For u = 0 To UBound(drList)
        If GetDriveTypeA(drList(u)) = 2 And drList(u) <> "A:\" Then
            ActiveDocument.SaveAs drList(u) & ActiveDocument.Name
            SaveMess = SaveMess & drList(u) & vbCr
        End If
    Next
            
   'save locally
    If ActPath > "" Then
        SaveMess = SaveMess & ActPath & vbCr

        ActiveDocument.SaveAs ActPath & "\" & ActiveDocument.Name
    Else
        SaveMess = SaveMess & D8x("SpeechFolder") & vbCr

        ActiveDocument.SaveAs D8x("SpeechFolder") & "\" & ActiveDocument.Name
    End If
    
    MsgBox SaveMess, vbInformation
End Sub



Sub FlowReceive(SideColor As Long)
'Receives flow segments from Excel Debate Synergy flow template.
    
    On Error GoTo Handler
    
    Dim PsChk, Ps, p
    Set PsChk = New DataObject
    PsChk.GetFromClipboard
    Ps = PsChk.GetText(1)
    
    Selection.Collapse 0
    Selection = Ps
    
    WordBasic.FormatStyle Name:="Black"
    
    With ActiveDocument
        .Styles("Black").BaseStyle = "Normal"
        .Styles("Black").Font.Shading.Texture = wdTextureNone
        .Styles("Normal").NextParagraphStyle = "Black"
        Selection.Style = .Styles("Normal")
    End With
    
    With Selection
        For Each p In .Paragraphs
            If Len(Trim(p.Range)) > 2 Then
                p.Range.Font.Shading.ForegroundPatternColor = SideColor
                p.Range.Font.Shading.Texture = 1000
            End If
        Next
    End With
    
    Set PsChk = Nothing
Handler:
End Sub
