Attribute VB_Name = "D8Format"
'I don't comment much - it was hard to write so it should be hard to read!

Sub D8_FormatSimilar()
'Selects all text in entire document that is similar to current text.

    On Error Resume Next
    If Selection.Font.Underline = 0 And Selection.Font.Size = ActiveDocument.Styles("normal").Font.Size Then
        Application.ScreenUpdating = False
        
        ActiveDocument.Content.Font.Shrink
        WordBasic.SelectSimilarFormatting
        ActiveDocument.Content.Font.Grow
    
        Application.ScreenUpdating = True
    Else
        WordBasic.SelectSimilarFormatting
    End If
End Sub

Sub D8_PasteURL()
'Pastes plain text with the URL from the top-most browser window.
    
    On Error Resume Next
    
    Dim sys32 As Object, procList As Object, proc As Object, _
        curID As Long, ffID As Long, _
         bwsr As Long, _
        CBChk As New DataObject, CurCB As New DataObject, cPos
     'ie As InternetExplorer, ieShell As Object, , ieShells As New Shell
   
    Selection.Collapse
    cPos = Selection.End
    Call D8_PasteText
    Selection.End = cPos
    
    CurCB.SetText Selection, 1
    
    
    'access system processes to find Firefox process handle
    Set sys32 = GetObject("winmgmts:")
    Set procList = sys32.execquery("SELECT * FROM win32_process")
    For Each proc In procList
        If proc.Name = "firefox.exe" Then ffID = proc.Handle
    Next
        
    If ffID Then 'find Firefox window handle
        bwsr = FindWindow("", "")
        Do While bwsr
            GetWindowThreadProcessId bwsr, curID
            If curID = ffID Then Exit Do
            bwsr = GetWindow(bwsr, 2)
        Loop
    Else 'or get IE as shell object, then find IE window handle
     '   For Each ieShell In ieShells.Windows
      '      If TypeName(ieShell.Document) = "HTMLDocument" Then Set ie = ieShell
       ' Next ieShell
            
        'If ie Is Nothing Then Exit Sub
        'bwsr = ie.HWND
    End If
    
    'bring up Firefox/IE
    SetForegroundWindow bwsr
    
    'alt+d
    keybd_event &H12, 0, 0, 0
    keybd_event Asc("D"), 0, 0, 0
    keybd_event &H12, 0, &H2, 0
    keybd_event Asc("D"), 0, &H2, 0
    'ctrl+c
    keybd_event &H11, 0, 0, 0
    keybd_event Asc("C"), 0, 0, 0
    keybd_event &H11, 0, &H2, 0
    keybd_event Asc("C"), 0, &H2, 0
       
    Sleep 350 'increase this value if nothing gets copied.
              'windows takes a bit to copy text to clipboard
            
    CBChk.GetFromClipboard
        If Not InStr(CBChk.GetText(1), "http") Then Sleep 1000
    
    'tab
    keybd_event &H9, 0, 0, 0
    keybd_event &H9, 0, &H2, 0
    
    Application.Activate
    CBChk.GetFromClipboard
    Selection = CBChk.GetText(1) & vbCrLf & vbCrLf
    Selection.Collapse
    CurCB.PutInClipboard
    
    Set sys32 = Nothing
    Set procList = Nothing
End Sub

Sub D8_PasteReturns()
' Pastes text without formatting and without returns
     
    Dim curPasteSetting
    curPasteSetting = D8x("Paste")
    
    D8s "Paste", True
    D8_PasteText
    D8s "Paste", curPasteSetting
    
    D8_RemoveReturns
End Sub
Sub D8_PasteText()
' Pastes text without formatting or Lexis Academic extra text.
    
    Dim PsChk As New DataObject, Ps, Ps1, Ps2, PsChk2 As New DataObject
  
    On Error GoTo Handler
    PsChk.GetFromClipboard
    If PsChk.GetFormat(1) = False Then Exit Sub
    Ps = PsChk.GetText(1)
    
    While InStr(Ps, "Enhanced Coverage Linking")
        Ps1 = RTrim(Left(Ps, InStr(Ps, "Enhanced Coverage Linking") - 1))
            If Mid(Ps1, Len(Ps1) - 1) = vbCrLf Then Ps1 = Left(Ps1, Len(Ps1) - 2)
            'gotta use second line test to make it work with Google Chrome
        Ps2 = LTrim(Mid(Ps, InStr(Ps, "Most Recent 60 Days") + 21))
            If Left(Ps2, 2) = vbCrLf Then Ps2 = Mid(Ps2, 3)
        
        Ps = Ps1 & " " & Ps2
    Wend
     
    If InStr(Ps, "Find Documents with Similar Topics") Then _
        Ps = Left(Ps, InStr(Ps, "Find Documents with Similar Topics") - 17)
         
    Ps = Replace(Ps, "Click here to return to the footnote reference.", "")
    
    '
    If InStr(Ps, " ") And Len(Selection) = 1 Then
        If Right(Ps, 1) <> " " Then Ps = Ps & " "
        
        If InStr(" ", Selection) Then
            Selection.MoveRight
        Else
            If Left(Ps, 1) <> " " Then Ps = " " & Ps
        End If
        
    End If
    
    Selection = Ps
    
    If D8x("Paste") = "False" Then _
        Selection.Collapse 0
    
Handler:
End Sub


Sub D8_RemoveReturns()
' Removes line breaks in selection.

    Application.ScreenUpdating = False
    
    With Selection
        If .Type < 2 Then Exit Sub
        If .Characters.Last = vbCr Then .MoveEnd , -1
        With .Find
            .Wrap = wdFindStop
            .ClearFormatting
            If InStr(Selection, Chr(10)) Then _
                .Execute findtext:=Chr(10), Wrap:=wdFindStop, replacewith:=Chr(13), Replace:=wdReplaceAll 'linefeeds to returns
            If InStr(Selection, Chr(45) & Chr(13)) Then _
                .Execute findtext:=Chr(45) & Chr(13), replacewith:="", Replace:=wdReplaceAll 'word wrap in pdf
            If InStr(Selection, Chr(172)) Then _
                .Execute findtext:=Chr(172), replacewith:="", Replace:=wdReplaceAll 'invisible word wrap in pdf
            If InStr(Selection, Chr(160)) Then _
                .Execute findtext:=Chr(160), replacewith:=" ", Replace:=wdReplaceAll  'non breaking spaces (used in some pdfs)
            If InStr(Selection, Chr(9)) Then _
                .Execute findtext:=Chr(9), replacewith:=" ", Replace:=wdReplaceAll 'tabs
            If InStr(Selection, Chr(11)) Then _
                .Execute findtext:=Chr(11), replacewith:=" ", Replace:=wdReplaceAll 'vertical tabs
            If InStr(Selection, Chr(12)) Then _
                .Execute findtext:=Chr(12), replacewith:=" ", Replace:=wdReplaceAll 'page breaks
            If InStr(Selection, Chr(13)) Then _
                .Execute findtext:=Chr(13), replacewith:=" ", Replace:=wdReplaceAll 'returns!
            
            While InStr(Selection, "  ")
                .Execute findtext:="  ", replacewith:=" ", Replace:=wdReplaceAll
            Wend
        End With
        If .Characters(1) = " " And _
            .Paragraphs(1).Range.Start = .Start Then _
            .Characters(1).Delete
        
    End With
    Application.ScreenUpdating = True
End Sub

Sub D8_FormatBox()
' Toggles box borders around selection.

    With Options
        .DefaultBorderLineStyle = 1
        .DefaultBorderLineWidth = 4
        .DefaultBorderColor = 0
    End With
    With Selection
        If .Type < 2 Then
            .StartOf wdWord
            .MoveEnd wdWord
        End If
        If .Font.Borders(1).LineStyle Then
            .Font.Borders(1).LineStyle = 0
        Else
            .Font.Borders(1).LineStyle = 1
        End If
    End With
End Sub

Sub D8_FormatHat()
' Sets selection as a main section heading, emphasized in the table of contents.

    With Selection
        If .Paragraphs.Count > 1 Or Len(Selection) > 40 Then
            If MsgBox("Are you sure you would like to make this large amount of text a Section Hat?", _
                vbYesNo + vbQuestion + vbDefaultButton2, "Hat Format") <> vbYes Then Exit Sub
        End If
        
        If .Paragraphs(1).Range.Start = ActiveDocument.Content.Start Then
            .HomeKey
            .TypeParagraph
        End If
        ActiveDocument.CopyStylesFromTemplate NormalTemplate.FullName
        .Style = ActiveDocument.Styles("Hat")
        .Range.Case = wdUpperCase
        .HomeKey
        If Selection <> "*" Then .TypeText "***"
        .HomeKey
        .TypeParagraph
        .MoveUp wdLine, 2
    End With
End Sub
Sub D8_FormatHeading()
' Sets selection as block heading.
    On Error Resume Next
    
    With Selection
    
        If .Paragraphs.Count > 1 Or Len(Selection) > 40 Then
            If MsgBox("Are you sure you would like to make this large amount of text a Block Heading?", _
                vbYesNo + vbQuestion + vbDefaultButton2, "Heading Format") <> vbYes Then Exit Sub
        End If
        
        
        If Not ActiveDocument.Styles("Heading 1").ParagraphFormat.PageBreakBefore Then .InsertBreak 7
        ActiveDocument.CopyStylesFromTemplate NormalTemplate.FullName
        If .Type < 2 Then .Paragraphs(1).Range.Select
     End With
     With Selection
    
        .Style = ActiveDocument.Styles("Heading 1")
        .Font.Reset
        If Len(.Paragraphs(1).Range) = 1 Then Exit Sub
        
        .EndKey
        .TypeParagraph
    End With
End Sub
Sub D8_FormatNotHeading()
' Removes selection from appearing in the document map.
    Dim selSize, selFont, selBold, selUndr, selAlgn, p As Paragraph
    
    With Selection
        .Range.ParagraphFormat.OutlineLevel = 10
        If .Range.ParagraphFormat.OutlineLevel = 10 Then Exit Sub
        If MsgBox("Would you also like to remove the text formatted " & _
            "using styles from appearing in the document map?", vbYesNo + _
            vbQuestion, "Remove From Document Map") <> vbYes Then Exit Sub
        
        For Each p In .Paragraphs
            If p.Range.ParagraphFormat.OutlineLevel <> 10 And _
                p.Range.ParagraphFormat.OutlineLevel <> 9999999 Then
                    With p.Range.Font
                        selSize = .Size
                        selFont = .Name
                        selBold = .Bold
                        selUndr = .Underline
                    End With
                    selAlgn = p.Range.ParagraphFormat.Alignment
                    
                    p.Range.Style = ActiveDocument.Styles("Normal")
                    With p.Range.Font
                        .Size = selSize
                        .Name = selFont
                        .Bold = selBold
                        .Underline = selUndr
                    End With
                    p.Range.ParagraphFormat.Alignment = selAlgn
            End If
        Next
    End With
End Sub
Sub D8_FormatHighlight()
' Toggles the highlighting of the selection.
    
    If Options.DefaultHighlightColorIndex = 0 Then Options.DefaultHighlightColorIndex = wdYellow
    
    With Selection
        If .Type < 2 Then
            .StartOf wdWord
            .MoveEnd wdWord
        End If
    End With

    WordBasic.Highlight
End Sub
Sub D8_FormatNormal()
' Sets selection to default font and clears all formatting except underlining.

    ActiveDocument.CopyStylesFromTemplate NormalTemplate.FullName
    With Selection
        .Style = ActiveDocument.Styles("Normal")
        .Range.HighlightColorIndex = 0
        .Font.Shading.Texture = wdTextureNone
        .Font.Name = ActiveDocument.Styles("Normal").Font.Name
        .Font.Size = ActiveDocument.Styles("Normal").Font.Size
    End With
End Sub
Sub D8_FormatToggle()
' Toggles the selection between underlined and non-underlined small size.

    On Error Resume Next
    Application.ScreenUpdating = False
    
    With Selection
        If .Type < 2 Then
            .StartOf wdWord
            .MoveEnd wdWord
        End If
        
        With .Font
            If .Underline = 0 Then
                .Size = ActiveDocument.Styles("Normal").Font.Size
                .Underline = 1
            Else
                .Size = D8x("Small")
                .Underline = 0
            End If
        End With
    End With
    Application.ScreenUpdating = True
End Sub
Sub D8_FormatToggleAuto()
' Automatically toggles any selected text between underlined and non-underlined small size.
 
    Dim SizeBig, SizeSmall
    SizeBig = ActiveDocument.Styles("Normal").Font.Size
    SizeSmall = D8x("Small")
    
    Do
        DoEvents
        SetCursor 9
        
        With Selection
            If .Type > 1 Then
                With .Font
                        If .Underline = 0 Then
                            .Size = SizeBig
                            .Underline = 1
                        Else
                            .Size = SizeSmall
                            .Underline = 0
                        End If
                End With
                .Collapse 0
            End If
        End With
        
    Loop Until GetKeyState(&H1B) < 0
    
    SetCursor 1

End Sub

Sub D8_FormatSmallAll()
'Makes all nonunderlined and nonbolded text in selection small size
'and all underlined or bolded text in selection normal size.

    Application.ScreenUpdating = False
             
    If Selection.Start <> Selection.Paragraphs(1).Range.Start Then _
        Selection.Paragraphs(1).Range.Select
    
    If Selection.Type < 2 Then Exit Sub
    
    
    With Selection.Range.Find
        .ClearFormatting
        .Text = ""
        .Wrap = wdFindStop
        .Replacement.Text = ""
        
        .Font.Underline = 1
        .ParagraphFormat.OutlineLevel = 10
        .Replacement.Font.Size = ActiveDocument.Styles("Normal").Font.Size
        .Execute Replace:=wdReplaceAll
        
        .Font.Underline = 0
        .Font.Bold = True
        .Execute Replace:=wdReplaceAll
        
        .Font.Bold = False
        .Font.Underline = 0
        .Replacement.Font.Size = D8x("Small")
        .Execute Replace:=wdReplaceAll
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
        
    Application.ScreenUpdating = True
    
End Sub

Sub D8_FormatSmallAllMore()
'Downsizes all nonunderlined and nonbolded text in selection.

    Application.ScreenUpdating = False
    
    
    Dim cEnd, cStart
    With Selection
        cEnd = Selection.End
        cStart = Selection.Start
        
        'fix multi-selection bug
        Selection.End = cEnd
        Selection.Start = cStart
        
        If .Font.Size <> 9999999 Then .Font.Shrink
        
        With .Find
            .ClearFormatting
            .Text = ""
            .Replacement.Text = ""
            .Font.Underline = 0
            .Font.Bold = False
            .ParagraphFormat.OutlineLevel = 10
        End With
        
        
        
        Do While .Find.Execute And .Start < cEnd
            .Font.Shrink
        Loop
        
        .End = cEnd
        .Start = cStart
        .Find.ClearFormatting
        .Find.Replacement.ClearFormatting
    End With
    
    Application.ScreenUpdating = True
End Sub
