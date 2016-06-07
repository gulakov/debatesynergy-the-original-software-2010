Attribute VB_Name = "d8advanced"
Sub D8_FixCiteRequest()
' Deletes all but the first and last words from the selection
' or current paragraph, then outputs in new document and copies the cite.
        
    On Error Resume Next
    Dim WordsInCite, X As Long: WordsInCite = D8x("CiteWords")
    
    Application.ScreenUpdating = False
    
    'copy to new document
    With Selection
        If .Type < 2 Then .MoveStart wdParagraph, -4
        Dim Pr: Pr = .Paragraphs.Count
        .StartOf wdParagraph
        .MoveEnd wdParagraph, Pr
        .Copy
    End With
    Documents.Add.Range.Paste
    Selection.EndOf wdStory
    Selection.TypeText vbCr & vbCr & vbCr & vbCr & vbCr
    If ActiveDocument.TablesOfContents.Count Then _
        ActiveDocument.TablesOfContents(1).Delete
    
    With Selection
        .StartOf wdStory
            
        While ActiveDocument.Range.End - .End > 5
            
           
            Do Until .Paragraphs.First.Range.Font.Underline = 9999999
                If ActiveDocument.Range.End - .End < 5 Then GoTo EndLoop
                
                .MoveStart wdParagraph
            Loop
            
            Do
                If ActiveDocument.Range.End - .End < 5 Then GoTo EndLoop
            .MoveEnd wdParagraph
            Loop Until .Paragraphs.Last.Range.Font.Underline <> 9999999
            
            .MoveEnd wdParagraph, -1
            
            
            For X = 1 To WordsInCite
            .MoveStart wdWord
                If (LCase(.Words.First) = UCase(.Words.First) Or Len(.Words.First) < 3) _
                    And Trim(.Words.First) <> "…" Then X = X - 1
                 If ActiveDocument.Range.End - .End < 5 Then GoTo EndLoop
            Next X
            
            For X = 0 To WordsInCite
            .MoveEnd wdWord, -1
                If (LCase(.Words.Last) = UCase(.Words.Last) Or Len(.Words.Last) < 3) _
                    And Trim(.Words.Last) <> "…" Then X = X - 1
                 If ActiveDocument.Range.End - .End < 5 Then GoTo EndLoop
            Next X
            .TypeText "… "
            
               
            .Move wdParagraph
        Wend
    End With
    
EndLoop:
    ActiveDocument.Range.Copy
    Application.ScreenUpdating = True
End Sub

Sub D8_WarrantAdd()

    Dim dm, cm As Comment
    dm = ActiveWindow.DocumentMap
    
    ActiveWindow.View.ShowRevisionsAndComments = True
    ActiveWindow.View.MarkupMode = wdBalloonRevisions
    Set cm = Selection.Comments.Add(Selection.Paragraphs(1).Range)
    
    ActiveWindow.DocumentMap = dm

    cm.ShowTip = True
    cm.Edit
    
    Fresh
    
    Set cm = Nothing

End Sub

Sub D8_WarrantToggle()

    With ActiveWindow.View
        If .ShowRevisionsAndComments Then
            .ShowRevisionsAndComments = False
        Else
            .ShowRevisionsAndComments = True
            .MarkupMode = wdBalloonRevisions
        End If
    End With
End Sub




Sub D8_TOC()
' Inserts (or updates) a front page containing the table of contents, and removing like entries.
   
    Dim t1, t2, tSum
    
    Application.ScreenUpdating = False
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    With ActiveDocument
        ' Set style of regular entries in TOC
        With .Styles("TOC 1")
            .BaseStyle = "Normal"
            .Font.Size = ActiveDocument.Styles("Normal").Font.Size
            .Font.Bold = ActiveDocument.Styles("Normal").Font.Bold
            .Font.Underline = ActiveDocument.Styles("Normal").Font.Underline
            .ParagraphFormat.LeftIndent = ActiveDocument.Styles("Normal").ParagraphFormat.LeftIndent
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.LineSpacing = 12
        End With
        
        ' Set style of Hat entries in TOC
        With .Styles("TOC 4")
            .BaseStyle = "Normal"
            .Font.Bold = True
            .Font.Underline = wdUnderlineSingle
            .ParagraphFormat.LeftIndent = InchesToPoints(0)
            .ParagraphFormat.SpaceBefore = 12
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.LineSpacing = 12
            .Font.Name = ActiveDocument.Styles("Normal").Font.Name
        End With

        If .TablesOfContents.Count Then
            .TablesOfContents(1).Update
            GoTo TOCFormat
        End If
    End With
             
    With Selection
        .HomeKey wdStory
        .InsertParagraphBefore
        .Style = ActiveDocument.Styles("Heading 1")
        
        If Right(ActiveDocument.Name, 4) = ".doc" Then .TypeText Replace(ActiveDocument.Name, ".doc", "")
        If Right(ActiveDocument.Name, 5) = ".docx" Then .TypeText Replace(ActiveDocument.Name, ".docx", "")
        If ActiveDocument.Path = "" Then .TypeText "Table of Contents"
        
        .TypeParagraph
        .ClearFormatting
        .TypeParagraph
    
        ActiveDocument.TablesOfContents.Add .Range, UseHeadingStyles:=True, _
            UpperHeadingLevel:=1, LowerHeadingLevel:=1, UseHyperlinks:=True, _
            AddedStyles:="Heading 1,1,Hat,4", UseOutlineLevels:=True
    End With

TOCFormat:
    
    With Selection
    Dim X
    If D8x("RemoveTOC") <> "True" Then
        .MoveLeft
        Application.ScreenUpdating = True
        Exit Sub
    End If
        .StartOf wdStory
        .GoTo wdGoToField, Name:="toc"
        .MoveEnd
        tSum = .Hyperlinks.Count
        .StartOf
        .MoveRight
        .MoveEnd wdLine
        t1 = LCase(Trim(Left(Selection, InStr(Selection, Chr(9)))))
        For X = 2 To tSum
            .MoveRight
            .MoveEnd wdLine
            t2 = LCase(Trim(Left(Selection, InStr(Selection, Chr(9)))))
            If t1 = t2 Then .Delete
            t1 = t2
        Next X
        ActiveDocument.TablesOfContents(1).UpdatePageNumbers
        .StartOf wdStory
        .GoTo wdGoToField, Name:="toc"
        .MoveRight
        .MoveEnd wdLine
        .Delete
        .MoveLeft
    End With
    Application.ScreenUpdating = True
    
End Sub


Sub D8_PageHeader()
' Inserts page header for all pages.
    On Error Resume Next
    Dim X
    Application.ScreenUpdating = False
    ActiveDocument.CopyStylesFromTemplate NormalTemplate.FullName
    
    With ActiveWindow.ActivePane.View
        If .SplitSpecial Then ActiveWindow.Panes(2).Close
        .Type = wdPrintView
    
        If ActiveDocument.Sections.Count > 1 Then
            If MsgBox("Display this page header in all sections?", vbYesNo) = vbYes Then
                    .SeekView = wdSeekMainDocument
                    Selection.StartOf wdStory
                    .SeekView = wdSeekCurrentPageHeader
                    For X = 2 To ActiveDocument.Sections.Count
                        .NextHeaderFooter
                        Selection.HeaderFooter.LinkToPrevious = True
                    Next X
            End If
        End If
    
        WordBasic.RemoveHeader
       .SeekView = wdSeekCurrentPageHeader
    End With
        
    With Selection
        .WholeStory
        .Delete
        .WholeStory
        .ParagraphFormat.TabStops.ClearAll
        
        'top line
        .Style = ActiveDocument.Styles("PageHeaderLine1")
        
        If Right(ActiveDocument.Name, 4) = ".doc" Then .TypeText Replace(ActiveDocument.Name, ".doc", "")
        If Right(ActiveDocument.Name, 5) = ".docx" Then .TypeText Replace(ActiveDocument.Name, ".docx", "")

        .TypeText vbTab
        
        .TypeText Mid(Application.UserName, InStr(Application.UserName, ",") + 1)
        
        .TypeParagraph
        
        'bottom line
        .Style = ActiveDocument.Styles("PageHeaderLine2")
        
        .Fields.Add .Range, wdFieldPage
        .TypeText "/"
        .Fields.Add .Range, wdFieldNumPages
        
        .TypeText vbTab

        .TypeText Left(Application.UserName, InStr(Application.UserName, ",") - 1)
    End With
    ActiveDocument.Repaginate
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    Application.ScreenUpdating = True
End Sub

'Sub D8_Timer()
'Displays debate timer.
 '   Read.Show
'End Sub


Sub D8_FixNoLinks()
' Deletes all hyperlinks in document.
    
    Dim DelCount As Long, i
    With ActiveDocument.Styles("Hyperlink").Font
        .Underline = 0
        .Color = wdColorAutomatic
    End With
    With ActiveDocument.Content
        For i = .Hyperlinks.Count To 1 Step -1
            If .Hyperlinks(i).Address > "" Then
                .Hyperlinks(i).Range.Font.Color = 0
                .Hyperlinks(i).Range.Font.Underline = 0
                 .Hyperlinks(i).Delete
                
                DelCount = DelCount + 1
            End If
        Next i
    End With
    StatusBar = DelCount & " hyperlinks deleted."
End Sub
Sub D8_FixPageContinued()
' Insert page break with a Continued note and current block title.

    Dim ContinueText, InsPos, BrChk
    ContinueText = D8x("Continues")
    
    Application.ScreenUpdating = False
    
    With Selection
        .Collapse 0
        If .Words(1).Start <> .Start Then .MoveRight wdWord
        .TypeParagraph
        .InsertParagraphAfter
        .ClearFormatting
        .TypeText ContinueText
        .TypeParagraph
        InsPos = .Start
        .GoTo wdGoToBookmark, Name:="\Page"
        .MoveEnd wdParagraph, 1 - .Paragraphs.Count
        .Copy
        .HomeKey
        BrChk = (.Range.ParagraphFormat.PageBreakBefore <> -1)
        .Start = InsPos
        If BrChk Then .InsertBreak wdPageBreak
        .Paste
        
        .TypeParagraph
        .Font.Underline = 0
        .TypeText ContinueText
        .TypeParagraph
    End With
    Application.ScreenUpdating = True
End Sub

Sub D8_FixAutoFormat()
' Formats entire document into the Normal Template
    
    Application.ScreenUpdating = False
    Dim pw, ph, tm, bm, lm, rm, li, fi, ri
    
    'margins
    With NormalTemplate.OpenAsDocument.Content
        With .PageSetup
            pw = .PageWidth
            ph = .PageHeight
            tm = .TopMargin
            bm = .BottomMargin
            lm = .LeftMargin
            rm = .RightMargin
        End With
        With .ParagraphFormat
            li = .LeftIndent
            fi = .FirstLineIndent
            ri = .RightIndent
        End With
    End With
    Documents(NormalTemplate.FullName).Close 0
    With ActiveDocument.Content
        With .PageSetup
                .PageWidth = pw
                .PageHeight = ph
                .TopMargin = tm
                .BottomMargin = bm
                .LeftMargin = lm
                .RightMargin = rm
        End With
        With .ParagraphFormat
                .LeftIndent = li
                .FirstLineIndent = fi
                .RightIndent = ri
        End With
    End With
    
    'styles
    ActiveDocument.CopyStylesFromTemplate NormalTemplate.FullName
    With Selection
        .StartOf wdStory
        With .Find
            .ClearFormatting
            .Text = ""
            .ParagraphFormat.OutlineLevel = 1
            .Format = True
            .Wrap = wdFindStop
        End With
            
        Do While .Find.Execute
            If .Range.Frames.Count = 0 Then .Style = "Heading 1"
            .Collapse 0
        Loop
    
    Call D8_FixBlankPages
        
        'toc
        If ActiveDocument.TablesOfContents.Count Then
            .GoTo wdGoToField, Name:="toc"
            .MoveEnd
            .Delete
            .GoTo wdGoToBookmark, Name:="\Page"
            .Delete
            Call D8_TOC
        End If
    End With
    ActiveDocument.Content.Font.Name = ActiveDocument.Styles("Normal").Font.Name
    Application.ScreenUpdating = True
End Sub

Sub D8_FixBlankPages()
' Removes all blank pages in document and displays number removed in status bar.
   
    Application.ScreenUpdating = False
    Dim CurS, CurE, PagesBefore, p
    ActiveDocument.Repaginate
    PagesBefore = ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
    With Selection
        CurS = .Start
        CurE = .End
        .HomeKey wdStory
            
        If ActiveDocument.Styles("Heading 1").ParagraphFormat.PageBreakBefore = True Then
            With .Find
                .ClearFormatting
                .Execute findtext:="^m", replacewith:="", Wrap:=wdFindContinue, Replace:=wdReplaceAll
                .Text = Chr(13)
                .ParagraphFormat.PageBreakBefore = True
                Do While .Execute
                    If Len(Selection.Paragraphs(1).Range) < 3 Then Selection.Delete
                    Selection.Collapse 0
                Loop
            End With
        ActiveDocument.Repaginate
        End If
        
        For p = 1 To ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
            .GoTo wdGoToBookmark, Name:="\Page"
            If Len(Trim(Selection)) < 2 Or .Words.Count = .Characters.Count Then
                .Delete
            Else
                .GoTo wdGoToPage
            End If
        Next p
        .Start = CurS
        .End = CurE
        .Find.ClearFormatting
    End With
    ActiveDocument.Repaginate
    StatusBar = PagesBefore - ActiveDocument.BuiltInDocumentProperties(wdPropertyPages) & " blank pages removed."
    Application.ScreenUpdating = True
    
End Sub
Sub D8_FixCaps()
' Formats the Selection in Title Case, Except Words Like "and".

    Application.ScreenUpdating = False
    Dim w
    With Selection.Range
        .Case = wdTitleWord
        For Each w In .Words
        If InStr(" a an and are as at but by for from in into is of off on onto or out the this to up with ", _
            " " & Trim(LCase(w)) & " ") And w <> .Words.First Then w.Case = wdLowerCase
        Next w
    End With
    Application.ScreenUpdating = True
End Sub



