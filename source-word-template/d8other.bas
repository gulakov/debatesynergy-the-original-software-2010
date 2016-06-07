Attribute VB_Name = "d8other"
Dim isScrolling As Boolean


Sub D8_ScrollDown()
' Scrolls down more smoothly.
    If isScrolling Then Exit Sub
    isScrolling = True
    
    ActiveWindow.VerticalPercentScrolled = ActiveWindow.VerticalPercentScrolled + 1
    
    Dim ScrollStart
    ScrollStart = Timer
    Do While Timer - ScrollStart < 0.6
        DoEvents
    Loop
    
    isScrolling = False
End Sub
Sub D8_PDFConvert()
' Converts the document to a .pdf file.

    On Error Resume Next
    With ActiveDocument
        Dim PDFName, i
        PDFName = Left(.FullName, InStr(.FullName, ".doc") - 1)
        If Dir(PDFName & ".pdf") > "" Then
            For i = 2 To 100
                If Dir(PDFName & " " & i & ".pdf") = "" Then Exit For
            Next i
            PDFName = PDFName & " " & i
        End If
        If .Path = "" Then .SaveAs
        .ExportAsFixedFormat PDFName & ".pdf", 17, True, CreateBookmarks:=1
        If Err Then MsgBox "You do not have PDF Exporting installed."

    End With
End Sub

Sub D8_MergeDocs()
'Merges multiple documents into a single document.

    Dim doNames As Boolean, d
    
    Select Case MsgBox("Control click to select multiple documents. All selected files will " & _
    "then be merged into a single document." & vbCr & vbCr & "Would you like to use the file" & _
        "names as block headers?", vbQuestion + vbYesNoCancel, "File Converter")
    
        Case vbYes: doNames = True
        Case vbNo: doNames = False
        Case Else: Exit Sub
    End Select
    
    With Application.FileDialog(1)
        .InitialView = 2
        .Title = "Control click to select multiple files"
        .AllowMultiSelect = True
        .ButtonName = "Merge"
        .Filters.Clear
        .Filters.Add "Documents", "*.doc,*.docx,*.rtf"
        
        If .show Then
            'create output document
            Application.ScreenUpdating = False
            Documents.Add
        
            'combine files
            For Each d In .SelectedItems
               If doNames Then
                    Selection.TypeText Replace(Right(d, InStr(StrReverse(d), "\") - 1), _
                        Right(d, InStr(StrReverse(d), ".")), "")
       
                    Call D8Format.D8_FormatHeading
                    Selection.TypeParagraph
                End If
                
                Selection.InsertFile d
            Next d
        End If
    End With
    
    Application.ScreenUpdating = True
End Sub
'BLOCK MOVE
Public Function D8_BlockSelect()
    
    Application.ScreenUpdating = False
    With Selection
        If .Range.Frames.Count Then
            .Range.Frames(1).Select
        Else
            .HomeKey
            
            While .Paragraphs(1).Range.ParagraphFormat.OutlineLevel <> 1 And _
                .Start <> ActiveDocument.Range.Start
                    .MoveUp wdParagraph, 1, 1
            Wend
           
            If .Paragraphs.Count = 1 Then .MoveEnd wdParagraph, 2
            
            While .Paragraphs(.Paragraphs.Count).Range.ParagraphFormat.OutlineLevel <> 1 And _
                .End <> ActiveDocument.Range.End
                    .MoveEnd wdParagraph
            Wend
            
            If .End <> ActiveDocument.Content.End Then .MoveUp wdParagraph, 1, 1
        End If
        
    End With
End Function
Sub D8_BlockStart()
'Moves current block or hat to after the table of contents.

    With Selection
        D8_BlockSelect
        .Cut
        .StartOf wdStory
        If ActiveDocument.TablesOfContents.Count Then
            .GoTo wdGoToField, Name:="toc"
            .MoveEnd
            .EndOf
            .TypeParagraph
            .MoveRight
        End If
        .Range.Paste
        If Selection = " " Then .Delete
        Application.ScreenUpdating = True
        .EndKey
        .HomeKey
    End With
End Sub
Sub D8_BlockEnd()
'Moves current block or hat to the end of the document.

    With Selection
        D8_BlockSelect
        .Cut
        .EndOf wdStory
        If Len(.Paragraphs(1).Range) > 1 Then .TypeParagraph
        .ClearFormatting
        .Range.Paste
        Application.ScreenUpdating = True
        .EndKey
        .HomeKey
    End With
End Sub
Sub D8_BlockUp()
'Moves current block or hat up.
  
   D8_BlockSelect
    With Selection
        If .Range.Frames.Count Then
            .Cut
            .MoveRight
            D8_BlockSelect
            .EndOf
            .MoveLeft
            
            .Paste
            .MoveLeft
        Else
            .Cut
            .MoveLeft
            D8_BlockSelect
            .StartOf
            
            If .Range.Frames.Count Then
                While .Range.Frames.Count And .Start > 0
                    .MoveUp
                Wend
                        
                If Len(.Paragraphs(1).Range) > 1 Then
                    .TypeParagraph
                    .ClearFormatting
                End If
            End If
            
            .Paste
            .MoveLeft
            D8_BlockSelect
            .StartOf
        End If
    End With
End Sub
Sub D8_BlockDown()
'Moves current block or hat down.

    D8_BlockSelect
    With Selection
        If .Range.Frames.Count Then
            .Cut
            .MoveRight
            D8_BlockSelect
            .EndOf
            .MoveLeft
            If Len(.Paragraphs(1).Range) > 1 Then
                .TypeParagraph
                .ClearFormatting
            End If
            .Paste
            .MoveLeft
        Else
            .Cut
            D8_BlockSelect
            .EndOf
            
            If .Range.Frames.Count Then
                While .Range.Frames.Count
                    .MoveUp
                Wend
                .MoveUp
                While Len(.Paragraphs(1).Range) = 1
                    .MoveUp
                Wend
            Else
                .MoveLeft
            End If
                       
            If Len(.Paragraphs(1).Range) > 1 Then
                .TypeParagraph
                .ClearFormatting
            End If
            
            .Paste
            .MoveLeft
            D8_BlockSelect
            .StartOf
        End If
    End With
End Sub

'OTHER
Sub D8_ViewRecovered()
'View folder with AutoRecovery documents.

    Dim r
    With Application.FileDialog(1)
        .InitialFileName = Options.DefaultFilePath(5) & "\AutoRecovery save of *.asd"
        .InitialView = 2
        .ButtonName = "Recover"
        If .show Then
            For Each r In .SelectedItems
                Documents.Open r
            Next r
        End If
    End With
End Sub
Sub D8_WindowCycle()
'Cycle through all open windows

    On Error Resume Next
    Dim i As Long
    For i = 1 To Documents.Count
        If Documents(i).Name = ActiveDocument.Name Then Exit For
    Next i
    
    If i = 1 Then i = Documents.Count + 1
    Documents(i - 1).Activate
    
End Sub


Sub D8_SaveSession()
' Saves currently open documents to be restored the next time Word starts.
    
    Dim d
    If MsgBox("Would you like to save all currently open documents? " & _
        vbCr & "They will be restored the next time Word starts.", _
        vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    For d = 1 To Documents.Count
        Documents(d).SaveAs
        D8s "Doc" & d, Documents(d).FullName, "SessionSave"
    Next d
    D8s "Date", "from " & DateTime.Date & " " & DateTime.Time, "SessionSave"
    D8s "Count", Documents.Count, "SessionSave"

    MsgBox "Your session has been saved and will open the next time Word starts.", vbApplicationModal
End Sub
Sub d8filesave(control As IRibbonControl, ByRef cancelDefault)
' This overrides the built-in Word command.
' Save by current date and user name.

    
    On Error Resume Next
    If ActiveDocument.Path > "" And control.id = "FileSave" Then
        ActiveDocument.Save
    Else
        Dim SN, SaveName
        SN = Application.UserName
        If InStr(SN, ",") Then SN = Left(SN, InStr(SN, ",") - 1)
        
        SaveName = Replace(ActiveDocument.Name, ".doc", "")
        
        If Left(SaveName, 8) = "Document" Then
            SaveName = ActiveDocument.Content.Paragraphs(1).Range
            SaveName = Trim(Left(SaveName, InStrRev(Left(SaveName, 30), " ")))
        End If
        
        With Dialogs(wdDialogFileSaveAs)
            .Name = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SaveName & " " & DateTime.Month(Now) & _
                "-" & DateTime.Day(Now) & " " & SN & ".doc"
            .Format = 0
            .show
        End With
    End If
End Sub
