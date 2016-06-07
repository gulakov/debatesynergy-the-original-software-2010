Attribute VB_Name = "d8startup"



Sub AutoExec()
    On Error Resume Next
    
    'initialize settings on first run
    FirstRun
     
    'change window caption
    Application.Caption = "Debate Synergy"
   
    'restore saved sessions
        If D8x("Count", "SessionSave") > "" Then
            Dim SaveName, d
            If MsgBox("Open the saved session " & D8x("Date", "SessionSave") & _
                "?", vbYesNo) = vbYes Then
                For d = 1 To D8x("Count", "SessionSave")
                    SaveName = D8x("Doc" & d, "SessionSave")
                    If Dir(SaveName) > "" Then Documents.Open SaveName
                    D8s "Doc" & d, "", "SessionSave"
                Next d
                D8s "Count", "", "SessionSave"
            ElseIf MsgBox("Delete the saved session?", vbYesNo) = vbYes Then _
                    D8s "Count", "", "SessionSave"
            End If
        End If
End Sub
Sub AutoOpen()
   On Error Resume Next
   
    'Show correct page count on open
    If D8x("PageCount") = "True" Then _
    ActiveDocument.Repaginate
    
    If D8x("startview") = "True" Then _
        ActiveWindow.View.Type = wdWebView
    
    
    'Go to the last edit position when opening a document
    If D8x("LastEdit") = "True" And ActiveDocument.Bookmarks.Exists("LastEdit") Then _
    Selection.GoTo wdGoToBookmark, Name:="LastEdit"
    
    'All documents open in 100% zoom.
    ActiveWindow.ActivePane.View.Zoom.Percentage = 100
      
    'Puts the active template name in the window title
    Application.Caption = ActiveDocument.BuiltInDocumentProperties(wdPropertyTemplate)
    If Application.Caption = "Normal.dotm" Then Application.Caption = "Debate Synergy"
   
    'All documents open with document map.
    ActiveWindow.DocumentMap = True
    
    'Use verdana for document map font.
    ActiveDocument.Styles("Document Map").Font.Name = "Verdana"
End Sub
Sub AutoNew()

     If D8x("startview") = "True" Then _
        ActiveWindow.View.Type = wdWebView
   
End Sub
Sub AutoClose()
    On Error Resume Next
    
    'Go to the last edit position when opening a document
    If D8x("LastEdit") = "True" Then _
        If ActiveDocument.Words.Count > 1 Then ActiveDocument.Bookmarks.Add "LastEdit", Selection.Range
End Sub
