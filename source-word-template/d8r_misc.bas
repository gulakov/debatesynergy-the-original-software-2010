Attribute VB_Name = "d8r_misc"

'EVERYTHING SEARCH
Sub s_go(control As IRibbonControl, Text As String)
    On Error GoTo Err
    
    Dim progPath
    progPath = D8x("EveryProg")
    
    If Dir(progPath) = "" Then GoTo Err
    
    ShellExecuteA 0, vbNullString, progPath, _
        "-search ""\""" & D8x("EveryPath") & "\\"" " & Text & """", vbNullString, 1
    
    Exit Sub
Err: MsgBox "The Everything Search Engine is not found. Please install it from voidtools.com " & _
    "and then point to its location in Debate Synergy options.", vbCritical, "Program Not Found"
    
End Sub


'WARRANTS
Sub c_show(control As IRibbonControl, ByRef returnedVal)
    With ActiveWindow.View
        returnedVal = (.ShowRevisionsAndComments And .MarkupMode = wdBalloonRevisions)
    End With
End Sub
Sub c_click(control As IRibbonControl, pressed As Boolean)
    With ActiveWindow.View
        If pressed Then
            .ShowRevisionsAndComments = True
            .MarkupMode = wdBalloonRevisions
        Else
            .ShowRevisionsAndComments = False
        End If
    End With
End Sub



'RATINGS
Sub q_rate(control As IRibbonControl)

    If Selection.Type < 2 Then
        MsgBox "This feature allows you to assign a Best, " & _
            "Medium, or Worst rating to a text selection in the document. You can then " & _
            "select to show all text of a particular rating." & vbCr & vbCr & _
            "You must select some text before running this feature.", vbInformation, "Quality Control"
      Exit Sub
    End If

  Dim i As Long
For i = 1 To 100
     If Not ActiveDocument.Bookmarks.Exists(control.id & "s" & i) Then Exit For
  Next i
   ActiveDocument.Bookmarks.Add control.id & "s" & i, Selection.Range
End Sub
Sub q_show(control As IRibbonControl, id As String, index As Integer)
    Dim i As Long

 If ActiveDocument.Bookmarks.Count = 0 Then _
      MsgBox "This feature enables you assign one of three quality ratings " & _
      "to a selection and then to select which quality of evidence to " & _
      "show temporarily." & vbCr & vbCr & "Before using this, please use " & _
       "the star buttons to assign some ratings.", vbInformation, "Quality Control"
  ActiveDocument.Content.Font.Hidden = Not (id = "q0s")
   For i = 1 To 100
       If ActiveDocument.Bookmarks.Exists(id & i) Then _
            ActiveDocument.Bookmarks(id & i).Range.Font.Hidden = False
   Next i
End Sub


'WINDOW CONTROL
Sub w_fill(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    
    Dim displayList As String, WinItemSize, w, d, docpath
    
    Set oList = New Collection
    
    w = 1
    For Each d In Documents
        oList.Add d.Name
        w = w + 1
        
        docpath = PurgePath(d.Path)
        If docpath = "" Then docpath = "Unsaved"
        
        If d = ActiveDocument Then
            displayList = displayList & "<checkBox getPressed=""w_wincheck"" id=""w" & w & _
                """ onAction=""w_refresh"" description=""" & docpath & """ label=""" & purge(d.Name) & """  />"
        ElseIf InStr(LCase(d.Name), "speech") Then
                displayList = displayList & "<button imageMso=""MicrosoftVisualFoxPro"" id=""w" & w & _
                    """ description=""" & docpath & """  label=""" & purge(d.Name) & """ onAction=""w_open"" />"
            Else
                displayList = displayList & "<button id=""w" & w & _
                    """ description=""" & docpath & """  label=""" & purge(d.Name) & """ onAction=""w_open"" />"
        End If
    Next
        
        
        
    If Documents.Count > 7 Then
        WinItemSize = "normal"
    Else
        WinItemSize = "large"
    End If
 
    returnedVal = "<menu itemSize=""" & WinItemSize & """ xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" & _
        "<button description=""Toggle between Print Layout and Paperless Layout."" id=""d8vswitch"" getImage=""w_viewimage"" getLabel=""w_viewlabel"" onAction=""w_viewswitch"" />" & _
        "<button description=""Toggle Full Screen mode. (Shortcut is `)"" id=""d8spread"" label=""Read Speech"" imageMso=""ZoomFitToWindow"" onAction=""main""/>" & _
        "<button description=""View the Excel flow and the Word speech side-by-side."" id=""winsbs"" image=""d8sbs"" label=""Flow Side-by-Side"" onAction=""w_flow"" />" & _
        "<menu supertip=""Show or hide view components."" id=""wmodes"" imageMso=""PageSizeGallery"" label=""Show/Hide"" description=""Show or hide the document map, rulers, and page separators."" >" & _
        "<checkBox supertip=""Toggle whether page separators are displayed."" id=""d8white"" getPressed=""w_whitepress"" label=""Page Separators"" onAction=""w_white"" />" & _
        "<checkBox idMso=""ViewDocumentMap"" /> <checkBox idMso=""ViewRulerWord"" />" & _
        "</menu> <menuSeparator id=""wsep"" />" & displayList & " </menu>"
    
    Fresh
        
End Sub
Sub w_viewswitch(control As IRibbonControl)
    With ActiveWindow.View
        If .Type = wdWebView Then
            .Type = wdPrintView
        Else
            .Type = wdWebView
        End If
    End With
    Fresh
End Sub
 Sub w_viewlabel(control As IRibbonControl, ByRef returnedVal)
    With ActiveWindow.View
        If .Type = wdWebView Then returnedVal = "Go to Print Layout"
        If .Type = wdPrintView Then returnedVal = "Go to Paperless Layout"
    End With
End Sub
 Sub w_viewimage(control As IRibbonControl, ByRef returnedVal)
    With ActiveWindow.View
        If .Type = wdWebView Then returnedVal = "ViewPrintLayoutView"
        If .Type = wdPrintView Then returnedVal = "CreateReportBlankReport"
    End With
End Sub

Sub w_wincheck(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub
Sub w_flow(control As IRibbonControl)
    w_flowresize
End Sub
Function w_flowresize()
    On Error Resume Next
    
    Dim rW, rH, oExcel As Object
    
    rW = 0.45 ' set percentage of screen width that the flow should occupy (default: 0.45)
    rH = 0.97 ' set percentage of screen height that both windows should occupy (default: 0.97)
     
    Set oExcel = GetObject(, "Excel.Application")
    With oExcel
        .WindowState = xlNormal
        .Width = 0.75 * Screen(0) * rW
        .Height = 0.75 * Screen(1) * rH
        .Left = 0
        .Top = 0
    End With
    With Application
        .WindowState = wdWindowStateNormal
        .Width = 0.75 * Screen(0) * (1 - rW)
        .Height = 0.75 * Screen(1) * rH
        .Left = 0.75 * Screen(0) * rW
        .Top = 0
        .Activate
    End With
    ActiveWindow.View.Type = wdWebView
    
    Application.ScreenRefresh
    Application.ScreenUpdating = True
    Set oExcel = Nothing
    
End Function
Sub w_refresh(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    
    Dim d As Document, ActDoc As Document
    
    If MsgBox("Close all documents other than this one?", _
        vbQuestion + vbYesNo, "Close All Others") = vbYes Then
        
        Set ActDoc = ActiveDocument
        For Each d In Documents
            If d <> ActDoc Then d.Close
        Next
    
    End If
    
    Fresh
End Sub
Sub w_open(control As IRibbonControl)
    On Error Resume Next
    Documents(Mid(control.id, 2) - 1).Activate
    
    Fresh
End Sub
Sub w_whitepress(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ActiveWindow.View.DisplayPageBoundaries
End Sub
Sub w_white(control As IRibbonControl, pressed As Boolean)
    ActiveWindow.View.DisplayPageBoundaries = pressed
End Sub


