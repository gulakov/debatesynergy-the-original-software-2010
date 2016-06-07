Attribute VB_Name = "D8Flow"
Sub D8_Comment()
' Inserts a sticky-note on cell
    
    On Error Resume Next
    
    With ActiveCell.Range("A1")
        On Error Resume Next
        .AddComment
            If Err Then Exit Sub
        .Comment.Visible = True
        .Comment.Text Text:=""
        .Comment.Shape.Select
    End With
    Selection.Font.Size = 8
    With Selection.ShapeRange
        .Fill.ForeColor.RGB = RGB(255, 255, 150)
        .Shadow.Visible = False
        .Fill.Transparency = 0.3
        .Line.Visible = 0
        .IncrementLeft -120
        .IncrementTop 8
        .ScaleWidth 0.6, msoFalse, msoScaleFromBottomRight
        .ScaleHeight 0.75, msoFalse, msoScaleFromTopLeft
    End With
    Application.DisplayCommentIndicator = xlCommentAndIndicator
End Sub
Sub D8_FlowAdd()
' Adds flow

    On Error GoTo Handler
    Dim ActSh, TypeSh, NewSh
    ActSh = ActiveSheet.Name
    TypeSh = "Offcase"
        If ActiveSheet.Tab.ColorIndex = 12 Then TypeSh = "Oncase"
    
    With Sheets(TypeSh)
        .Visible = True
        .Copy After:=Sheets(ActiveSheet.Index)
        .Visible = 2
    End With
    
SetName:
    NewSh = Trim(Replace(ActSh, Val(Right(ActSh, 2)), "")) & " " & _
        Trim(Str(Val(Right(ActSh, 2)) + 1 - (Val(Right(ActSh, 2)) = 0)))
    
    On Error Resume Next
    ActiveSheet.Name = NewSh
    If Err Then
        ActSh = NewSh
        GoTo SetName
    End If
    Exit Sub
    
Handler:
    Sheets(ActSh).Copy After:=Sheets(ActiveSheet.Index)
End Sub
Sub D8_FlowHide()
' Hides current sheet

    On Error Resume Next
    If ActiveSheet.Name = "Casebook" Then _
        If MsgBox("Are you sure you want to hide the Casebook sheet?" & _
        vbCr & vbCr & "Go to Format > Sheet > Unhide to restore any sheet.", _
        vbYesNo) = vbNo Then Exit Sub
         
    On Error Resume Next
    ActiveWindow.SelectedSheets.Visible = False
End Sub
Sub D8_Group()
' Group/ungroup cells
        
    On Error Resume Next
    With Selection.Areas(1).Columns(1).Offset(, (Selection.Areas(1).Columns(1).EntireColumn.Address = "$C:$C")).Borders(xlEdgeLeft)
        If .LineStyle <> xlDash Then
            .LineStyle = xlDash
            .Weight = xlMedium
            Select Case Selection.Areas(1).Cells(1, 1).Font.Color
                Case Blue, Blue3: .Color = 13893632
                Case Red, Red3: .Color = -16777024
            End Select
        Else
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .Color = 0
        End If
    End With
    FlowRefresh
End Sub
Sub D8_MoveDown()
' Moves cell down
    Application.ScreenUpdating = False
    On Error GoTo Handler
    
    With Selection.Areas(1)
        If WorksheetFunction.CountA(.Rows(.Rows.Count).Offset(1)) Then _
            If MsgBox("Are you sure you want to copy over a cell with text?", 4 + 0) <> 6 Then Exit Sub
                
        If .Columns(1).EntireColumn.Address = "$C:$C" Then .Cells(1, 1).Offset(, -1).ClearContents
              
        .Copy
        .Rows(1).Offset(1).Select
        ActiveSheet.Paste
                
        .Rows(1).ClearContents
    End With
Handler:
End Sub
Sub D8_MoveExtend()
' Copies cell over two columns
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Dim CL
    If ActiveSheet.Name = "Cross-x" Or ActiveSheet.Name = "Casebook" Then Exit Sub
    
    With Selection.Areas(1)
        CL = .EntireColumn.Cells(1, 1)
        If CL = "2NR" Or CL = "2AR" Then Exit Sub
        
        If WorksheetFunction.CountA(.Offset(, 2)) Then _
            If MsgBox("Are you sure you want to copy over a cell with text?", 4 + 0) <> 6 Then Exit Sub
        
        .Copy
        ActiveSheet.Paste .Columns(1).Offset(, 2 - (CL = "1AC" Or CL = "1NC"))
        
        .Offset(, 2 - (CL = "1AC" Or CL = "1NC")).Font.Italic = True
        
    End With
    FlowRefresh
End Sub
Sub D8_MoveLeft()
' Copies cells left a collumn
    On Error Resume Next
    Application.ScreenUpdating = False
    
    With Selection.Areas(1)
        If .Columns(1).EntireColumn.Address = "$A:$A" Then Exit Sub
        
        If WorksheetFunction.CountA(.Columns(1).Offset(, -1 + (.Columns(1).EntireColumn.Address = "$C:$C"))) Then _
            If MsgBox("Are you sure you want to copy over a cell with text?", 4 + 0) <> 6 Then Exit Sub
        
        .Cut
        .Columns(1).Offset(, -1 + (.Columns(1).EntireColumn.Address = "$C:$C")).Select
        ActiveSheet.Paste
    End With
End Sub
Sub D8_MoveUp()
' Moves cells up
    Application.ScreenUpdating = False
    On Error GoTo Handler
    
    With Selection.Areas(1)
        If .Rows(1).EntireRow.Address = "$1:$1" Then Exit Sub
        
        If WorksheetFunction.CountA(.Rows(1).Offset(-1)) Then _
            If MsgBox("Are you sure you want to copy over a cell with text?", 4 + 0) <> 6 Then Exit Sub
                
        If .Columns(1).EntireColumn.Address = "$C:$C" Then .Cells(1, 1).Offset(, -1).ClearContents
              
        .Copy
        .Cells(1, 1).Offset(-1).Select
        ActiveSheet.Paste
                
        .Rows(.Rows.Count).ClearContents
    End With
    
Handler:
End Sub
Sub D8_Number()
' Inserts next 2AC/1NC number
    
    On Error Resume Next
    If ActiveSheet.Name = "Cross-x" Or ActiveSheet.Name = "Casebook" Then Exit Sub
    
    With Selection.Cells(1, 1)
        If .EntireColumn.Address = "$C:$C" Then
            .Offset(2, -1).FormulaR1C1 = "=MAX(R2C:R[-1]C)+1"
            .Offset(2, -1).Font.Name = "Arial Narrow"
            .Offset(2).Select
            FlowRefresh
            LineRefresh
        Else
            For x = 1 To 20
                If .Offset(x).Borders(xlEdgeTop).LineStyle = xlDash Then
                    .Offset(x).Select
                    Exit Sub
                End If
            Next x
        End If
    End With
End Sub
Sub D8_Row()
' Inserts row

    On Error Resume Next
    Selection.Rows(Selection.Rows.Count).Offset(1).EntireRow.Insert xlDown, 1
End Sub
Sub D8_RowDelete()
' Erases selected rows
        
    On Error Resume Next
    Dim TextFound As Boolean, rC, aPrompt
    rC = Selection.Rows.Count
    
    For R = 1 To rC
        If WorksheetFunction.CountA(Selection.Rows(R).EntireRow) Then TextFound = True
    Next R
    
    aPrompt = "Are you sure you want to delete " & rC & " rows with text?"
    If rC = 1 Then aPrompt = "Are you sure you want to delete a row with text?"
    
    If TextFound Then If MsgBox(aPrompt, 4 + 0) <> 6 Then Exit Sub

    Selection.EntireRow.Delete
    Selection.Rows(1).Select
    FlowRefresh
End Sub
Sub D8_RowOverview()
' Inserts three rows at the top

    On Error Resume Next
    For x = 1 To 3
        Application.ScreenUpdating = False
        Rows(2).Insert xlDown, 1
    Next x
    ActiveCell.EntireColumn.Rows(2).Select
End Sub
Sub D8_Save(control As IRibbonControl, ByRef cancelDefault)
Attribute D8_Save.VB_ProcData.VB_Invoke_Func = "s\n14"
' This overrides the default command
' Saves by tournament name
    
   ' On Error GoTo Handler
    If ActiveWorkbook.Path > "" Then
        ActiveWorkbook.Save
    Else
        Dim SaveName, fileSaveName, FPath
        
        Dim wd As New Word.Application
            FPath = wd.System.PrivateProfileString(Application.TemplatesPath & _
                "\D8.ini", "Flow", "FPath")
        Set wd = Nothing
        ActiveWorkbook.Activate
        
        SaveName = Trim(Sheets("Casebook").Range("B2"))
        If SaveName = "" Then SaveName = "Debate Flow"
        
        fileSaveName = Application.GetSaveAsFilename(FPath & SaveName & " " & _
            Month(Now) & "-" & Day(Now), "Flow WITH Macros (*.xlsm), *.xlsm,Flow WITHOUT Macros, *.xlsx", 1, "Save Flow", "Save Flow")
    
        Application.DisplayAlerts = False
        If fileSaveName <> False Then ActiveWorkbook.SaveAs _
            fileSaveName, 52 + (Right(fileSaveName, 4) = "xlsx")
   
    End If
    Exit Sub
Handler:
    Application.Dialogs(5).Show
End Sub
Sub D8_Speech()
' Send selection to Speech

    On Error Resume Next
    Dim rW, rH, oWord As Object, IsOpen As Boolean, Side, Sp, D
    

    rW = 0.45 ' set percentage of screen width that the flow should occupy (default: 0.45)
    rH = 0.97 ' set percentage of screen height that both windows should occupy (default: 0.97)
    
    On Error GoTo Handler
    
    Set oWord = GetObject(, "Word.Application")
    
    If oWord.Documents.Count = 0 Then
        oWord.Quit
        Exit Sub
    End If
    
    'try to find speech - if not found use top document
    For D = oWord.Documents.Count To 1
        If InStr(LCase(oWord.Documents(D).Name), "speech") Then _
            Sp = oWord.Documents(D).Name
    Next D
    If Sp = "" Then Sp = oWord.Documents(1).Name
    
    SideColor = ActiveCell.Range("A1").Font.Color
    With Application
        .WindowState = xlNormal
        .Width = 0.75 * Screen(0) * rW
        .Height = 0.75 * Screen(1) * rH
        .Left = 0
        .Top = 0
    End With
    With oWord
        .WindowState = wdWindowStateNormal
        .Width = 0.75 * Screen(0) * (1 - rW)
        .Height = 0.75 * Screen(1) * rH
        .Left = 0.75 * Screen(0) * rW
        .Top = 0
    End With
    
    Selection.Copy
    oWord.Activate
    oWord.Documents(Sp).Activate
   
    oWord.Run "FlowReceive", SideColor
    
    Set oWord = Nothing
    Exit Sub
Handler: MsgBox "Debate Synergy is not open in Microsoft Word. " & _
    "Try running this macro again or restarting both Excel and Word.", vbApplicationModal
End Sub
Sub D8_Star()
' Highlights cell
    
    On Error Resume Next
    With Selection.Interior
        If .Pattern = xlNone Then
            Selection.Font.Bold = True
            .Pattern = xlSolid
            
            If Selection.Font.Color = Blue Or Selection.Font.Color = Blue3 Then .Color = 16763080
            If Selection.Font.Color = Red Or Selection.Font.Color = Red3 Then .Color = 13158655
        Else
            Selection.Font.Bold = False
            .Pattern = xlNone
        End If
    End With
End Sub




