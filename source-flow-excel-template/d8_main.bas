Attribute VB_Name = "d8_main"
#If Win64 Then
    Public Declare PtrSafe Function ShellOpen Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As LongPtr) As LongPtr
    Public Declare PtrSafe Function LoadCurson Lib "user32" Alias "LoadCursorA" (ByVal ins As LongPtr, ByVal Name As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal Cur As LongPtr) As LongPtr
    Public Declare PtrSafe Function Screen Lib "user32" Alias "GetSystemMetrics" (ByVal R As LongPtr) As LongPtr
#Else
    Public Declare Function ShellOpen Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
    Public Declare Function LoadCurson Lib "user32" Alias "LoadCursorA" (ByVal ins As Long, ByVal Name As Long) As Long
    Public Declare Function SetCursor Lib "user32" (ByVal Cur As Long) As Long
    Public Declare Function Screen Lib "user32" Alias "GetSystemMetrics" (ByVal R As Long) As Long
#End If
Public Const Red = 204
Public Const Blue = 13369344
Public Const Red3 = 395485
Public Const Blue3 = 13893632
    
Dim wd As Word.Application
Private Sub Auto_Open()
    With Application
        'edit the keyboard shortcuts here
        .OnKey "^{RETURN}", "D8_Number"
        .OnKey "^{DOWN}", "D8_MoveDown"
        .OnKey "^{UP}", "D8_MoveUp"
        .OnKey "^{LEFT}", "D8_MoveLeft"
        .OnKey "^{RIGHT}", "D8_MoveExtend"
        .OnKey "{INSERT}", "D8_Row"
        .OnKey "^{DEL}", "D8_RowDelete"
        .OnKey "^{INSERT}", "D8_RowOverview"
        .OnKey "`", "D8_Speech"
        .OnKey "^g", "D8_Group"
        .OnKey "^{BACKSPACE}", "D8_FlowHide"
        .OnKey "^=", "D8_FlowAdd"
        .OnKey "{F1}", "D8_Star"
        .OnKey "{F2}", "D8_Comment"
        .OnKey "{F12}", "D8_Casebook"
        
        .Caption = "Debate Synergy Flow"
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
        .EnableAutoComplete = False
        .ErrorCheckingOptions.UnlockedFormulaCells = False
        .ErrorCheckingOptions.EmptyCellReferences = False
    End With
    
    If InStr(Right(ActiveWorkbook.Name, 4), "xlt") Then _
        MsgBox "You are currently editing the Debate Synergy flow template. " & _
        "Any changes made here will be reflected in all future flows opened from this template. " & _
        vbCr & vbCr & "To open a flow based on this template, double click on the " & _
        "template file or, to open in this template by default, rename the template file " & _
        "as Book.xltm and place it into folder: " & Application.StartupPath, vbInformation, "Debate Synergy Flow"
    
    
    'first run
    
    If Dir(Application.TemplatesPath & "\D8.ini") = "" Then
        Set wd = New Word.Application
    
        f8s "FPath", Replace(Application.DefaultFilePath & "\Flows\", "\\", "\")
        f8s "SkipRows", True
        f8s "ABC", True
        f8s "Voters", True
        f8s "Authors", True
        f8s "FlowTitle", True
        
        Set wd = Nothing
        ActiveWorkbook.Activate
    End If
   
End Sub
Sub main(control As IRibbonControl)
    Select Case Mid(control.ID, 3, 2)
        
        Case "sp": D8_Speech
        
        Case "st": D8_Star
        Case "cm", "cm2": D8_Comment
        Case "nm": D8_Number
        Case "gp": D8_Group
        
        Case "rw", "rw2": D8_Row
        Case "ro": D8_RowOverview
        Case "rd": D8_RowDelete
        Case "fh", "fh2": D8_FlowHide
        Case "fa": D8_FlowAdd
        Case "mr", "mr2": D8_MoveExtend
        Case "ml": D8_MoveLeft
        Case "mu": D8_MoveUp
        Case "md": D8_MoveDown
        
        Case "cb", "cb2": D8_Casebook
          
        Case "op": Config.Show
    End Select
End Sub
Public Sub f8s(Setting As String, Data As Variant)
    On Error Resume Next
    wd.System.PrivateProfileString(Application.TemplatesPath & "\D8.ini", "Flow", Setting) = Data
End Sub


Public Function LineRefresh()
'end of flow line
    
    Dim R, F, xR
    F = 71 - (ActiveSheet.Tab.ColorIndex = 12)
    
    With ActiveSheet.Range("A2:" & Chr(F) & 500)
        
        For R = .Rows.Count To 1 Step -1
            If WorksheetFunction.CountA(.Rows(R)) Then Exit For
        Next R
        
        For xR = R + 3 To 1 Step -1
            If .Rows(xR).Columns(1).Interior.Pattern = xlPatternLinearGradient Then _
                .Rows(xR).Interior.Pattern = xlNone
        Next xR

        .Range("A" & R + 2 & ":" & Chr(F + 1) & 2000).Borders(xlInsideVertical).LineStyle = xlNone
    
        For T = R + 3 To 1 Step -1
            With Range("A:" & Chr(F + 1)).Rows(T).Borders(xlInsideVertical)
                If .LineStyle = xlNone Then
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                Else
                    Exit For
                End If
            End With
        Next T
        Range("B:C").Borders(xlInsideVertical).LineStyle = xlNone
        
        With .Rows(R + 2).Interior
            .Pattern = xlPatternLinearGradient
            .Gradient.Degree = 270
            With .Gradient.ColorStops
                .Clear
                .Add(0).Color = 8421504
                .Add(1).Color = 16777215
            End With
         End With
     End With
End Function

Public Function FlowRefresh()
    Application.ScreenUpdating = False
    With ActiveSheet
    
        .Rows.AutoFit
        
        If .Name = "Cross-x" Or .Name = "Casebook" Then Exit Function
        
        'refresh red/blue colors
        Range("$A:$A, $D:$D, $F:$F, $H:$H").Font.Color = Red - (.Tab.ColorIndex = 12) * 13369140
        Range("$B:$B, $C:$C, $E:$E, $G:$G").Font.Color = Blue + (.Tab.ColorIndex = 12) * 13369140
        
        'clear lines
        Range("A2:H150").Borders(xlInsideHorizontal).LineStyle = xlNone
        Range("B:C").Borders(xlInsideVertical).LineStyle = xlNone
        
        'draw horizontal lines above every 2ac/1nc number, until grouping line
        For Each R In Range("B3:" & Chr(71 - (ActiveSheet.Tab.ColorIndex = 12)) & 150).Rows
            If Val(R.Columns(1).Text) Then
                With R.Columns(1)
                    If Val(.Text) < WorksheetFunction.Max(.EntireColumn) Then
                        .Offset(1).FormulaR1C1 = "¯"
                        .Offset(1).Font.Name = "Symbol"
                    ElseIf .Offset(1).FormulaR1C1 > "" Then
                        .Offset(1).FormulaR1C1 = ""
                    End If
                End With
                
                For c = 1 To R.Columns.Count
                    If R.Columns(c).Borders(xlEdgeLeft).LineStyle = xlDash And _
                        R.Offset(-1).Columns(c).Borders(xlEdgeLeft).LineStyle = xlDash Then Exit For
                Next c
                                        
                For E = 1 To c - 1
                    R.Columns(E).Borders(xlEdgeTop).LineStyle = xlDash
                Next E
            End If
        Next R
            
    End With
End Function


