Attribute VB_Name = "d8casebook"
Sub D8_Casebook()
'Exports aff, neg, and judge data for the round to the central casebook.
    
    On Error Resume Next

    Dim Flow As Object, B As Object, bA As Object, bJ As Object, bN As Object, _
        R As Long, x As Long, rNum, aNum, J As Long, Mess As String, FPath As String
    
    Dim wd As New Word.Application
    FPath = wd.System.PrivateProfileString(Application.TemplatesPath & _
        "\D8.ini", "Flow", "FPath")
    
    Set wd = Nothing
    ActiveWorkbook.Activate
    
    Set Flow = ActiveWorkbook.Sheets("Casebook")
        If Err Then
            MsgBox "Could not find the sheet titled Casebook.", vbCritical
            Exit Sub
        End If
        
    'Output message
    If Flow.Range("B1") > "" Then Mess = "affirmative "
    If Flow.Range("C1") > "" Then Mess = Mess & "negative "
    If Flow.Range("B18") > "" Then Mess = Mess & "judge"
    Mess = Replace(StrReverse(Replace(StrReverse(Trim(Mess)), " ", _
        " dna ", 1, 1)), " negative ", ", negative, ")
   
    If Mess > "" Then
        If MsgBox("This round's data for the " & Mess & _
            " will be saved to the casebook.", _
            vbOKCancel + vbInformation, "Casebook") <> vbOK Then Exit Sub
    Else
        MsgBox "You must enter team or judge names to save that data to the casebook."
        Exit Sub
    End If
    
    
    If Dir(FPath & "Casebook.xls") > "" Then
        Set B = Workbooks.Open(FPath & "Casebook.xls")

    Else 'first time run
        
        Workbooks.Add
        For s = 2 To Sheets.Count
            Sheets(s).Visible = False
        Next s
        
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        
        With Columns("A:A")
            .ColumnWidth = 10
            .VerticalAlignment = xlTop
            .WrapText = False
            .Font.Color = -10053325
            .Font.Bold = True
        End With
        
        With Columns("B:B")
            .ColumnWidth = 6
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlTop
            .WrapText = False
            .Font.Color = -10053325
        End With
        
        With Columns("C:C")
            .ColumnWidth = 100
            .WrapText = True
        End With
        
        Sheets(1).Name = "Affs"
        Sheets(1).Copy After:=Sheets(1)
        Sheets(1).Copy After:=Sheets(1)
        Sheets(2).Name = "Negs"
        Sheets(3).Name = "Judges"
        
        MkDir FPath
        ActiveWorkbook.SaveAs FPath & "Casebook.xls", xlExcel8
        Set B = ActiveWorkbook
    End If
    
    
    Set bA = B.Sheets("Affs")
    Set bN = B.Sheets("Negs")
    Set bJ = B.Sheets("Judges")
   
    
    'AFFIRMATIVE
    If Flow.Range("B1") > "" Then
        'insert blank rows
        rNum = 4
        For x = 1 To 8
            If Flow.Range("B" & 9 + x) > "" Then rNum = rNum + 1
        Next
        For R = 1 To rNum
            bA.Rows(1).Insert
        Next
        'draw a line
        With bA.Range("A1:C1").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = -10053325
            .Weight = xlHairline
        End With
        'transfer data
        bA.Range("A1") = Flow.Range("B1")
        bA.Range("B2") = "Tournament"
        bA.Range("C2") = Flow.Range("B2") & " (" & Date & ")"
        bA.Range("B3") = "Plan"
        bA.Range("C3") = Flow.Range("B8")
        If rNum > 4 Then bA.Range("B4") = "Advantages"
        For x = 0 To 7
            If Flow.Range("B" & 10 + x) > "" Then
                bA.Range("C" & 4 + aNum) = Flow.Range("B" & 10 + x)
                aNum = aNum + 1
            End If
        Next
    End If
    
    
    'NEGATIVE
    If Flow.Range("C1") > "" Then
        'insert blank rows
        rNum = 4
        For x = 1 To 8
            If Flow.Range("C" & 9 + x) > "" Then rNum = rNum + 1
        Next
        For R = 1 To rNum
            bN.Rows(1).Insert
        Next
        'draw a line
        With bN.Range("A1:C1").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = -10053325
            .Weight = xlHairline
        End With
        'transfer data
        bN.Range("A1") = Flow.Range("C1")
        bN.Range("B2") = "Tournament"
        bN.Range("C2") = Flow.Range("B2") & " (" & Date & ")"
        bN.Range("B3") = "2NR"
        bN.Range("C3") = Flow.Range("C8")
        If rNum > 4 Then bN.Range("B4") = "1NC List"
        aNum = 0
        For x = 0 To 7
            If Flow.Range("C" & 10 + x) > "" Then
            
                bN.Range("C" & 4 + aNum) = Flow.Range("C" & 10 + x)
                aNum = aNum + 1
            End If
        Next
    End If
    
    
            
            
    'JUDGES
    For J = 66 To 68
        If Flow.Range(Chr(J) & "18") > "" Then
            'insert blank rows
            rNum = 4
            For x = 20 To 200
                If Flow.Range(Chr(J) & x) > "" Then rNum = rNum + 1
            Next
            For R = 1 To rNum
                bJ.Rows(1).Insert
            Next
            'draw a line
            With bJ.Range("A1:C1").Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = -10053325
                .Weight = xlHairline
            End With
            'transfer data
            bJ.Range("A1") = Flow.Range(Chr(J) & "18")
            bJ.Range("B2") = "Tournament"
            bJ.Range("C2") = Flow.Range("B2") & " (" & Date & ")"
            bJ.Range("B3") = "Decision"
            bJ.Range("C3") = Flow.Range(Chr(J) & "19") & " (Aff: " & _
                Flow.Range("B1") & ", Neg: " & Flow.Range("C1") & ")"
            If rNum > 4 Then bJ.Range("B4") = "Comments"
            aNum = 0
            For x = 20 To 200
                If Flow.Range(Chr(J) & x) > "" Then
                    bJ.Range("C" & 4 + aNum) = Flow.Range(Chr(J) & x)
                    aNum = aNum + 1
                End If
            Next
        End If
    Next J
    
    'save casebook
    bA.Select
    B.Save
End Sub
