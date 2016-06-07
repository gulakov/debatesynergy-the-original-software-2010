VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Casebook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
    'save on tournament name entry
    If Target.Areas(1).Cells(1, 1).Address = "$B$2" And Trim(Range("B2")) > "" Then
        Dim fileSaveName, FPath, wd As New Word.Application
        FPath = wd.System.PrivateProfileString(Application.TemplatesPath & _
            "\D8.ini", "Flow", "FPath")
        Set wd = Nothing
        ActiveWorkbook.Activate

        fileSaveName = Application.GetSaveAsFilename(FPath & Trim(Range("B2")) & " " & _
            Month(Now) & "-" & Day(Now), "Flow WITH Macros (*.xlsm), *.xlsm,Flow WITHOUT Macros, *.xlsx", 1, "Save Flow", "Save Flow")
    
        Application.DisplayAlerts = False
        If fileSaveName <> False Then ActiveWorkbook.SaveAs _
            fileSaveName, 52 + (Right(fileSaveName, 4) = "xlsx")
    End If
    
    'format prep time
    If Target.Areas(1).Cells(1, 1).Address = "$B$5" And InStr(Range("B5"), ":") = 0 Then
        Select Case Len(Range("B5"))
            Case Is > 2: Range("B5") = "'" & Left(Range("B5"), 1) & ":" & Mid(Range("B5"), 2)
            Case 2: Range("B5") = "'0:" & Range("B5")
            Case 1: Range("B5") = "'" & Range("B5") & ":00"
        End Select
    End If
    
    If Target.Areas(1).Cells(1, 1).Address = "$C$5" And InStr(Range("C5"), ":") = 0 Then
        Select Case Len(Range("C5"))
            Case Is > 2: Range("C5") = "'" & Left(Range("C5"), 1) & ":" & Mid(Range("C5"), 2)
            Case 2: Range("C5") = "'0:" & Range("C5")
            Case 1: Range("C5") = "'" & Range("C5") & ":00"
        End Select
    End If
End Sub

