VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VTub 
   Caption         =   "Virtual Tub Converter"
   ClientHeight    =   8230.001
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   4230
   OleObjectBlob   =   "VTub.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VTub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blocks As New Collection
        

Private Sub bbrowse_Click()
    Dim sh As New Shell32.Shell
    Set dPick = sh.BrowseForFolder(0, "Name the expando and select where to save.", 0, D8x("VTub"))
    If dPick Is Nothing Then Exit Sub
    
    DocTitle.Value = dPick.Self.Path
    
End Sub

Private Sub bconvert_Click()
    
  '  Application.ScreenUpdating = False
    
    Dim sPath As String, b As New Block, fso As New FileSystemObject
    sPath = DocTitle.Value & "\"
    
    If Not fso.FolderExists(sPath) Then MkDir sPath
    
    For n = 1 To blocks.Count
        Set b = blocks(n)
        b.Save sPath
    Next
    
    ShellExecuteA 0, vbNullString, sPath, vbNullString, vbNullString, 1
    x_doRefresh
    Fresh
    
    End
    
End Sub



Private Sub bdel_Click()
   
    Dim j As Long
    j = 0
    For i = 0 To blist.ListCount - 1
       If blist.Selected(i) Then
            blocks.Remove j + 1
            j = j - 1
       End If
       j = j + 1
    Next
    
    fillList
End Sub

Private Sub bgroup_Click()
    Dim b As Block, f As String, n As String
    
    f = purge(InputBox("Enter a name for this expando", "Expando Name"))
    If f = "" Then Exit Sub
    
    
    For i = 0 To blist.ListCount - 1
       If blist.Selected(i) Then
            Set b = blocks(i + 1)
            n = b.FullName
            
            b.FullName = Left(n, InStrRev(n, "\")) & f & _
                "\" & Mid(n, InStrRev(n, "\") + 1)
        
       End If
    Next
    
    fillList
End Sub

Private Sub blist_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 If KeyCode = 27 Then End
End Sub

Private Sub blist_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim b As New Block
    Set b = blocks(blist.ListIndex + 1)
    
    b.selectRange
    
    Dim n
    n = 0
    
    For i = 0 To blist.ListCount - 1
        If blist.Selected(i) Then n = n + 1
    Next
    
    bmerge.Enabled = n > 1
    

End Sub

Private Sub bmerge_Click()
    
    Dim bMain As Block, b As Block, j As Long
    j = 0
    
    For i = 0 To blist.ListCount - 1
        If blist.Selected(i) Then
            If bMain Is Nothing Then
                Set bMain = blocks(i + 1)
            Else
                Set b = blocks(j + 1)
                bMain.addBlock b
                blocks.Remove j + 1
                j = j - 1
            End If
            
        End If
        j = j + 1
    Next
    
    fillList
End Sub



Private Sub brename_Click()
    
    If blist.ListIndex = -1 Then Exit Sub
    
    Dim b As New Block, n As String
    Set b = blocks(blist.ListIndex + 1)
    
    n = purge(InputBox("Enter new name", "New Name", b.FullName))
    If n = "" Then Exit Sub
    
    b.FullName = n
    
    fillList
    
End Sub

Private Sub Move_SpinDown()
    Dim sel As New Collection, index As Long
    index = blist.ListIndex
    
    For X = blist.ListCount - 1 To 0 Step -1
        If blist.Selected(X) Then
            
        
            If X = blist.ListCount - 1 Then
                blocks.Add blocks(X + 1), , 1
                blocks.Remove (X + 2)
                sel.Add 0
            Else
                blocks.Add blocks(X + 1), , , X + 2
                blocks.Remove (X + 1)
                sel.Add X + 1
            
            End If
        
        End If
    Next
    

    fillList
    
    For Each X In sel
        blist.Selected(X) = True
    Next
    blist.ListIndex = index
    
End Sub

Private Sub Move_SpinUp()
    Dim sel As New Collection, index As Long
    index = blist.ListIndex
    
    For X = 0 To blist.ListCount - 1
        If blist.Selected(X) Then
            
        
            If X = 0 Then
                blocks.Add blocks(X + 1)
                blocks.Remove (X + 1)
                sel.Add blocks.Count - 1
            Else
                blocks.Add blocks(X + 1), , X
                blocks.Remove (X + 2)
                sel.Add X - 1
            
            End If
        
        End If
    Next
    

    fillList
    
    For Each X In sel
        blist.Selected(X) = True
    Next
    blist.ListIndex = index
    
End Sub

    

Private Sub UserForm_Initialize()
    getBlocks
    fillList
    
    DocTitle.Value = D8x("VTub") & purge(ActiveDocument.Name, False)
End Sub


Private Sub getBlocks()

    Dim p As Paragraph, b As Block, t As String
    
    For Each p In ActiveDocument.Paragraphs
       

        If p.Range.ParagraphFormat.OutlineLevel = 1 Then
            t = purge(p.Range.Text)
            If t <> vbCr And t <> "" And t <> " " Then
                Set b = New Block
                b.setHead p, True
                            
                blocks.Add b
            
            End If
        ElseIf Not b Is Nothing Then
            b.addParagraph p
        End If
        
        
    
    Next
    
End Sub

Sub fillList()
    
    Dim b As Block, index As Long
    
  '  index = blist.ListIndex
    
    blist.Clear
    
    
    For i = 1 To blocks.Count
        Set b = blocks(i)
       
        blist.AddItem b.FullName
    Next
    
    If index > -1 And index < blist.ListCount Then
        'blist.Selected(index) = True
       ' blist.ListIndex = index
    End If
End Sub

