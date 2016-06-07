VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Block"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FullName As String
Public includeHead As Boolean

Private Head As Paragraph
Public Bodies As New Collection
    
Public Sub addParagraph(p As Paragraph)
    If Bodies.Count = 0 Then Exit Sub
    
    
    Bodies.Add ActiveDocument.Range(Bodies(1).Start, p.Range.End), , 1
    Bodies.Remove 2
    
End Sub

Public Sub addBlock(b As Block)
    Dim p As Range
    For Each p In b.Bodies
        Bodies.Add p
    Next
    

End Sub


Public Sub setHead(p As Paragraph, doIncludeHeading As Boolean)
    
    Set Head = p
    
    FullName = purge(p.Range.Text)
    Bodies.Add ActiveDocument.Range(p.Range.End, p.Range.End)
    includeHead = doIncludeHeading

End Sub

Public Sub Save(sPath As String)
    
    Dim p As Range, n As String, t As String, fso As New FileSystemObject, _
        num As Long, nPath As String
    
    n = FullName
    While InStr(n, "\")
        t = t + Left(n, InStr(n, "\"))
        If Not fso.FolderExists(sPath & t) Then MkDir sPath & t
        n = Mid(n, InStr(n, "\") + 1)
    Wend
    
    
    num = fso.GetFolder(sPath & Left(FullName, InStrRev(FullName, "\"))).Files.Count + 1
    
    nPath = sPath & Left(FullName, InStrRev(FullName, "\")) & Format(num, "00#") & "_" & n & ".doc"
    
     
    If Dir(nPath) > "" Then
         Documents.Open nPath, False, False, False
         Selection.EndKey
    Else
        Documents.Add
    End If
    
    If includeHead And Len(Head.Range.Text) > 1 Then
        Head.Range.Copy
        Selection.Paste
    End If
    
    For Each p In Bodies
        If Len(p.Text) > 1 Then
            p.Copy
            Selection.Paste
        End If
    Next
    
    
    
    
    
    ActiveDocument.SaveAs nPath, wdFormatDocument, , , False
    ActiveDocument.Close 0
End Sub

Public Sub selectRange()
    Head.Range.Select
End Sub




