Attribute VB_Name = "d8r_exp"
Dim xm

'EXPANDOS
Sub x_fill(control As IRibbonControl, ByRef returnedVal)
     On Error Resume Next

    
    Dim vPath, fso As New FileSystemObject, abr As TextStream, _
        xml As String, VTub As String
    
    
    vPath = NormalTemplate.Path & "\VTub.ini"
    VTub = D8x("VTub")
    
    If Dir(vPath) = "" Then
        Set abr = fso.CreateTextFile(vPath, True)
        xm = 0
        If Not fso.FolderExists(VTub) Then MkDir VTub
        xml = x_getList(VTub, 1)
        abr.Write xml
        abr.Close
        Set abr = Nothing
    Else
        
        xml = fso.OpenTextFile(vPath).ReadAll
        If Err Then x_doRefresh
    End If
    
    
    returnedVal = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" & _
        xml & "<menuSeparator id=""xpsep"" />" & _
       " <button id=""d8expfresh"" image=""d8exprefresh"" label=""Refresh Virtual Tub"" " & _
        " onAction=""x_refresh"" /> " & _
        " <button id=""expCreate"" image=""d8expcreate"" label=""Convert to Expando"" " & _
        " onAction=""x_create"" />  " & _
        " <button id=""d8star"" image=""d8star"" label=""Add Shortcut..."" " & _
        " onAction=""x_shortcut"" />" + X + " </menu>"
    
End Sub

Function x_getList(Path As Variant, depth As Long) As String

    Dim d As String, Attr As Long, nm As String, fd As New Collection, fl As New Collection, ext As String
    
    d = Dir(Path, vbDirectory + vbNormal)
    Do
        Attr = GetAttr(Path & d)
        If Attr = 16 Then
            If d <> "." And d <> ".." And depth < 5 Then fd.Add d
        ElseIf Attr = 32 Then
             If Not InStr(d, "&") > 0 And Len(d) > 3 Then fl.Add d
        End If
        d = Dir()
    Loop Until d = ""
    
    'recursion thru folders
    For Each p In fd
          xm = xm + 1
        
          nm = p
          If InStr(nm, "_") > 1 Then _
             If val(Left(nm, InStr(nm, "_") - 1)) > 0 Then _
                 nm = Mid(nm, InStr(nm, "_") + 1)
    
         nm = purge(nm)
          
         If depth < 2 Then
            img = "d8tub"
         Else
            img = "d8exp"
         End If
          
          
         If nm > "" Then _
            out = out & "<splitButton id=""xf" & xm & """><button tag=""" & Path & p & _
                """ onAction=""x_openfolder"" id=""xb" & xm & """ image=""" & _
                img & """ label=""" & nm & """ /><menu id=""xmf" & xm & """>" & x_getList(Path & _
                p & "\", depth + 1) & "</menu></splitButton>"
            
    Next
    
    For Each p In fl
        c = c + 1
        xm = xm + 1
        d = p
         
        
         'recognize doc/docx/rtf
         ext = Right(d, 4)
         If ext = ".doc" Or ext = "docx" Or ext = ".rtf" Then
             img = "imageMso=""FileNew"""
         Else
             img = "image=""d8star"""
         End If
          
         If InStr(p, "_") > 1 Then _
             If val(Left(d, InStr(d, "_") - 1)) > 0 Then _
                 d = Mid(d, InStr(d, "_") + 1)
    
         d = purge(d, False)
         
         If d > "" Then _
             out = out & "<button tag=""" & Path & p & """ id=""xp" & xm & _
                 """ " & img & " label=""" & d & """ onAction=""x_open""/>"
                 
    Next
    
    'return
    x_getList = out

End Function

Sub x_refresh(control As IRibbonControl)
   
    x_doRefresh
    
End Sub

Function x_doRefresh()
    Dim vPath, fso As New FileSystemObject, _
        abr As TextStream, xml As String, VTub As String
    
    VTub = D8x("VTub")
    vPath = NormalTemplate.Path & "\VTub.ini"
    Set abr = fso.CreateTextFile(vPath, True)
    xm = 0
    
    If Not fso.FolderExists(VTub) Then MkDir VTub
    
    
    abr.Write x_getList(D8x("VTub"), 1)
    abr.Close
    
    Set abr = Nothing
    Set fso = Nothing
    Fresh True
End Function



Sub x_create(control As IRibbonControl)
    
    
    
    VTub.show
    
    Exit Sub
    Application.ScreenUpdating = False
    
    Dim nDoc As Document, nPath, pkname As String, curName As String, _
        fso As New FileSystemObject, doHeader As Boolean, num As Integer, _
        sh As Shell32.Shell, dPick As Shell32.Folder2
        
    num = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    curName = Replace(ActiveDocument.Name, ".docx", "")
    curName = Replace(curName, ".doc", "")
    curName = D8x("VTub") & curName
    
    If ActiveDocument.Path = "" Then
        Set sh = New Shell32.Shell
        Set dPick = sh.BrowseForFolder(0, "Select which expando to save the document to, " & _
            "or create a new expando folder.", 0, D8x("VTub"))
        
        If dPick Is Nothing Then Exit Sub
        curName = dPick.Self.Path
        
    End If
    
    
    
    Select Case MsgBox("This document will be converted to an expando and placed in " & curName & _
        ". All the blocks in this document will be separated into .doc files placed in the folder" & _
        " and represent expando pockets. Feel free to edit the files and to create subfolders." & vbCr & vbCr & _
        "Would you like block headers (which appear in the document map) to be included in the output files?", _
        vbQuestion + vbYesNoCancel, "Expando Coverter")
    
        Case vbYes: doHeader = True
        Case vbNo: doHeader = False
        Case Else: Exit Sub
    End Select
    
    'make folder for these files
    If Not fso.FolderExists(curName) Then MkDir curName
    
    With Selection
        .HomeKey wdStory
        Do
            .Collapse 0
            With .Find
                .ClearFormatting
                
                .Text = ""
                .ParagraphFormat.OutlineLevel = 1
                .Forward = True
                .Execute
                
                If .Found = False Then
                    Selection.StartOf wdStory
                    .Execute
                End If
                .ClearFormatting
            End With
            
            
            pkname = purge(Selection, True)
            
            If doHeader Then
                .MoveEnd wdParagraph
            Else
                .Collapse 0
            End If
            
            While .Paragraphs(.Paragraphs.Count).Range.ParagraphFormat.OutlineLevel <> 1 And _
                .End <> ActiveDocument.Range.End
                    .MoveEnd wdParagraph
            Wend
            
            If .End <> ActiveDocument.Content.End Then .MoveUp wdParagraph, 1, 1
        
            If Len(pkname) > 0 Then
                
                nPath = curName & "\" & Format(num, "00#") & _
                    "_" & pkname & ".doc"
                nPathBefore = curName & "\" & Format(num - 1, "00#") & _
                    "_" & pkname & ".doc"
               
                Selection.Copy
                
                If Dir(nPathBefore) > "" Then
                     nPath = nPathBefore
                     Set nDoc = Documents.Open(nPath, False, False, False)
                     nDoc.Activate
                     Selection.EndKey
                     Selection.Paste
                 Else
                     Set nDoc = Documents.Add
                     nDoc.Content.Paste
                 End If
                
                nDoc.SaveAs nPath, wdFormatDocument
                nDoc.Close 0
                
                num = num + 1
                
            End If
    
        Loop Until .End = ActiveDocument.Content.End
        .HomeKey wdStory
    
    End With
    
    
    'Open the folder
    ShellExecuteA 0, vbNullString, curName, vbNullString, vbNullString, 1
        
    Application.ScreenUpdating = True
    x_doRefresh
    Fresh
    
    Set fso = Nothing
End Sub

Sub x_open(control As IRibbonControl)
    On Error Resume Next
    
    Dim pt As String, ext As String
    pt = control.Tag
    
    Application.ScreenUpdating = False

    
    If GetKeyState(&H11) < 0 Then
        If MsgBox("Would you like to delete the pocket """ & pt & "?", _
            vbApplicationModal + vbYesNo, "Delete Pocket") = vbYes Then
                Kill pt
                x_doRefresh
                Fresh
            End If
        Exit Sub
    End If
    
    'open file if not doc/docx/rtf
    ext = Right(pt, 4)
    If ext <> ".doc" And ext <> "docx" And ext <> ".rtf" Then
        ShellExecuteA 0, vbNullString, pt, vbNullString, vbNullString, 1
        Exit Sub
    End If
    
    If GetKeyState(&H10) < 0 Then
        If Selection.Type > 1 Then
            Selection.Copy
            Documents.Open pt, False, False, False
            Selection.MoveStart wdStory
            Selection.Paste
            ActiveDocument.Save
        Else
            Documents.Open pt, False, False, False
        End If
        Exit Sub
    End If

    If Selection.Type > 1 Then Selection.Collapse 0
    Selection.InsertFile control.Tag
    
    Application.ScreenUpdating = True
    x_doRefresh
    Fresh
End Sub

Sub x_openfolder(control As IRibbonControl)
   ' On Error Resume Next
    
    Dim pt As String, fso As New FileSystemObject
    pt = control.Tag
    
    
    If GetKeyState(&H11) < 0 Then
        If MsgBox("Would you like to delete the expando """ & pt & "?", _
            vbApplicationModal + vbYesNo, "Delete Expando") = vbYes Then
                fso.DeleteFolder pt, True
                x_doRefresh
                Fresh
            End If
        Exit Sub
    End If
    
    
    If GetKeyState(&H10) < 0 Then
        ShellExecuteA 0, vbNullString, pt, vbNullString, vbNullString, 1
        Exit Sub
    End If

    
    If Selection.Type < 2 Then D8_BlockSelect
    
    Dim pkname As String
    pkname = purge(Left(InputBox("Enter title for adding a new pocket with selection.", _
        "Pocket Title", purge(Left(Selection.Paragraphs(1).Range, 50))), 50))
        
    If pkname = "" Then Exit Sub
    
    
    Dim d As Files, max As Long, vl As Long, n As File
    Set d = fso.GetFolder(pt).Files
    For Each n In d
        nm = n.Name
        If InStr(nm, "_") > 1 Then
            vl = val(Left(nm, InStr(nm, "_") - 1))
            If vl > max Then max = vl
        End If
    Next
   
   
    nPath = pt & "\" & Format(max + 1, "00#") & "_" & pkname & ".doc"
                   
    
    Selection.Copy
    
    
    Set nDoc = Documents.Add
    nDoc.Content.Paste
     
    nDoc.SaveAs nPath, wdFormatDocument
    nDoc.Close 0
    
                
     x_doRefresh
    Fresh

End Sub


Sub x_shortcut(control As IRibbonControl)
    On Error Resume Next
    Dim filePath, ws As Object, scut As Object, docName
    Set ws = CreateObject("WScript.Shell")
            
    
    With Application.FileDialog(3)
        .AllowMultiSelect = False
        .ButtonName = "Bookmark"
        .InitialFileName = Options.DefaultFilePath(wdDocumentsPath)
        .InitialView = 2
        .Title = "Bookmark a Favorite Document"
        If .show Then
            filePath = .SelectedItems(1)
            docName = Mid(filePath, InStrRev(filePath, "\") + 1)
           
    
            Set scut = ws.CreateShortcut(D8x("VTub") & docName & ".lnk")
            scut.TargetPath = filePath
            scut.IconLocation = filePath & ", 0"
            scut.WorkingDirectory = D8x("VTub")
            scut.Save

            
        End If
    End With
    
    x_doRefresh
    Fresh
    
    Set ws = Nothing
End Sub


