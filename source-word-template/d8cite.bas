Attribute VB_Name = "d8cite"

Sub D8_CiteMagic()
' Creates cite from Lexis/Lexis Law selection or from a selection with
' author lastname (preferably) IN BOLD, title IN QUOTES, author firstname,
' date, url, and quals. For Lexis Law, the journal abbreviation counts as URL.
     
   
    
    Dim cite: cite = D8x("Cite")
    Dim s: s = Selection
    If Len(s) = 1 Then Exit Sub
    If Len(s) > 400 Then If MsgBox("Processing a large amount of text into a cite may take " & _
        "longer than usual.", vbOKCancel, "CiteMagic") <> vbOK Then Exit Sub
    If Len(s) > 5000 Then Exit Sub
    
    Dim tAuthor, eAuthor, lAuthor, tAuthorMain, AuthorFound As Boolean
    Dim tURL, sURL, eURL
    Dim tTitle, TitleWords
    Dim vDate, tDate, tDate2, tyear, sDate, PgChk, sPgChk
    Dim kQuals, sQuals, eQuals, tQuals, w, p
    
    ' non-numerical dates
    sDate = " jan feb mar apr may jun jul aug sep sept oct nov dec " & _
            "winter spring summer fall " & _
            "january february march april may june july august september october november december "
            
    ' if this is before a number, then it's not a date
    sPgChk = " p pp page pages pg vol no v n accessed $ at am pm"
            
    ' start of a qualifications line
    kQuals = _
        " is was senior associate professor fellow assistant lecturer ceo staff" & _
        " strategist specialist worked directed correspondent president author" & _
        " director prof asst editor analyst degree administrator served" & _
        " member institute economist reporter head heads newspaper deputy " & _
        " advocate colonel officer founder founded visiting journalist former" & _
        " retired expert executive manager doctoral candidate chief" & _
        " contributor student blogger chair chairman major general ambassador " & _
        " phd secretary physicist engineer research office school department " & _
        " writer teacher advisor award winner center commentator" & _
        " rand brookings heritage cato un aei forbes nyt cbo "
    
    
    On Error GoTo CiteOutput

'RECOGNIZE LEXIS NEWS CITES
    If InStr(s, "LENGTH:") And (InStr(s, "BIO:") = 0 Or InStr(s, "SECTION:")) Then

        s = Replace(s, Chr(11), Chr(13))
            
        'get publication, ie "author last name" from first line OR from bolded text
        If Selection.Font.Bold = 9999999 Then
            For w = 1 To Selection.Words.Count
                If Selection.Words(w).Font.Bold = True _
                And Trim(Selection.Words(w)) <> "LENGTH" _
                And Trim(Selection.Words(w)) <> "SECTION" _
                And Trim(Selection.Words(w)) <> "BYLINE" _
                And Trim(Selection.Words(w)) <> ":" Then _
                tAuthorMain = tAuthorMain & Selection.Words(w)
            Next w
            If Len(tAuthorMain) > 2 Then AuthorFound = True
        End If
            
        If AuthorFound = False Then
            s = Replace(s, "Publication Logo", "")
            While Left(s, 1) = vbCr
                s = Mid(s, 2)
            Wend
            tAuthorMain = Trim(Left(s, InStr(5, s, vbCr) - 1))
            If InStr(tAuthorMain, "-") Then tAuthorMain = Trim(Left(tAuthorMain, InStr(tAuthorMain, "-") - 1))
        End If
        
        'get byline, ie "author first name"
        If InStr(s, "BYLINE:") Then tAuthor = Mid(s, InStr(s, "BYLINE:") + 8, _
            InStr(InStr(s, "BYLINE:"), s, vbCr) - InStr(s, "BYLINE:") - 8)
            
        'get title
        tTitle = s
        If InStr(tTitle, "BYLINE:") Then tTitle = Replace(tTitle, Mid(tTitle, InStr(tTitle, "BYLINE:")), "")
        If InStr(tTitle, "SECTION:") Then tTitle = Replace(tTitle, Mid(tTitle, InStr(tTitle, "SECTION:")), "")
        If InStr(tTitle, "LENGTH:") Then tTitle = Replace(tTitle, Mid(tTitle, InStr(tTitle, "LENGTH:")), "")
        tTitle = Right(tTitle, InStr(3, StrReverse(tTitle), vbCr))
        tTitle = Replace(tTitle, Chr(13), "")
        tTitle = "“" & Trim(tTitle) & "”"
         
        'get date
        With Selection
            For w = 1 To .Words.Count
                vDate = val(.Words(w))
                If tyear = "" And vDate > 1970 Then tyear = Trim(.Words(w))
                If tDate = "" And vDate > 0 And vDate < 32 Then tDate = Trim(.Words(w))
                If tDate2 = "" And InStr(sDate, " " & Trim(LCase(.Words(w))) & " ") Then tDate2 = Trim(.Words(w))
                If tyear > "" And tDate > "" And tDate2 > "" Then w = .Words.Count
            Next w
        End With
        tDate = Trim(tDate2 & " " & tDate)
        
        tURL = "lexis"
        GoTo CiteOutput
    End If


'RECOGNIZE LEXIS LAW CITES
    If InStr(s, "LENGTH:") And InStr(s, "SECTION:") = 0 Then
        
        s = Replace(s, Chr(11), Chr(13))
        
        tURL = Right(Replace(s, Mid(s, InStr(s, "LENGTH:") - 2), ""), _
            InStr(3, StrReverse(Replace(s, Mid(s, InStr(s, "LENGTH:")), "")), vbCr) - 3)
        
        
        On Error Resume Next
        If InStr(3, StrReverse(Replace(s, Mid(s, InStr(s, tURL)), "")), vbCr) - 3 > 0 Then
        
            tDate = Right(Replace(s, Mid(s, InStr(s, tURL) - 2), ""), _
                InStr(3, StrReverse(Replace(s, Mid(s, InStr(s, tURL)), "")), vbCr) - 3)
        Else
            tDate = Left(s, InStr(s, vbCr))
        End If
        
        tyear = StrReverse(val(StrReverse(tDate)))
        tDate = Trim(Replace(tDate, tyear, ""))
        
        tAuthor = Mid(s, InStr(s, "NAME:") + 6, _
            InStr(InStr(s, "NAME:") + 6, s, vbCr) - InStr(s, "NAME:") - 6)
        
        On Error Resume Next
        If InStr(tAuthor, "+") - 1 > 0 Then _
        If InStr(StrReverse(Trim(Left(tAuthor, InStr(tAuthor, "+") - 1))), " ") - 1 > 0 Then _
            tAuthorMain = Right(Trim(Left(tAuthor, InStr(tAuthor, "+") - 1)), _
                InStr(StrReverse(Trim(Left(tAuthor, InStr(tAuthor, "+") - 1))), " ") - 1)
       
        If InStr(tAuthor, "*") - 1 > 0 Then _
        If InStr(StrReverse(Trim(Left(tAuthor, InStr(tAuthor, "*") - 1))), " ") - 1 > 0 Then _
            tAuthorMain = Right(Trim(Left(tAuthor, InStr(tAuthor, "*") - 1)), _
                InStr(StrReverse(Trim(Left(tAuthor, InStr(tAuthor, "*") - 1))), " ") - 1)
        
        
        tAuthor = Replace(tAuthor, "*", "")
        tAuthor = Replace(tAuthor, "+", "")
        tAuthor = Replace(tAuthor, "by", "")
        tAuthor = Replace(tAuthor, tAuthorMain, "")
        tAuthor = Trim(Replace(tAuthor, "  ", " "))
       
        tQuals = Mid(s, InStr(s, "BIO:") + 5)
        If InStr(tQuals, ". ") > 5 Then tQuals = Left(tQuals, InStr(tQuals, ". ") - 1)
        If InStr(tQuals, vbCr) > 5 Then tQuals = Left(tQuals, InStr(tQuals, vbCr) - 1)
        tQuals = Trim(Replace(tQuals, "*", ""))
        tQuals = Trim(Replace(tQuals, "+", ""))
        
        tTitle = Mid(s, InStr(InStr(s, "LENGTH:"), s, vbCr) + 2, InStr(InStr(InStr(s, "LENGTH:"), s, vbCr) _
        + 2, s, vbCr) - InStr(InStr(s, "LENGTH:"), s, vbCr))
        
        tTitle = Replace(tTitle, "Article:", "")
        tTitle = Replace(tTitle, "ARTICLE:", "")
        tTitle = Replace(tTitle, vbCr, "")
        tTitle = Replace(tTitle, "*", "")
        tTitle = "“" & Trim(tTitle) & "”"
        
        tURL = tURL & ", Lexis Law"
        GoTo CiteOutput
    End If


'RECOGNIZE REGULAR CITES

'get author if bold
     With Selection
        For w = 1 To .Words.Count
            If .Words(w).Font.Bold = True And Len(.Words(w)) > 2 And val(.Words(w)) = 0 Then
                tAuthorMain = tAuthorMain & .Words(w)
                If w > 1 Then
                    If .Words(w - 1).Font.Bold = False And (.Words(w - 1).Case = 2 Or _
                            .Words(w - 1).Case = 4 Or .Words(w - 1).Case = 1) Then _
                                tAuthor = .Words(w - 1)
        
                    End If
                s = Replace(s, tAuthor, "")
                s = Replace(s, tAuthorMain, "")
            End If
        Next w
        If tAuthorMain > "" Then
            AuthorFound = True
            Selection = s
        End If
    End With
    
'get title if in quotes
    If InStr(s, "“") > 0 And InStr(s, "”") > 0 Then
        tTitle = Mid(s, InStr(s, "“"), InStr(s, "”") - InStr(s, "“") + 1)
        s = Replace(s, tTitle, "")
        Selection = s
    Else
        For Each p In Selection.Paragraphs
            For Each w In p.Range.Words
                If w.Case = wdTitleWord Or w.Case = wdUpperCase Then TitleWords = TitleWords + 1
            Next
            
            If p.Range.Words.Count > 6 And TitleWords / p.Range.Words.Count > 0.6 Then
                tTitle = Replace(p.Range, vbCr, "")
                s = Replace(s, tTitle, "")
                Selection = s
                Exit For
            End If
        Next
    End If
    
'get url
    sURL = InStr(s, "http")
    If sURL = 0 Then sURL = InStr(s, "www")
    If sURL Then
        If eURL = 0 And InStr(sURL, s, ".html") And Mid(s, InStr(sURL, s, ".html") + 5, 1) <> "?" _
            Then eURL = InStr(sURL, s, ".html") + 5
        
        If eURL = 0 And InStr(sURL, s, ".htm") And Mid(s, InStr(sURL, s, ".htm") + 5, 1) <> "?" _
            And Mid(s, InStr(sURL, s, ".htm") + 4, 1) <> "?" Then eURL = InStr(sURL, s, ".htm") + 4
        
        If eURL = 0 And InStr(sURL, s, ".pdf") And Mid(s, InStr(sURL, s, ".pdf") + 4, 1) <> "?" _
            Then eURL = InStr(sURL, s, ".pdf") + 4
    
        If eURL = 0 Then eURL = InStr(sURL, s, vbCrLf)
        If eURL = 0 Then eURL = InStr(sURL, s, vbCr)
        If eURL = 0 Then eURL = InStr(sURL, s, vbLf)
        If eURL = 0 Then eURL = InStr(sURL, s, " ")
        
        tURL = Mid(s, sURL, eURL - sURL)
        s = Replace(s, tURL, "")
        Selection = s
    End If

     
'get date
    With Selection
        For w = 1 To .Words.Count
            If InStr(sDate, " " & Trim(LCase(.Words(w))) & " ") Then
                tDate = tDate & " " & Trim(.Words(w))
                s = Replace(s, .Words(w), "", Count:=1)
            End If
            If "’" = Left(.Words(w), 1) And val(Mid(.Words(w), 2)) > 0 Then
                tyear = .Words(w)
                s = Replace(s, .Words(w), "", Count:=1)
            End If
            If "‘" = .Words(w) And .Words(w) <> .Words.Last Then
                tyear = .Words(w + 1)
                s = Replace(s, .Words(w + 1), "", Count:=1)
            End If
            
            vDate = val(.Words(w))
            
            If (vDate < 32 And vDate > 0) Or vDate > 1990 Or (vDate > 90 And vDate < 100) Then
                PgChk = 0
                If w > 3 And w < .Words.Count - 3 Then
                    PgChk = InStr(sPgChk, " " & Trim(LCase(.Words(w - 1))) & " ") + _
                            InStr(sPgChk, " " & Trim(LCase(.Words(w - 2))) & " ") + _
                            InStr(sPgChk, " " & Trim(LCase(.Words(w + 1))) & " ") + _
                            InStr(sPgChk, " " & Trim(LCase(.Words(w + 2))) & " ") + _
                            InStr(sPgChk, " " & Trim(LCase(.Words(w + 3))) & " ")
                End If
            
                If PgChk = 0 Then
                    If vDate > 1970 Then
                    
                        If tyear > "" Then tDate = tDate & " " & tyear
                        tyear = .Words(w)
                    End If
                    If tyear = "" And (vDate < 12 And vDate > 0) Or (vDate > 90 And vDate < 100) Then tyear = .Words(w)
                    If tyear <> .Words(w) Then tDate = tDate & " " & Trim(.Words(w))
                    s = Replace(s, .Words(w), "", Count:=1)
                End If
            End If
            
        Next w
        Selection = s
        tDate = Mid(tDate, 2)
        If tyear = "" Then tyear = Right(tDate, InStr(StrReverse(tDate), " "))
        If tyear = "" Then
            tyear = tDate
            tDate = ""
        End If
    End With
  
' get author if in format: Bob Smith or Bob Q. Smith
    eAuthor = 11
    With Selection
    
    If InStr(LCase(s), "by") Then 'By line takes precedence
    For w = .Words.Count To 1 Step -1
            
            If LCase(Trim(.Words(w))) = "by" Then
                tAuthor = .Words(w + 1)
                If tAuthorMain = "" Then tAuthorMain = .Words(w + 2)
                If Len(tAuthorMain) < 3 Then tAuthorMain = .Words(w + 4)
                AuthorFound = True
            End If
            If AuthorFound Then Exit For
    Next w
    End If
     
           
        For w = 1 To .Words.Count - 4
            
            If LCase(Trim(.Words(w))) = "by" Then
                tAuthor = .Words(w + 1)
                If tAuthorMain = "" Then tAuthorMain = .Words(w + 2)
                If Len(tAuthorMain) < 3 Then tAuthorMain = .Words(w + 4)
                AuthorFound = True
            End If
            
            If InStr(LCase(s), "by") = 0 And AuthorFound = False And (.Words(w).Case = 2 Or .Words(w).Case = 4 Or _
                .Words(w).Case = 1) And .Words(w + 1).Case = 2 Or .Words(w + 1).Case = 1 Then
            
                lAuthor = Len(.Words(w)) + Len(.Words(w + 1))
                If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, vbCr) - lAuthor
                If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, vbCrLf) - lAuthor
                If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, vbLf) - lAuthor
                If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, "(") - lAuthor
                If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, " - ") - lAuthor
                
                If eAuthor < 10 Then
                    tAuthor = .Words(w)
                    If tAuthorMain = "" Then tAuthorMain = .Words(w + 1)
                    AuthorFound = True
                End If
           End If
                
                
            If InStr(LCase(s), "by") = 0 And AuthorFound = False And (.Words(w).Case = 2 Or .Words(w).Case = 4 Or _
                .Words(w).Case = 1) And Trim(.Words(w + 2)) = "." Then
                lAuthor = Len(.Words(w)) + Len(.Words(w + 1)) + Len(.Words(w + 2)) + Len(.Words(w + 3))
                    
                    If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, vbCr) - lAuthor
                    If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, vbCrLf) - lAuthor
                    If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, vbLf) - lAuthor
                    If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, "(") - lAuthor
                    If eAuthor > 10 Then eAuthor = InStr(InStr(s, .Words(w) & .Words(w + 1)), s, " - ") - lAuthor
                    
                    If eAuthor < 10 Then
                        tAuthor = .Words(w) & .Words(w + 1) & .Words(w + 2)
                        tAuthorMain = .Words(w + 3)
                        AuthorFound = True
                    End If
            End If
            If AuthorFound Then w = .Words.Count
        Next w
        s = Replace(s, tAuthor, "")
        s = Replace(s, tAuthorMain, "")
        Selection = s
    End With
    
'get qualifications
    With Selection
        For w = 1 To .Words.Count
        If w > .Words.Count Then Exit For
        If InStr(kQuals, " " & Trim(LCase(.Words(w))) & " ") Then
            
            sQuals = InStr(s, .Words(w))
            
            If eQuals < 10 Then eQuals = InStr(sQuals, s, "." & vbCrLf)
            If eQuals < 10 Then eQuals = InStr(sQuals, s, "." & vbCr)
            If eQuals < 10 Then eQuals = InStr(sQuals, s, "." & vbLf)
            If eQuals < 10 Then eQuals = InStr(sQuals, s, ", “")
            If eQuals < 10 Then eQuals = InStr(sQuals, s, " (")
            If eQuals < 10 Then eQuals = InStr(sQuals, s, " [")
            If eQuals < 10 Then eQuals = InStr(sQuals, s, ". ")
            If eQuals < 10 Then eQuals = InStr(sQuals, s, ", ")
            If eQuals < 10 Then eQuals = InStr(sQuals, s, vbCrLf)
            If eQuals < 10 Then eQuals = InStr(sQuals, s, vbCr)
            If eQuals < 10 Then eQuals = InStr(sQuals, s, vbLf)
            
            If eQuals > 9 Then tQuals = Mid(s, sQuals, eQuals - sQuals)
            s = Replace(s, tQuals, "")
            Selection = s
            
            w = .Words.Count
        End If
        Next w
    End With
    
' create output

CiteOutput:
    Dim t, tMain

    If val(tyear) > 1970 Then tyear = val(Mid(Trim(tyear), 3))
    
    t = cite
    t = Replace(t, "AuthorLast", tAuthorMain)
    t = Replace(t, "Year", tyear)
    t = Replace(t, "Quals", tQuals)
    t = Replace(t, "AuthorFirst", tAuthor)
    t = Replace(t, "Date", tDate)
    t = Replace(t, "Title", tTitle)
    t = Replace(t, "URL", tURL)

    t = Replace(t, Chr(13), " ")
    t = Replace(t, " ,", ",")
    t = Replace(t, ",,", ",")
    t = Replace(t, ",”,", ",”")
    t = Replace(t, ", ,", ",")
    t = Replace(t, ", ,", ",")
    t = Replace(t, "-  (", "(")
    t = Replace(t, "- (", "(")
    t = Replace(t, "–  (", "(")
    t = Replace(t, "– (", "(")
    t = Replace(t, "(, ", "(")
    t = Replace(t, "  (", " (")
    t = Replace(t, ",, ", ", ")
    t = Replace(t, "(, ", "(")
    t = Replace(t, "  ", " ")
    
    tMain = Left(t, InStr(t, tyear) + Len(tyear) - 1)
    t = Replace(t, tMain, "")
    
    With Selection
        .Font.Bold = False
        .Collapse 1
        .Font.Bold = True
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .TypeText tMain
        Selection = t
        .Font.Bold = False
        .EndOf
        .TypeParagraph
    End With
End Sub

