VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim Disp, wd As New Word.Application, LastLine As Boolean

Private Function D8x(Setting As String) As String
    On Error Resume Next
    D8x = wd.System.PrivateProfileString(Application.TemplatesPath & "\D8.ini", "Flow", Setting)
End Function
Private Sub D8s(Setting As String, Data As Variant)
    On Error Resume Next
    wd.System.PrivateProfileString(Application.TemplatesPath & "\D8.ini", "Flow", Setting) = Data
End Sub


Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
On Error Resume Next

With Target.Areas(1)
    
      
    'block broken formulas
    If .HasFormula And .Text = "#NAME?" Then Target = "'" & .Formula
    
    'refresh on cell delete
    If .Text = "" Then
        If Err Then FlowRefresh
        FlowRefresh
        Application.ScreenUpdating = True
    End If
    
    'create abc lists
    If D8x("ABC") = "True" Then
        Dim Num
        
        If .Rows(1).EntireRow.Address <> "$1:$1" And .Text > "" Then
            If Len(.Offset(-1).Text) Then Num = Chr(Asc(Left(.Offset(-1).Text, 1)) + 1)
        
            If Left(.Text, 2) = "=-" Or Left(.Text, 1) = "-" Then
                If Mid(.Offset(-1).Text, 2, 1) <> ")" Then Num = "a"
                Target = Num & ") " & Mid(.Text, 3 + (Left(.Text, 1) = "-"))
            End If
            
            If Mid(.Offset(-1).Text, 2, 1) = ")" And Mid(.Text, 2, 1) <> ")" Then Target = Num & ") " & .Text
        End If
    End If

    'auto-replace shorthand values
    Dim fxU, fxB, fxI, fAll, fS, fR, fF, fT, WordSplit, fso As New FileSystemObject, abr As TextStream, SPath
    
    SPath = Application.TemplatesPath & "\FlowShorthand.txt"
        Err = False
        Dir (SPath)
        If Err Then SPath = "C:\FlowShorthand.txt"
    
    If Dir(SPath) = "" Then
        Set abr = fso.CreateTextFile(SPath, True)
            
        'reset to these default values
        abr.Write Replace( _
        "adv=advantage,alt=alternative,ans=answer,arg=argument,args=arguments,b/c=because,b4=before," & _
        "biod=biodiversity,bpo=best policy option,c/a=cross-apply,c/i=counter-interpretation,calc=calculus," & _
        "*cd=concede,chks=checks,ci=counter-interpretation,comp=competitiveness," & _
        "condo=conditionality,cp=counterplan,cps=counterplans,cts=courts,cx=cross-x,d8=debate," & _
        "da=disad,das=disads,def=definition,defo=deforestation,demo=democracy,deo=deontology,dispo=dispositionality," & _
        "dtr=deterrence,econ=economy,edu=education,env=environment,enviro=environment,ev=evidence,eval=evaluate," & _
        "*ex=extend,exn=extinction,*ext=extend,fair=fairness,fw=framework,fx=effects,g=good,geno=genocide,gov=government," & _
        "gr=ground,gvmt=government,gw=global warming,heg=hegemony,i/l=internal link,!i=impact,il=internal link,ilaw=international law," & _
        "indopak=india-pakistan,inev=inevitable,inf=infinite,inh=inherency,*int=intrinsic,interp=interpretation,intl=international," & _
        "k/t=key to,k=kritik,k2=key to,*kl=kills,l=link,lbl=line-by-line,lit=literature,lm=limits,mag=magnitude,mideast=middle east," & _
        "mt=moving target,n/u=non-unique,nb=net benefit,nm=normal means,noko=north korea,nu=non-unique,nw=nuclear war,o/v=overview," & _
        "*o/w=outweighs,o=offense,ov=overview,overpop=overpopulation,*ow=outweighs,!p=perm,pa=potential abuse,pc=political capital," & _
        "pol cap=political capital,polcap=political capital,pot=potential,pov=poverty,pp=presidential powers,ppl=people,pq=prerequisite," & _
        "pred=predictability,prereq=prerequisite,prob=probability,ptx=politics,quals=qualifications,*rc=root cause,rev=revolution," & _
        "rez=resolution,rt=right,rtp=reasons to prefer,rw=real world,*s=solves,*sev=severance,soc=social,solv=solvency,sopo=soft power," & _
        "spec=specify,sq=status quo,std=standards,strat=strategy,sv=solvency,*t/=turn,t=topicality,t/o=takeout,tf=timeframe,tix=politics,tspec=topic-specific," & _
        "u=uniqueness,util=utilitarianism,_!v=voter,!v2l=value to life,_!vi=voting issue,vio=violation,w/=with,xt=extra topicality", ",", vbCrLf)
    End If
    
    Set abr = fso.OpenTextFile(SPath)
    fAll = Split(abr.ReadAll, vbCrLf)
    
    For x = 0 To UBound(fAll)
        fxU = 0: fxB = 0: fxI = 0
        
        
        If Target.Text > "" Then fT = LCase(Target)
            
        For Y = 0 To 3
            fT = Replace(fT, Array(".", ",", "=", "-")(Y), " ")
        Next Y
    
        If InStr(fAll(x), "=") Then fF = Left(fAll(x), InStr(fAll(x), "=") - 1)
        
            fxU = InStr(Left(fF, 3), "_")
            fxI = InStr(Left(fF, 3), "*")
            fxB = InStr(Left(fF, 3), "!")
            fF = Replace(Left(fF, 3), "_", "") & Mid(fF, 4)
            fF = Replace(Left(fF, 3), "*", "") & Mid(fF, 4)
            fF = Replace(Left(fF, 3), "!", "") & Mid(fF, 4)
        
        fS = InStr(" " & fT & " ", " " & fF & " ")
        
        If fS Then
            fR = Mid(fAll(x), InStr(fAll(x), "=") + 1)
        
            If fR <> "  " Then .Characters(fS, Len(fF)).Text = fR
            
            fS = InStr(fS, Target, fR)
            If fxU Then .Characters(fS, Len(fR)).Font.Underline = 2
            If fxI Then .Characters(fS, Len(fR)).Font.Italic = True
            If fxB Then .Characters(fS, Len(fR)).Font.Bold = True
            
        End If
    Next x
    
    
'DO EVERYTHING BELOW ONLY FOR FLOW SHEETS
    If sh.Name = "Cross-X" Or sh.Name = "Casebook" Or _
        .Columns(1).EntireColumn.Address = "$B:$B" Then Exit Sub
        
    'change sheet name based on entry in cell A4
    If D8x("FlowTitle") = "True" Then
        If .Address = "$A$4" And .Text > "" Then
            sh.Name = .Text
           Select Case ActiveSheet.Tab.Color
                Case Red, Red3
                    For x = 1 To 8
                        If Sheets("Casebook").Range("C" & 9 + x) = "" Then Exit For
                    Next
                    If Sheets("Casebook").Range("C" & 8 + x) <> .Text Then _
                        Sheets("Casebook").Range("C" & 9 + x) = .Text
                Case Blue, Blue3
                    For x = 1 To 8
                        If Sheets("Casebook").Range("B" & 9 + x) = "" Then Exit For
                    Next
                    If Sheets("Casebook").Range("B" & 8 + x) <> .Text Then _
                        Sheets("Casebook").Range("B" & 9 + x) = .Text
            End Select
        End If
    End If
    
    'indent cell containing author
    If D8x("Authors") = "True" Then
        .IndentLevel = -(Val(StrReverse(.Text)) > 0)
    End If
    
    'auto-star voting issues
    If D8x("Voters") = "True" Then
        If (InStr(" " & .Text & " ", " voter ") And InStr(" " & .Text & " ", " no voter ") = 0) Or _
            (InStr(" " & .Text & " ", " vi ") And InStr(" " & .Text & " ", " no vi ") = 0) Or _
            (InStr(" " & .Text & " ", " voting issue ") And InStr(" " & .Text & " ", " no voting issue ") = 0) Or _
            InStr(" " & .Text & " ", " conditionality ") Or _
            InStr(" " & .Text & " ", " dispositionality ") Then
                
                .Font.Bold = True
                .Interior.Pattern = xlSolid
                If .Font.Color = Blue Or .Font.Color = Blue3 Then .Interior.Color = 16763080
                If .Font.Color = Red Or .Font.Color = Red3 Then .Interior.Color = 13158655
        End If
    End If
    
    'auto-number when skipped two lines and below #1
    If D8x("SkipRows") = "True" Then
        Err = False
        If .Rows(1).EntireRow.Address <> "$1:$1" And _
            .Rows(1).EntireRow.Address <> "$2:$2" And _
            .Columns(1).EntireColumn.Address = "$C:$C" Then
            With .Cells(1, 1)
                If (.Text > "" And .Offset(-1) & .Offset(-2) & _
                    .Offset(-1, -1) = "" And Not Val(.Offset(-2, -1))) _
                    And _
                    (WorksheetFunction.Max(Range("B2:B" & _
                        (Val(Mid(.Rows(1).EntireRow.Address, 2)) - 1))) _
                        Or WorksheetFunction.Max(Range("B:B")) = 0) Then

                    If Err = False Then
                        .Offset(, -1).Font.Name = "Arial Narrow"
                        .Offset(, -1).FormulaR1C1 = "=MAX(R2C:R[-1]C)+1"
                    End If
                    FlowRefresh
                End If
            End With
        End If
    End If
    
    LineRefresh
    
End With
End Sub
Private Sub Workbook_SheetActivate(ByVal sh As Object)
'display sheet name in center of screen on sheet switch

    
    If InStr(sh.Name, "case") Then GoTo Handler
    On Error GoTo Handler
    
    Set Disp = ActiveSheet.Shapes.AddTextbox(1, (Application.Width) / _
        (2 - (Application.Version = 11)), (Application.Height) / 4, 100, 70)
    
    With Disp.TextFrame
        .AutoSize = True
        .HorizontalAlignment = xlHAlignCenter
        .Characters.Text = sh.Name
        With .Characters.Font
        .Size = 48
            Select Case ActiveSheet.Tab.ColorIndex
                Case 10, 3: .Color = Red
                Case 12, 5: .Color = Blue
                Case 57, 50: .Color = 6723891
            End Select
        End With
    End With
    
    Application.OnTime Now + TimeValue("00:00:01"), "ThisWorkbook.clearDisp"

Handler:
End Sub
Private Sub clearDisp()
    On Error Resume Next
    Disp.Delete
End Sub
Private Sub Workbook_SheetDeactivate(ByVal sh As Object)
    clearDisp
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, cancel As Boolean)
    Application.DisplayAlerts = False
End Sub

Private Sub Workbook_NewSheet(ByVal sh As Object)
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    D8_FlowAdd
End Sub


