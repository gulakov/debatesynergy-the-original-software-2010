Public Class TimerWin
    Dim FlashColor = Color.FromArgb(255, 0, 0), _
        BlueColor = Color.FromArgb(151, 194, 238), _
        LightColor = Color.FromArgb(198, 217, 238), _
        NormalColor = BlueColor, _
        FlashInt = 0, ignoreResize = False, prevWidth = 0, d8 = "DebateSynergyTimer", T = DateTime.Now.Second

    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, _
         ByVal lpWindowName As String) As Long
    Private Declare Function GetClassNameA Lib "user32" (ByVal hWnd As Long, _
         ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


    'ticker
    Private Sub Timer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer.Tick

        'check second intervals against system timer
        If T = DateTime.Now.Second Then Exit Sub
        T = DateTime.Now.Second

        Dim Type = Mid(Me.Text, 1, 2), Time = Time2Sec(Disp.Text)

        'terminal alert
        If (Time <= 0 And Not oTU.Checked) Or (Time >= 3600 And oTU.Checked) Then
            If Type = "Co" And c000.Checked Or Type = "Re" And r000.Checked Or _
                Type = "Cr" And x000.Checked Or (Type = "Af" Or Type = "Ne") And p000.Checked Then

                Timer.Stop()
                Start.Image = d8timer.My.Resources.Resources.go1
                If oAB.Checked Then My.Computer.Audio.Play(My.Resources.beep_final, AudioPlayMode.Background)
                If oAV.Checked Then My.Computer.Audio.Play(My.Resources._000, AudioPlayMode.Background)
                If oAF.Checked Then MsgBox("Time up!", MsgBoxStyle.SystemModal + MsgBoxStyle.Exclamation)
            End If

            Select Case Type
                Case "Af", "Ne", "Re" : bR.PerformClick()
                Case "Co" : bX.PerformClick()
                Case "Cr" : bC.PerformClick()
            End Select
            Exit Sub
        End If

        'save prep
        SaveP()

        'alerts
        If Not oTU.Checked And (Type = "Co" And ((c030.Checked And Time = 30) Or _
            (c100.Checked And Time = 60) Or (c130.Checked And Time = 90) Or _
            (c200.Checked And Time = 120) Or (c300.Checked And Time = 180) Or _
            (c400.Checked And Time = 240) Or (c500.Checked And Time = 300) Or _
            (c600.Checked And Time = 360)) Or _
            Type = "Cr" And ((x030.Checked And Time = 30) Or _
            (x100.Checked And Time = 60) Or (x130.Checked And Time = 90) Or _
            (x200.Checked And Time = 120)) Or _
            Type = "Re" And ((r030.Checked And Time = 30) Or _
            (r100.Checked And Time = 60) Or (r130.Checked And Time = 90) Or _
            (r200.Checked And Time = 120) Or (r300.Checked And Time = 180) Or _
            (r400.Checked And Time = 240) Or (r500.Checked And Time = 300)) Or _
            (Type = "Af" Or Type = "Ne") And ((p030.Checked And Time = 30) Or _
            (p100.Checked And Time = 60) Or (p130.Checked And Time = 90) Or _
            (p200.Checked And Time = 120) Or (p300.Checked And Time = 180) Or _
            (p400.Checked And Time = 240) Or (p500.Checked And Time = 300) Or _
            (p600.Checked And Time = 360))) Then

            'beep
            If oAB.Checked Then BeepAlert()

            'voice
            If oAV.Checked Then
                Dim Snd = My.Resources._030
                Select Case Time
                    Case 30 : Snd = My.Resources._030
                    Case 60 : Snd = My.Resources._100
                    Case 90 : Snd = My.Resources._130
                    Case 120 : Snd = My.Resources._200
                    Case 180 : Snd = My.Resources._300
                    Case 240 : Snd = My.Resources._400
                    Case 300 : Snd = My.Resources._500
                    Case 360 : Snd = My.Resources._600
                End Select
                My.Computer.Audio.Play(Snd, AudioPlayMode.Background)
            End If

            'flash
            If oAF.Checked Then
                Disp.BackColor = FlashColor
                FlashInt = 4
            End If
        End If

        'clear flash alert
        If FlashInt > 0 Then
            FlashInt = FlashInt - 1
            If FlashInt = 0 Then Disp.BackColor = NormalColor
        End If

        'time display
        Disp.Text = Sec2Time(Time - 1 - 2 * (oTU.Checked))

        'save prep (again)
        SaveP()
    End Sub

    'common operations
    Function Time2Sec(ByVal iTime As String) As Integer
        Return Val(Mid(iTime, 1, 2)) * 60 + Val(Mid(iTime, InStr(iTime, ":") + 1, 2))
    End Function
    Function Sec2Time(ByVal iSec As Integer) As String
        Dim m As Integer = Math.Floor(iSec / 60)
        Return m & ":" & Format(iSec - m * 60, "0#")
    End Function
    Sub BeepAlert()
        My.Computer.Audio.Play(My.Resources.beep, AudioPlayMode.Background)
    End Sub
    Sub SaveP(Optional ByVal doRev As Boolean = False)
        If doRev Then
            If Disp.ForeColor = AffP.BackColor Then Disp.Text = AffP.Text
            If Disp.ForeColor = NegP.BackColor Then Disp.Text = NegP.Text
        Else
            If Disp.ForeColor = AffP.BackColor Then AffP.Text = Disp.Text
            If Disp.ForeColor = NegP.BackColor Then NegP.Text = Disp.Text
        End If
    End Sub
    Function convertInput() As String
        Dim Tx = Disp.Text, Sec = Val(Strings.Right(Tx, 2)), ln = Len(Tx), ins = InStr(Tx, ":")

        If ins > 0 Then
            Tx = Val(Strings.Left(Tx, ins - 1)) * 60 + Val(Strings.Right(Tx, ln - ins))
        ElseIf ln < 2 Then
            Tx = Val(Tx) * 60
        ElseIf ln = 2 Then
            Tx = Sec
        Else
            Tx = Val(Strings.Left(Tx, ln - 2)) * 60 + Sec
        End If
        Tx = Math.Min(5940, Val(Tx))

        Return Sec2Time(Tx)
    End Function

    'start button
    Private Sub Start_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Start.Click
        Disp.BackColor = NormalColor

        If Timer.Enabled Then
            Timer.Stop()
            Start.Image = d8timer.My.Resources.Resources.go1
        Else
            If oAB.Checked Or oAV.Checked Then BeepAlert()

            'correct time input
            Disp.Text = convertInput()

            T = DateTime.Now.Second
            Timer.Start()
            Start.Image = d8timer.My.Resources.Resources.go2
            If oWS.Checked Then SizeDown()
        End If
    End Sub

    'buttons to set times
    Sub TimeSet(ByVal s, ByVal name)
        If Timer.Enabled Then Start.PerformClick()
        If Disp.Text <> "0:00" Then oBack.Text = "Back to " + Me.Text + " (" + Disp.Text + ")"
        Me.Text = name
        Disp.Text = s.Text
        SizeUp()
        Start.Focus()
        If s.BackColor = LightColor Then
            Disp.ForeColor = Color.Black
        Else
            Disp.ForeColor = s.BackColor
        End If
    End Sub
    Private Sub bX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bX.Click
        TimeSet(sender, "Cross-X")
    End Sub
    Private Sub bC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bC.Click
        TimeSet(sender, "Constructive")
    End Sub
    Private Sub bR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bR.Click
        TimeSet(sender, "Rebuttal")
    End Sub
    Private Sub AffP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AffP.Click
        TimeSet(sender, "Aff Prep")
    End Sub
    Private Sub NegP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NegP.Click
        TimeSet(sender, "Neg Prep")
    End Sub

    'time input
    Private Sub Disp_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Disp.TextChanged

        Dim x As Integer
        For x = 1 To 255
            If x < 48 Or x > 58 And InStr(Disp.Text, Chr(x)) Then Disp.Text = Replace(Disp.Text, Chr(x), "")
        Next x

        If Me.Text = "Aff Prep" Then
            AffP.Text = convertInput()
        ElseIf Me.Text = "Neg Prep" Then
            NegP.Text = convertInput()
        End If

    End Sub
    Private Sub Disp_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Disp.MouseUp
        If Timer.Enabled Then Start.PerformClick()
    End Sub

    
    'options - alerts
    Private Sub oAB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oAB.CheckedChanged
        oAV.Enabled = Not oAB.Checked
        If oAB.Checked Then oAV.Checked = False
    End Sub
    Private Sub oAV_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oAV.CheckedChanged
        oAB.Enabled = Not oAV.Checked
        If oAV.Checked Then oAB.Checked = False
    End Sub

    'options - time
    Private Sub oTU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oTU.Click
        If oTU.Checked And Not Timer.Enabled Then Disp.Text = "0:00"
    End Sub
    Private Sub oBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oBack.Click
        If Timer.Enabled Then Start.PerformClick()

        Dim t = oBack.Text, Type = Mid(t, 9, InStr(t, "(") - 10)
        If Disp.Text <> "0:00" Then oBack.Text = "Back to " + Me.Text + " (" + Disp.Text + ")"
        Disp.Text = Mid(t, InStr(t, "(") + 1, InStr(t, ")") - InStr(t, "(") - 1)
        Me.Text = Type

        Select Case Type
            Case "Aff Prep" : Disp.ForeColor = AffP.BackColor
            Case "Neg Prep" : Disp.ForeColor = NegP.BackColor
            Case Else : Disp.ForeColor = Color.Black
        End Select
    End Sub
    Private Sub oReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oReset.Click
        prof(AffP, "p")
        prof(NegP, "p")
        SaveP(True)
    End Sub

    'options - right click
    Private Sub AffPMinus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AffPMinus.Click
        AffP.Text = Sec2Time(Math.Max(Time2Sec(AffP.Text) - 5, 0))
        SaveP(True)
    End Sub
    Private Sub NegPMinus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NegPMinus.Click
        NegP.Text = Sec2Time(Math.Max(Time2Sec(NegP.Text) - 5, 0))
        SaveP(True)
    End Sub
    Private Sub AffPReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AffPReset.Click
        prof(AffP, "p")
        SaveP(True)
    End Sub
    Private Sub NegPReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NegPReset.Click
        prof(NegP, "p")
        SaveP(True)
    End Sub

    'options - window
    Private Sub oWT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oWT.CheckedChanged
        Me.TopMost = oWT.Checked
        oWS.Enabled = oWT.Checked
        If Not oWT.Checked Then
            oWR.Enabled = False
            oWS.Checked = False
            oWR.Checked = False
        End If
    End Sub
    Private Sub oWS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oWS.CheckedChanged
        oWR.Enabled = oWS.Checked
        If Not oWS.Checked Then oWR.Checked = False
    End Sub
    Private Sub oAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oAbout.Click
        About.Show()
        About.TopMost = oWT.Checked
    End Sub

    'options - switch
    Private Sub LoadProfile()
        prof(AffP, "p")
        prof(NegP, "p")
        prof(bC, "c")
        prof(bX, "x")
        prof(bR, "r")
    End Sub
    Private Sub prof(ByVal obj As Object, ByVal val As String)
        Dim p As String = "p1"
        If oSwitch.Text = "Switch to Primary Profile" Then p = "p2"
        Dim Data As String = GetSetting(d8, "Options", p + val)
        If Data = "" Then Exit Sub
        obj.Text = Data + ":00"
    End Sub
    Private Sub oConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oConfig.Click
        ProfileConfig.Show()
    End Sub
    Private Sub oSwitch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles oSwitch.Click
        If oSwitch.Text.Contains("Secondary") Then
            oSwitch.Text = "Switch to Primary Profile"
        Else
            oSwitch.Text = "Switch to Secondary Profile"
        End If

        LoadProfile()
    End Sub


    'load/unload
    Private Sub op(ByVal obj As Object, ByVal val As String, Optional ByVal spec As Boolean = False)
        Dim Data As String = GetSetting(d8, "Options", val)
        If Data = "" Then Exit Sub
        If spec Then
            obj.text = Data
        Else
            obj.Checked = (Data = "True")
        End If
    End Sub
    Private Sub os(ByVal obj As Object, ByVal val As String, Optional ByVal spec As Boolean = False)
        If spec Then
            SaveSetting(d8, "Options", val, obj)
        Else
            SaveSetting(d8, "Options", val, obj.checked)
        End If
    End Sub
    Private Sub TimerWin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        

        'profile
        op(oSwitch, "Switch", True)
        LoadProfile()

        'alerts
        op(c000, "c000") : op(c030, "c030") : op(c100, "c100") : op(c130, "c130")
        op(c200, "c200") : op(c300, "c300") : op(c400, "c400") : op(c500, "c500") : op(c600, "c600")
        op(r000, "r000") : op(r030, "r030") : op(r100, "r100") : op(r130, "r130")
        op(r200, "r200") : op(r300, "r300") : op(r400, "r400") : op(r500, "r500")
        op(p000, "p000") : op(p030, "p030") : op(p100, "p100") : op(p130, "p130")
        op(p200, "p200") : op(p300, "p300") : op(p400, "p400") : op(p500, "p500") : op(p600, "p600")
        op(x000, "x000") : op(x030, "x030") : op(x100, "x100") : op(x130, "x130") : op(x200, "x200")
        op(oAV, "oAV") : op(oAB, "oAB") : op(oAF, "oAF")

        'options
        op(oTU, "oTU") : op(oWT, "oWT") : op(oWS, "oWS") : op(oWR, "oWR")

        'window location
        Me.Location = New Point(Math.Max(Val(GetSetting(d8, "Options", "wX", 20)), 0), _
                                Math.Max(Val(GetSetting(d8, "Options", "wY", 20)), 0))
        Me.Width = Val(GetSetting(d8, "Options", "wS", 200))

        'saved times
        op(Disp, "Disp", True) : op(AffP, "AffP", True) : op(NegP, "NegP", True)
        op(Me, "Type", True) : op(oBack, "Back", True)

        If Me.Text = "Aff Prep" Then Disp.ForeColor = AffP.BackColor
        If Me.Text = "Neg Prep" Then Disp.ForeColor = NegP.BackColor
    End Sub
    Private Sub TimerWin_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        SaveSetting(d8, "Options", "Switch", oSwitch.Text)

        os(c000, "c000") : os(c030, "c030") : os(c100, "c100") : os(c130, "c130")
        os(c200, "c200") : os(c300, "c300") : os(c400, "c400") : os(c500, "c500") : os(c600, "c600")
        os(r000, "r000") : os(r030, "r030") : os(r100, "r100") : os(r130, "r130")
        os(r200, "r200") : os(r300, "r300") : os(r400, "r400") : os(r500, "r500")
        os(p000, "p000") : os(p030, "p030") : os(p100, "p100") : os(p130, "p130")
        os(p200, "p200") : os(p300, "p300") : os(p400, "p400") : os(p500, "p500") : os(p600, "p600")
        os(x000, "x000") : os(x030, "x030") : os(x100, "x100") : os(x130, "x130") : os(x200, "x200")
        os(oAV, "oAV") : os(oAB, "oAB") : os(oAF, "oAF")

        os(oTU, "oTU") : os(oWT, "oWT") : os(oWS, "oWS") : os(oWR, "oWR")

        os(Me.Location.X, "wX", True) : os(Me.Location.Y, "wY", True) : os(Me.Width, "wS", True)

        os(Disp.Text, "Disp", True) : os(AffP.Text, "AffP", True) : os(NegP.Text, "NegP", True)
        os(Me.Text, "Type", True) : os(oBack.Text, "Back", True)
    End Sub


    'auto shrink window
    Sub SizeUp()
        If Not oWS.Checked Then Exit Sub

        Me.Activate()

        If oWR.Checked And Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable Then Exit Sub

        ignoreResize = True

        If oWR.Checked Then
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Disp.BackColor = BlueColor
            NormalColor = BlueColor
            ControlPanel.Visible = True
            Menubar.Visible = True
        End If

        If prevWidth > 200 Then
            Me.Width = prevWidth
            Me.Height = prevWidth
            prevWidth = 0
        End If

        ignoreResize = False
    End Sub
    Sub SizeDown()
        If Not Timer.Enabled Or Not oWS.Checked Then Exit Sub


        ignoreResize = True
        prevWidth = Me.Width

        If oWR.Checked Then
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            Disp.BackColor = Color.DimGray
            NormalColor = Color.DimGray

            ControlPanel.Visible = False
            Menubar.Visible = False
        Else
            Me.Height = Disp.Height * 0.75 - 35 * Not oWR.Checked

            ignoreResize = False
        End If

    End Sub
    Private Sub Disp_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Disp.MouseMove
        SizeUp()
    End Sub
    Private Sub TimerWin_Deactivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Deactivate
        SizeDown()
    End Sub

    Private Sub AffP_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles AffP.MouseMove
        Me.Activate()
    End Sub
    Private Sub NegP_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NegP.MouseMove
        Me.Activate()
    End Sub
    Private Sub Start_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Start.MouseMove
        Me.Activate()
    End Sub
    Private Sub bX_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles bX.MouseMove
        Me.Activate()
    End Sub
    Private Sub bC_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles bC.MouseMove
        Me.Activate()
    End Sub
    Private Sub bR_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles bR.MouseMove
        Me.Activate()
    End Sub
    'resize window
    Private Sub TimerWin_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize

        If oWR.Checked And Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None Then Exit Sub

        If oWR.Checked Then
            ignoreResize = (Me.WindowState = FormWindowState.Minimized)
        End If

        If (Me.WindowState = FormWindowState.Minimized) Then ignoreResize = True
        If ignoreResize Then Exit Sub

        Me.Height = Me.Width
        Me.Width = Me.Height
        If Me.Height < 200 Then Exit Sub


        Dim x = Me.ClientSize.Height / 172, y = Me.ClientSize.Width / 194

        If y = 0 Or x = 0 Then Exit Sub

        Disp.Left = -2 * y
        Disp.Top = -11 * x
        Disp.Height = 81 * x
        Disp.Width = 196 * y
        Disp.Font = New Font(Disp.Font.Name, 50 * y, FontStyle.Bold, Disp.Font.Unit, Disp.Font.GdiCharSet)
        ControlPanel.Top = 62 * x
        ControlPanel.Height = 88 * x
        ControlPanel.Width = 194 * y
        Menubar.Top = 150 * x
        Menubar.Height = 22 * x
        Menubar.Width = 194 * y
        Menubar.Padding = New Padding(20 * y, 2 * x, 0, 0)
        Menubar.Font = New Font(Menubar.Font.Name, 8 * x, 0, Menubar.Font.Unit, Menubar.Font.GdiCharSet)

        AffP.Top = 2 * x
        AffP.Left = 2 * y
        AffP.Height = 41 * x
        AffP.Width = 62 * y
        AffP.Font = New Font(AffP.Font.Name, 13 * y, FontStyle.Bold, AffP.Font.Unit, AffP.Font.GdiCharSet)
        NegP.Top = 45 * x
        NegP.Left = 2 * y
        NegP.Height = 41 * x
        NegP.Width = 62 * y
        NegP.Font = New Font(NegP.Font.Name, 13 * y, FontStyle.Bold, NegP.Font.Unit, NegP.Font.GdiCharSet)
        Start.Top = 2 * x
        Start.Left = 66 * y
        Start.Height = 84 * x
        Start.Width = 62 * y


        bX.Top = 59 * x
        bX.Left = 130 * y
        bX.Height = 27 * x
        bX.Width = 62 * y
        bX.Font = New Font(bX.Font.Name, 10 * y, FontStyle.Bold, bX.Font.Unit, bX.Font.GdiCharSet)
        bC.Top = 2 * x
        bC.Left = 130 * y
        bC.Height = 27 * x
        bC.Width = 62 * y
        bC.Font = New Font(bC.Font.Name, 10 * y, FontStyle.Bold, bC.Font.Unit, bC.Font.GdiCharSet)
        bR.Top = 31 * x
        bR.Left = 130 * y
        bR.Height = 26 * x
        bR.Width = 62 * y
        bR.Font = New Font(bR.Font.Name, 10 * y, FontStyle.Bold, bR.Font.Unit, bR.Font.GdiCharSet)
    End Sub
End Class