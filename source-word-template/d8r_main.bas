Attribute VB_Name = "d8r_main"
Dim D8r As IRibbonUI

Public Sub D8Init(ribbon As IRibbonUI)
    Set D8r = ribbon
End Sub
Function Fresh(Optional doAlert As Boolean = False)
    If Not D8r Is Nothing Then D8r.Invalidate
End Function
Sub toolbar03(control As IRibbonControl, ByRef returnedVal)
    If D8x("Toolbar") > "" Then
        returnedVal = D8x("Toolbar")
    Else
        returnedVal = False
    End If
End Sub

'MAIN BUTTONS
Sub main(control As IRibbonControl)
    
    On Error Resume Next
    Select Case Mid(Replace(control.id, val(StrReverse(control.id)), ""), 3)
        Case "sess": D8_SaveSession
        Case "recover": D8_ViewRecovered
        Case "pdf": D8_PDFConvert
        Case "merge": D8_MergeDocs
        Case "options": Config.show
        
        Case "spsave": D8_SaveUSB
        Case "spbr": D8_SpeechMarker
        Case "spauto": D8_FolderAutoOpen
        Case "spread": D8_SpeechRead
        
        Case "rpaste": D8_PasteText
        Case "rpasteurl": D8_PasteURL
        Case "rpastereturns": D8_PasteReturns
        Case "rpastecb": Application.ShowClipboard
        Case "rreturns": D8_RemoveReturns
        Case "rcite": D8_CiteMagic
        Case "fheading": D8_FormatHeading
        Case "fheadingnot": D8_FormatNotHeading
        Case "fhat": D8_FormatHat
        Case "ftoggle": D8_FormatToggle
        Case "fsmall": D8_FormatSmallAll
        Case "fnormal": D8_FormatNormal
        Case "fclear": Selection.ClearFormatting
        Case "fhighlite": D8_FormatHighlight
        Case "fbox": D8_FormatBox
        Case "fsimilar": D8_FormatSimilar
        Case "fsmallmore": D8_FormatSmallAllMore
      
        Case "xauto": D8_FixAutoFormat
        Case "xblnk": D8_FixBlankPages
        Case "xcaps": D8_FixCaps
        Case "xcreq": D8_FixCiteRequest
        Case "xlink": D8_FixNoLinks
        Case "xcont": D8_FixPageContinued
        
        Case "toc": D8_TOC
        Case "pageheader": D8_PageHeader
        Case "rtimer": Application.OnTime Now + TimeValue("00:00:00"), "D8_Timer"
        
        Case "cadd": D8_WarrantAdd
        Case "fresh": Fresh
        Case "fsmallauto": D8_FormatToggleAuto
        
        
        Case Else: MsgBox Replace(control.id, val(StrReverse(control.id)), "")
    End Select
End Sub

Sub sp_main(control As IRibbonControl, id As String, index As Integer)
    Select Case id
        Case "d8spadd2": Fresh: D8_SpeechSend
        Case "d8spblock": D8_BlockSelect: D8_SpeechSend
        Case "d8spsave": D8_SaveUSB
        Case "d8spbr": D8_SpeechMarker
        Case Else: D8_SpeechNew (Mid(id, 3))
    End Select
End Sub

Sub show(control As IRibbonControl, ByRef returnedVal)
     If D8x("x1") = "" Then FirstRun
     
    Select Case control.id
        Case "d8view": returnedVal = D8x("x1")
        Case "d8win": returnedVal = D8x("x2")
        Case "d8spread": returnedVal = D8x("x3")
        Case "winsbs": returnedVal = D8x("x4")
        Case "d8speech": returnedVal = D8x("x5")
        Case "d8spsave2": returnedVal = D8x("x6")
        Case "d8spbr2": returnedVal = D8x("x7")
        Case "d8expmain": returnedVal = D8x("x8")
        Case "d8form": returnedVal = D8x("x9")
        Case "d8pastemain": returnedVal = D8x("x10")
        Case "d8rreturns": returnedVal = D8x("x11")
        Case "d8rcite": returnedVal = D8x("x12")
        Case "d8ftogglemain": returnedVal = D8x("x13")
        Case "d8fnormalmain": returnedVal = D8x("x14")
        Case "d8fheadingmain": returnedVal = D8x("x15")
        Case "d8fhighlitemain": returnedVal = D8x("x16")
        Case "d8use": returnedVal = D8x("x17")
        Case "d8format": returnedVal = D8x("x18")
        Case "d8fixmain": returnedVal = D8x("x19")
        Case "d8fixmainsep": returnedVal = D8x("x19")
        Case "d8fcommain": returnedVal = D8x("x20")
        Case "d8pageheader": returnedVal = D8x("x21")
        Case "d8toc": returnedVal = D8x("x22")
        Case "qualsep", "quallabel", "showonly", "q1", "q2", "q3": returnedVal = D8x("x23")
        Case "fssep", "fsinput", "fslabel": returnedVal = D8x("x24")
        Case "d8options7": returnedVal = D8x("x25")
        Case Else: returnedVal = True
    End Select
    
End Sub

