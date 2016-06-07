Attribute VB_Name = "d8_main"
' Debate Synergy 1.5
' Copyright © 2010 Alex Gulakov
' Contact alexgulakov@gmail.com
'
' Debate Synergy is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License 3.0 as published by
' the Free Software Foundation.
'
' Debate Synergy is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License 3 for more details.
'
' You should have received a copy of the GNU General Public License 3
' along with Debate Synergy. If not, see <http://www.gnu.org/licenses/gpl-3.0.txt>.


' This module contains core APIs and functions.
' This module must be imported to your project
' for most of the Debate Synergy macros to work.
' You also need to enable the same references in
' Tools > References as those used by this project.
'
' If copying a Debate Synergy macro to your project,
' include appropriate attribution to Alex Gulakov
' under the terms of the GNU General Public License 3.0.

#If Win64 Then
   'open files
    Public Declare PtrSafe Function ShellExecuteA Lib "shell32.dll" (ByVal HWND As LongPtr, ByVal Oper As String, ByVal FName As String, ByVal Par As String, ByVal Dir As String, ByVal nShowCmd As LongPtr) As LongPtr
    
    'url paste
    Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal HWND As Long, procID As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" (ByVal Class As String, ByVal winName As String) As Long
    Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal HWND As Long, ByVal Cmd As Long) As Long
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal HWND As LongPtr) As LongPtr
    Public Declare PtrSafe Sub keybd_event Lib "user32" (ByVal keyC As Byte, ByVal ext As Byte, ByVal UpDown As LongPtr, ByVal other As LongPtr)
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)
    
    'usb save
    Public Declare PtrSafe Function GetLogicalDriveStringsA Lib "kernel32" (ByVal BuffL As LongPtr, ByVal buffS As String) As LongPtr
    Public Declare PtrSafe Function GetDriveTypeA Lib "kernel32" (ByVal dr As String) As LongPtr
    
    'timer
    Public Declare PtrSafe Function Beep Lib "kernel32" (ByVal fq As LongPtr, ByVal ln As LongPtr) As LongPtr
    Public Declare PtrSafe Function sndPlaySoundA Lib "winmm" (ByVal fl As String, ByVal sync As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal HWND As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal wFlags As LongPtr) As LongPtr
    
    'cursor
    Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal Cur As LongPtr) As LongPtr
    
    'window resize
    Public Declare PtrSafe Function Screen Lib "user32" Alias "GetSystemMetrics" (ByVal r As LongPtr) As LongPtr
    
    'detect pressed keys
    Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal testK As LongPtr) As Integer
#Else
    'open files
    Public Declare Function ShellExecuteA Lib "shell32.dll" (ByVal HWND As Long, ByVal Oper As String, ByVal FName As String, ByVal Par As String, ByVal Dir As String, ByVal nShowCmd As Long) As Long
    
    'url paste
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal HWND As Long, procID As Long) As Long
    Public Declare Function FindWindowA Lib "user32" (ByVal Class As String, ByVal winName As String) As Long
    Public Declare Function GetWindow Lib "user32" (ByVal HWND As Long, ByVal Cmd As Long) As Long
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal HWND As Long) As Long
    Public Declare Sub keybd_event Lib "user32" (ByVal keyC As Byte, ByVal ext As Byte, ByVal UpDown As Long, ByVal other As Long)
    Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
    
    'usb save
    Public Declare Function GetLogicalDriveStringsA Lib "kernel32" (ByVal BuffL As Long, ByVal buffS As String) As Long
    Public Declare Function GetDriveTypeA Lib "kernel32" (ByVal dr As String) As Long
    
    'timer
    Public Declare Function Beep Lib "kernel32" (ByVal fq As Long, ByVal ln As Long) As Long
    Public Declare Function sndPlaySoundA Lib "winmm" (ByVal fl As String, ByVal sync As Long) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal HWND As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
    'cursor
    Public Declare Function SetCursor Lib "user32" (ByVal Cur As Long) As Long
    
    'window resize
    Public Declare Function Screen Lib "user32" Alias "GetSystemMetrics" (ByVal r As Long) As Long
    
    'detect pressed keys
    Public Declare Function GetKeyState Lib "user32" (ByVal testK As Long) As Integer
    
#End If



'read/save private settings
Public Function D8x(Setting As String, Optional Section As String) As String
    On Error Resume Next
    Dim Data
    If Section = "" Then Section = "Main"
    Data = System.PrivateProfileString(NormalTemplate.Path & "\D8.ini", Section, Setting)
    Select Case Data
        Case "True": D8x = True
        Case "False": D8x = False
        Case Else
        If val(Data) Then Data = val(Data)
        D8x = Data
    End Select
    
End Function
Public Sub D8s(Setting As String, Data As Variant, Optional Section As String)
    On Error Resume Next
    If Section = "" Then Section = "Main"
    System.PrivateProfileString(NormalTemplate.Path & "\D8.ini", Section, Setting) = Data
End Sub

'purge non-alphanumerics and extension
Public Function purge(iStr As String, Optional skipExt As Boolean) As String
    Dim X
    If Not skipExt Then iStr = Left(iStr, Len(iStr) - InStr(StrReverse(iStr), "."))
    For X = 1 To 255
        If (X < 48 Or X > 57) And X <> 32 And X <> 40 And X <> 41 And (X < 65 Or X > 90) _
        And (X < 97 Or X > 122) Then iStr = Replace(iStr, Chr(X), " ")
    Next X
    While InStr(iStr, "  ")
        iStr = Replace(iStr, "  ", " ")
    Wend
    purge = Trim(iStr)
End Function

'purge for file paths
Public Function PurgePath(iStr As String) As String
    Dim X
    For X = 1 To 255
        If (X < 48 Or X > 58) And X <> 32 And X <> 92 _
        And X <> 47 And X <> 45 And (X < 65 Or X > 90) _
        And (X < 97 Or X > 122) Or X = 38 Then iStr = Replace(iStr, Chr(X), " ")
    Next X
    PurgePath = Trim(iStr)
End Function

'set default settings on first run


Public Function FirstRun()

    If InStr(Application.UserName, ",") = 0 Then _
        Application.UserName = Application.UserName & ", Team " & Year(Now)
            
    Options.SaveInterval = 1
    Application.DefaultSaveFormat = "Doc"
    Options.AutoFormatAsYouTypeApplyNumberedLists = False
    
    
    
    If D8x("SpeechFolder") = "" Then
        Dim desk
        desk = Replace(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\", "\\", "\")
        If Mid(desk, 2, 2) <> ":\" Then desk = "C:\"
        D8s "SpeechFolder", desk
    End If
    
    If D8x("EveryProg") = "" Then
        Dim ev
        ev = Replace(Options.DefaultFilePath(wdUserTemplatesPath), "microsoft\templates", "Debate Synergy\Everything.exe")
        If Mid(ev, 2, 2) <> ":\" Then ev = "C:\Program Files\Everything\Everything.exe"
        D8s "everyProg", ev
        
    End If
    
    
    If D8x("EveryPath") = "" Then
        Dim ep
        ep = Replace(Options.DefaultFilePath(wdDocumentsPath), "\\", "\")
        If Mid(ep, 2, 2) <> ":\" Then ep = "C:\"
        D8s "everyPath", ep
    End If
    
    
    If D8x("VTub") = "" Then
        Dim VTub As String, fso As New FileSystemObject
        Set fso = CreateObject("Scripting.FileSystemObject")
            
        VTub = Replace(Options.DefaultFilePath(wdDocumentsPath), "\\", "\")
        If Mid(VTub, 2, 2) <> ":\" Then VTub = "C:\"
        VTub = VTub & "\Virtual Tub\"
        
        If Not fso.FolderExists(VTub) Then MkDir VTub
        
        D8s "VTub", Replace(VTub, "\\", "\")
        
        Set fso = Nothing
    End If
    
    
    
    If D8x("Cite") = "" Then D8s "Cite", "AuthorLast Year – Quals (AuthorFirst, Date, Title, URL)"
    If D8x("CiteWords") = "" Then D8s "CiteWords", 5
    
    If D8x("Small") = "" Then D8s "Small", 8
    If D8x("Continues") = "" Then D8s "Continues", "[CONTINUED]"
    If D8x("RemoveTOC") = "" Then D8s "RemoveTOC", True
    
    If D8x("Header") = "" Then D8s "Header", False
    If D8x("PageCount") = "" Then D8s "PageCount", False
    If D8x("Toolbar") = "" Then D8s "Toolbar", True
    If D8x("Paste") = "" Then D8s "Paste", True
    If D8x("LastEdit") = "" Then D8s "LastEdit", False
    If D8x("startview") = "" Then D8s "startview", True
   
        
    
    If D8x("FPath", "Flow") = "" Then
        Dim flw
        flw = Replace(Options.DefaultFilePath(wdDocumentsPath) & "\Flows\", "\\", "\")
        If Mid(flw, 2, 2) <> ":\" Then flw = "C:\Flows\"
        D8s "FPath", flw, "Flow"
    End If
    
    If D8x("SkipRows", "Flow") = "" Then D8s "SkipRows", True, "Flow"
    If D8x("ABC", "Flow") = "" Then D8s "ABC", True, "Flow"
    If D8x("Voters", "Flow") = "" Then D8s "Voters", True, "Flow"
    If D8x("Authors", "Flow") = "" Then D8s "Authors", True, "Flow"
    If D8x("FlowTitle", "Flow") = "" Then D8s "FlowTitle", True, "Flow"
    
    
    If D8x("x1") = "" Then
        D8s "x1", True
        D8s "x2", True
        D8s "x3", False
        D8s "x4", False
        D8s "x5", True
        D8s "x6", False
        D8s "x7", False
        D8s "x8", True
        D8s "x9", True
        D8s "x10", True
        D8s "x11", True
        D8s "x12", True
        D8s "x13", True
        D8s "x14", True
        D8s "x15", True
        D8s "x16", True
        D8s "x17", True
        D8s "x18", True
        D8s "x19", True
        D8s "x20", True
        D8s "x21", True
        D8s "x22", True
        D8s "x23", False
        D8s "x24", True
        D8s "x25", False
    End If

End Function




