VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebBrowser Tricks"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   5
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editing Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   1695
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Cut"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Copy"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Paste"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Select All"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "The ""about: "" Method"
      Height          =   855
      Left            =   1920
      TabIndex        =   16
      Top             =   720
      Width           =   6375
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "Add Your text"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text Size"
      Height          =   615
      Left            =   1920
      TabIndex        =   15
      Top             =   1680
      Width           =   6375
      Begin VB.CommandButton cmdText 
         Caption         =   "Tiny"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Small"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Medium"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Big"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Huge"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6480
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdFindFiles 
      Caption         =   "Find Files or Folders"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdFindTxt 
      Caption         =   "Find (on this page)"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdPSC 
      Caption         =   "PSC"
      Height          =   495
      Left            =   7440
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   2400
      Width           =   6255
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Refresh"
      Height          =   495
      Index           =   3
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Stop"
      Height          =   495
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Forward"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser Brow 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   8175
      ExtentX         =   14420
      ExtentY         =   6588
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   6720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label5 
      Caption         =   "Auto Complete Example"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label LblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   6720
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Forwards As Boolean
Dim Backwards As Boolean
Dim nonav As Boolean
Private Sub Brow_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
'sets our back/forward buttons enabled state
On Error Resume Next
Select Case Command
    Case CSC_NAVIGATEFORWARD
            Forwards = Enable
    Case CSC_NAVIGATEBACK
            Backwards = Enable
End Select
cmdNav(0).Enabled = Backwards
cmdNav(1).Enabled = Forwards
If Command = -1 Then Exit Sub

End Sub

Private Sub Brow_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Dim temp As String
nonav = True
temp = Left(Brow.LocationURL, 5)
'If using the about method dont show in the address combo
If Left(Brow.LocationURL, 6) = "about:" Then
    nonav = False
    Exit Sub
End If
'get rid of garbage in the web address - looks neater
temp = ReplaceAll(Brow.LocationURL, "%20", " ")
If Left(temp, 4) = "http" Then temp = Right(temp, Len(temp) - 7)
If Left(temp, 4) = "file" Then temp = Right(temp, Len(temp) - 8)
'make sure the new address is not already in the address combo
z = 0
For X = 0 To cboAddress.ListCount - 1
    If cboAddress.List(X) = temp Then
        z = 1
        Exit For
    End If
Next X
If z = 0 Then
    'if it's a new address add to combo list
    'and move combolist to show this address
    cboAddress.AddItem temp
    cboAddress.ListIndex = cboAddress.ListCount - 1
Else
    'if it's already in the list dont add it
    'but move the combolist to display
    'the old entry
    cboAddress.ListIndex = X
End If
    'put just the pages name in our title bar
    'not the full address
Me.Caption = Brow.LocationName
nonav = False

End Sub

Private Sub Brow_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    If PB.Value = 100 Then
        PB.Visible = False
    Else
        PB.Visible = True
    End If
    If Progress = -1 Then PB.Value = 100
    If Progress > 0 And ProgressMax > 0 Then
        PB.Value = Progress * 100 / ProgressMax
    End If

End Sub

Private Sub Brow_StatusTextChange(ByVal Text As String)
LblStatus = Text
'show the loading status
End Sub

Private Sub cboAddress_Click()
If nonav = False Then
    'when a new listitem is clicked - go there
        Brow.Navigate2 cboAddress.Text
End If

End Sub

Private Sub cboAddress_GotFocus()
'select the text so the user can easily enter a new address
SendKeys String:="{HOME}+{END}", Wait:=True

End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        'enter was pressed - lets go
        If nonav = False Then
            Brow.Navigate2 cboAddress.Text
        End If
End Select

End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        'enter was pressed - lets go
        If nonav = False Then
            Brow.Navigate2 cboAddress.Text
        End If
End If

End Sub

Private Sub cboAddress_KeyUp(KeyCode As Integer, Shift As Integer)
'This is my auto-complete method
'You'll have to wade through the logic behind this
Dim curlentxt As Integer
curlentxt = Len(cboAddress.Text)
Select Case KeyCode
    Case vbKeyReturn, vbKeyBack, vbKeyShift, vbKeyClear, vbKeyDelete
        Exit Sub
    Case Else
        If nonav = False Then
           Dim tempaddress As String
            Dim tempcount As Integer
            For X = 0 To cboAddress.ListCount - 1
                If Left(UCase(cboAddress.List(X)), 7) = "HTTP://" Then
                    tempaddress = UCase(Right(UCase(cboAddress.List(X)), Len(UCase(cboAddress.List(X))) - 7))
                    tempcount = 7
                End If
                If Left(UCase(cboAddress.List(X)), 8) = "FILE:///" Then
                    tempaddress = UCase(Right(UCase(cboAddress.List(X)), Len(UCase(cboAddress.List(X))) - 8))
                    tempcount = 8
                End If
                If UCase(cboAddress.Text) = UCase(Left(tempaddress, curlentxt)) Then
                    cboAddress.Text = cboAddress.List(X)
                    cboAddress.SelStart = curlentxt + tempcount
                    cboAddress.SelLength = Len(cboAddress.List(X)) - (curlentxt + tempcount)
                    Exit For
                ElseIf UCase(cboAddress.Text) = UCase(Left(cboAddress.List(X), curlentxt)) Then
                    cboAddress.Text = cboAddress.List(X)
                    cboAddress.SelStart = curlentxt
                    cboAddress.SelLength = Len(cboAddress.List(X)) - curlentxt
                    Exit For
                End If
            Next X
        End If
End Select

End Sub

Private Sub cmdEdit_Click(Index As Integer)
'Editing functions - easy isn't it
Select Case Index
    Case 0
        Brow.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DONTPROMPTUSER
    Case 1
        Brow.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER
    Case 2
        Brow.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DONTPROMPTUSER
    Case 3
        Brow.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End Select

End Sub

Private Sub cmdFindFiles_Click()
On Error Resume Next 'Spits an error without this line
Brow.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DONTPROMPTUSER

End Sub

Private Sub cmdFindTxt_Click()
'Mimics ctl+f keys being pressed to bring up the
'find text dialog
Brow.SetFocus
SendKeys "^f", True

End Sub

Private Sub cmdNav_Click(Index As Integer)
Select Case Index
    Case 0
        Brow.GoBack
    Case 1
        Brow.GoForward
    Case 2
        Brow.Stop
    Case 3
        Brow.Refresh
End Select
Brow.SetFocus
End Sub

Private Sub cmdOpen_Click()
'Standard common dialog stuff
    On Error GoTo woops
        With CommonDialog1
           .DialogTitle = "Open Local Web Page"
           .CancelError = True
           .Filter = "Web Pages (*.htm;*.html)|*.htm;*.html|All files (*.*)|*.*"
           .ShowOpen
        If Len(.Filename) = 0 Then Exit Sub
        If FileExists(.Filename) Then Brow.Navigate .Filename
        End With
woops:

End Sub

Private Sub cmdPSC_Click()
'lets go get some more code !
Brow.Navigate "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=1"
End Sub

Private Sub cmdSave_Click()
On Error GoTo woops
If Brow.LocationURL = "" Then Exit Sub
Brow.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
woops:

End Sub

Private Sub cmdText_Click(Index As Integer)
'this is how we change the text size - accepts 0 to 4 as long
Brow.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(Index)
End Sub



Private Sub Form_Load()
'Uses vbs' settings to store MRUs'
'The nonav variable stops the browser navigating while we fiddle
'with its' address combo.
nonav = True
temp = GetSetting(App.Title, "Settings", "MRUcount", "")
If temp = "" Then GoTo carryon
m = Val(temp)
For X = 0 To m - 1
    temp = GetSetting(App.Title, "Settings", "MRU " + Str(X), "")
    If temp <> "" Then cboAddress.AddItem temp
Next X
If cboAddress.ListCount > 0 Then cboAddress.ListIndex = 0
nonav = False
carryon:
LoadPage 'writes a web page and saves it to the app.path
'This is what you see when the program starts
Brow.Navigate App.Path + "\Blank.htm"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'save MRUs' and delete the opening web page
Dim m As Integer
m = 0
For X = cboAddress.ListCount - 1 To 0 Step -1
    If m = 100 Then Exit For
    If cboAddress.List(X) <> "" Then
        SaveSetting App.Title, "Settings", "MRU " + Str(m), cboAddress.List(X)
        m = m + 1
    End If
Next X
SaveSetting App.Title, "Settings", "MRUcount", Str(m)
If FileExists(App.Path + "\Blank.htm") Then Kill App.Path + "\Blank.htm"

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'the about method demo
Brow.Navigate "about: " + Text1.Text
End Sub

Public Function ReplaceAll(SourceString As String, ReplaceThis As String, WithThis As String)
'used to clean web addresses
    Dim temp As Variant
    temp = Split(SourceString, ReplaceThis)
    ReplaceAll = Join(temp, WithThis)
End Function
Function FileExists(ByVal Filename As String) As Integer
'used to stop errors if a file does not exist
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(Filename)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, MB_OK, "Error"
                End
            End If
    End Select
End Function



Public Sub LoadPage()
'writing a web page demo
'standard HTML source code here
Open App.Path + "\Blank.htm" For Output As #1
Print #1, "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
Print #1, "<html><head><title>IE Lite</title></head><body>"
Print #1, "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & ">"
Print #1, "<body bgcolor=" & Chr(34) & "#8DDAF3" & Chr(34) & " text=" & Chr(34) & "#8000FF" & Chr(34) & " link=" & Chr(34) & "#8000FF" & Chr(34) & " vlink=" & Chr(34) & "#8000FF" & Chr(34) & " alink=" & Chr(34) & "#8000FF" & Chr(34) & ">"
Print #1, "<p align=center><font color=#FF0000 size=5>" + "Web Browser Tricks"; "</font></p><br>"
Print #1, "<p align=center><font color=#FF0000 size=5>" + "MrBobo 2000"; "</font></p><br><br>"
Print #1, "<p align=left><font color=#FF0000 size=5>" + "This Example Application Demonstates :"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "1. Enable/Disable Forward and back Buttons"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "2. Cut, Copy, Paste and Select All edit functions"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "3. Auto-Complete Address Box"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "4. Saving and Loading MRUs"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "5. Avoiding duplicates in Combo/List boxes"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "6. Standard Navigation Buttons"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "7. The 'about:' Navigation method"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "8. Opening and Saving Web Pages"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "9. Using Explorers' Find File Dialog"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "10. Finding text on current page"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "11. Sizing text on Web pages"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "12. Showing a progress guage"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "13. Showing Status text"; "</font></p><br>"
Print #1, "<p align=left><font color=#000000 size=4>" + "14. Creating HTML pages at runtime"; "</font></p><br>"
Close #1

End Sub
