VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TextToHtml 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web design tool"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7470
   Icon            =   "TextToHtml.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "TextToHtml.frx":0442
   ScaleHeight     =   5610
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmd 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8281
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   2
      MousePointer    =   99
      DisableNoScroll =   -1  'True
      TextRTF         =   $"TextToHtml.frx":3F4A
      MouseIcon       =   "TextToHtml.frx":400E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuline2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnucopy 
         Caption         =   "Cop&y"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "E&mpty clipboard"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuprev 
      Caption         =   "Browser previe&w"
   End
   Begin VB.Menu mnuinfo 
      Caption         =   "&Info"
      Begin VB.Menu mnuhelpme 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "TextToHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
Salvat = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Salvat = False Then
  If vbYes = MsgBox("You forgot to save the file..." & vbCrLf & "Do you want to save it ?", vbCritical + vbYesNo, "File") Then Call mnusave_Click
End If

ChDir App.Path
Open "TempPrev.html" For Output As #6
Close #6
Kill "TempPrev.html"
End
End Sub


Private Sub mnuabout_Click()
About.Show 1
End Sub

Private Sub mnuclear_Click()
Clipboard.Clear
End Sub

Private Sub mnucopy_Click()
Clipboard.SetText (RTB1.SelText)
End Sub

Private Sub mnucut_Click()
Clipboard.SetText (RTB1.SelText)
RTB1.SelText = ""
End Sub


Private Sub mnuexit_Click()
If Salvat = False Then
  If vbYes = MsgBox("You forgot to save the file..." & vbCrLf & "Do you want to save it ?", vbCritical + vbYesNo, "File") Then Call mnusave_Click
End If

ChDir App.Path
Open "TempPrev.html" For Output As #5
Close #5
Kill "TempPrev.html"
End
End Sub


Private Sub mnuhelpme_Click()
frmAjutor.Show 1
End Sub

Private Sub mnuopen_Click()
If Salvat = False Then
  If vbYes = MsgBox("You forgot to save the file..." & vbCrLf & "Do you want to save it ?", vbCritical + vbYesNo, "File") Then Call mnusave_Click
End If

Setari.Show 1
End Sub

Private Sub mnupaste_Click()
RTB1.SelText = Clipboard.GetText
End Sub


Private Sub mnuprev_Click()
ChDir App.Path
Open "TempPrev.html" For Output As #4
  Print #4, RTB1.Text
Close #4

ShellExecute 0&, vbNullString, "TempPrev.html", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub mnusave_Click()
ChDir App.Path
On Error GoTo ErrorHandler
cmd.Filter = "All files(*.*)|*.*|HTML Files(*.html)|*.html|HTM Files(*.htm)|*.htm|Text Files(*.txt)|*.txt"
cmd.FilterIndex = 0
cmd.ShowSave
FileName = cmd.FileName

If LCase(Right(FileName, 3)) = "txt" Or LCase(Right(FileName, 4)) = "html" Or LCase(Right(FileName, 3)) = "htm" Then GoTo Etapa
MsgBox "Format not supported", vbCritical, "Save"
Exit Sub

Etapa:
If Verificare(FileName) = True Then
     If vbNo = MsgBox("The following file exists !" & vbCrLf & FileName & vbCrLf & "Do you want to replace it ?", vbCritical + vbYesNo, "Save") Then Exit Sub
End If
Open FileName For Output As #2
  Print #2, RTB1.Text
Close #2

Salvat = True
Exit Sub
ErrorHandler:
  Exit Sub
End Sub


Private Sub RTB1_Change()
Salvat = False
End Sub


