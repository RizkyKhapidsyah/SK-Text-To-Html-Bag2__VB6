VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Setari 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "Setari.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Setari.frx":0442
   ScaleHeight     =   4350
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Color the tags for better viewing"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   4440
      Picture         =   "Setari.frx":174A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Text color"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1560
      Picture         =   "Setari.frx":1B8C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Page title"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      Picture         =   "Setari.frx":1FCE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Page BGImage"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   4440
      Picture         =   "Setari.frx":2710
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Page BGColor"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   3000
      Picture         =   "Setari.frx":2B52
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Active link color"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   3000
      Picture         =   "Setari.frx":2F94
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Visited link color"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1560
      Picture         =   "Setari.frx":33D6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Link color"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      Picture         =   "Setari.frx":3818
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Open text file"
      Height          =   1095
      Left            =   120
      Picture         =   "Setari.frx":3C5A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detect all special characters"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   4440
      Picture         =   "Setari.frx":409C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Leave the text as I typed it"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   3000
      Picture         =   "Setari.frx":44DE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detect all hyperlinks"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1560
      Picture         =   "Setari.frx":4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Height          =   615
      Left            =   3000
      Picture         =   "Setari.frx":4D62
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      Picture         =   "Setari.frx":51A4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Ready"
      Top             =   3720
      Width           =   2775
   End
End
Attribute VB_Name = "Setari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Timpul As Long
Dim Mesaj As String
If Option1.Value = True Then
  Mesaj = Mesaj & "Type= Leave the text as I typed it" & vbCrLf
ElseIf Option2.Value = True Then
  Mesaj = Mesaj & "Type= Detect all special characters" & vbCrLf
End If
Mesaj = Mesaj & "Detect links= " & Format(Check1.Value, "Yes/No") & vbCrLf
Mesaj = Mesaj & "Page title= " & Titlu & vbCrLf
Mesaj = Mesaj & "Text color= " & CulText & vbCrLf
Mesaj = Mesaj & "Background: " & Fond & vbCrLf
Mesaj = Mesaj & "Link color= " & LinkX & vbCrLf
Mesaj = Mesaj & "Visited link color= " & VLinkX & vbCrLf
Mesaj = Mesaj & "Active link color= " & ALinkX & vbCrLf
Mesaj = Mesaj & "Identify tags by color= " & Format(Check2.Value, "Yes/No")

If vbNo = MsgBox("Are you sure this are the settings you want ?" & vbCrLf & vbCrLf & Mesaj, vbQuestion + vbYesNo, "Final preparations") Then Exit Sub
Setari.Hide
MsgBox "I'm ready to process your data...." & vbCrLf & "If you selected a long file then grab a Coke and wait..." & vbCrLf & "I will announce you when I'm done.", vbInformation, "Ready"
Timpul = Timer

If Option1.Value = True And Check1.Value = 0 Then
   ContinutFisier = "<pre>" & vbCrLf & ContinutFisier & vbCrLf & "</pre>" & vbCrLf
ElseIf Option1.Value = True And Check1.Value = 1 Then
   ContinutFisier = "<pre>" & vbCrLf & PreTag(ContinutFisier) & vbCrLf & "</pre>" & vbCrLf
ElseIf Option2.Value = True And Check1.Value = 0 Then
   ContinutFisier = "<p>" & vbCrLf & CautareSimpla(ContinutFisier) & vbCrLf & "</p>" & vbCrLf
ElseIf Option2.Value = True And Check1.Value = 1 Then
   ContinutFisier = "<p>" & vbCrLf & CautareAvansata(ContinutFisier) & vbCrLf & "</p>" & vbCrLf
End If


Static strTempHTML As String
strTempHTML = "  <!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">" & vbCrLf & "<HTML>" & vbCrLf & "<HEAD>" & vbCrLf & "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & ">" & vbCrLf & "<TITLE>" & Titlu & "</TITLE>" & vbCrLf & "<META name=" & Chr(34) & "description" & Chr(34) & " content=" & Chr(34) & "Personal page" & Chr(34) & ">" & vbCrLf & "<META name=" & Chr(34) & "author" & Chr(34) & " content=" & Chr(34) & "Text to Html converter by Classicmanpro" & Chr(34) & ">" & vbCrLf & "</head>" & vbCrLf & "<body " & Fond & " text=" & Chr(34) & CulText & Chr(34) & " link=" & Chr(34) & LinkX & Chr(34) & " vlink=" & Chr(34) & VLinkX & Chr(34) & " alink=" & Chr(34) & ALinkX & Chr(34) & ">" & vbCrLf
ContinutFisier = strTempHTML & ContinutFisier & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf

TextToHtml.RTB1.Text = ContinutFisier

If Setari.Check2.Value = 1 Then
  Call Colorare
End If
Unload Setari

Timpul = (Timer - Timpul) / 60
MsgBox "Done in " & Timpul & " minutes."

Salvat = False
End Sub

Private Sub Command10_Click()
Paleta.Show 1
CulText = CuloareTemp
End Sub

Private Sub Command2_Click()
Unload Setari
End Sub

Private Sub Command3_Click()
ChDir App.Path
On Error GoTo MYE
cmd1.Filter = "All Files(*.*)|*.*|Text Files(*.txt)|*.txt"
cmd1.FilterIndex = 0
cmd1.ShowOpen
FileName = cmd1.FileName

If LCase(Right(FileName, 3)) = "txt" Then GoTo Urmatorul
MsgBox "Format not supported..." & vbCrLf & "Only Text Files are accepted...", vbCritical, "Open"
Exit Sub

Urmatorul:
If Verificare(FileName) = False Then
   MsgBox "The text file you requested is missing...", vbCritical, "Open"
   Exit Sub
End If

Open FileName For Input As #3
 ContinutFisier = Input$(LOF(3), 3)
Close #3

Command1.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Exit Sub
MYE:
 Exit Sub
End Sub

Private Sub Command4_Click()
Paleta.Show 1
Fond = " bgcolor=" & Chr(34) & CuloareTemp & Chr(34)
End Sub


Private Sub Command5_Click()
Paleta.Show 1
LinkX = CuloareTemp
End Sub

Private Sub Command6_Click()
Paleta.Show 1
VLinkX = CuloareTemp
End Sub


Private Sub Command7_Click()
Paleta.Show 1
ALinkX = CuloareTemp
End Sub


Private Sub Command8_Click()
ChDir App.Path
On Error GoTo Jos1
cmd1.Filter = "All Files(*.*)|*.*|Gif Files(*.gif)|*.gif|Jpg Files(*.jpg)|*.jpg|Png files(*.png)|*.png"
cmd1.FilterIndex = 0
cmd1.ShowOpen
FileName = cmd1.FileName

If LCase(Right(FileName, 3)) = "gif" Or LCase(Right(FileName, 3)) = "png" Or LCase(Right(FileName, 3)) = "jpg" Then GoTo SeC
MsgBox "Unsupported format", vbCritical
Exit Sub

SeC:
Dim I As Integer
For I = 1 To Len(FileName)
  If Mid(FileName, I, 1) = "\" Then Mid(FileName, I, 1) = "/"
Next I

FileName = "file:///" & FileName
Fond = " background=" & Chr(34) & FileName & Chr(34)

MsgBox "The path of your Image is ABSOLUTE..." & vbCrLf & "If you want to modify it you'll find it inside the <body background=....> tag (after the conversion)", vbInformation, "Important"
Exit Sub
Jos1:
  Exit Sub
End Sub


Private Sub Command9_Click()
Titlu = InputBox("Insert your title...", "Head Setting", "My Page")
End Sub


