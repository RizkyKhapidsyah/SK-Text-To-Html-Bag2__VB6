VERSION 5.00
Begin VB.Form frmAjutor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmAjutor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAjutor.frx":0442
   ScaleHeight     =   4500
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmAjutor.frx":1CDA
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmAjutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmAjutor
End Sub
