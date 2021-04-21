VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Created by Classicmanpro"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "About.frx":0442
   ScaleHeight     =   2985
   ScaleWidth      =   2970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "About.frx":0A7B
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday 12 - Dec - 2001"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   120
      MouseIcon       =   "About.frx":0EBD
      MousePointer    =   99  'Custom
      Picture         =   "About.frx":11C7
      Stretch         =   -1  'True
      ToolTipText     =   "Visit Classic software"
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text to HTML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1905
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload About
End Sub



