VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Paleta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color bank"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "Paleta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Paleta.frx":0442
   ScaleHeight     =   3750
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   3435
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "Red"
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   327682
      LargeChange     =   15
      Max             =   255
      SelStart        =   255
      TickStyle       =   1
      TickFrequency   =   15
      Value           =   255
   End
   Begin ComctlLib.Slider Slider3 
      Height          =   510
      Left            =   480
      TabIndex        =   5
      ToolTipText     =   "Blue"
      Top             =   2160
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   900
      _Version        =   327682
      LargeChange     =   15
      Max             =   255
      SelStart        =   255
      TickStyle       =   1
      TickFrequency   =   15
      Value           =   255
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Green"
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   327682
      LargeChange     =   15
      Max             =   255
      SelStart        =   255
      TickStyle       =   1
      TickFrequency   =   15
      Value           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Ready"
      Default         =   -1  'True
      Height          =   255
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ready"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Text            =   "FFFFFF"
      ToolTipText     =   "Selected color"
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color code"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1560
      TabIndex        =   10
      Top             =   2760
      Width           =   1080
   End
End
Attribute VB_Name = "Paleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rosu As String
Dim Verde As String
Dim Albastru As String
Private Sub Command1_Click()
CuloareTemp = "#" & Text1.Text
Unload Paleta
End Sub

Private Sub Command2_Click()
Unload Paleta
End Sub









Private Sub Slider1_Change()
Picture1.BackColor = RGB(Slider1.Value, Slider2.Value, Slider3.Value)
Rosu = Hex(Slider1.Value)
Verde = Hex(Slider2.Value)
Albastru = Hex(Slider3.Value)
If Len(Rosu) <= 1 Then Rosu = "0" & Rosu
If Len(Verde) <= 1 Then Verde = "0" & Verde
If Len(Albastru) <= 1 Then Albastru = "0" & Albastru
Text1.Text = Rosu & Verde & Albastru
End Sub












Private Sub Slider2_Change()
Picture1.BackColor = RGB(Slider1.Value, Slider2.Value, Slider3.Value)
Rosu = Hex(Slider1.Value)
Verde = Hex(Slider2.Value)
Albastru = Hex(Slider3.Value)
If Len(Rosu) <= 1 Then Rosu = "0" & Rosu
If Len(Verde) <= 1 Then Verde = "0" & Verde
If Len(Albastru) <= 1 Then Albastru = "0" & Albastru
Text1.Text = Rosu & Verde & Albastru
End Sub



Private Sub Slider3_Change()
Picture1.BackColor = RGB(Slider1.Value, Slider2.Value, Slider3.Value)
Rosu = Hex(Slider1.Value)
Verde = Hex(Slider2.Value)
Albastru = Hex(Slider3.Value)
If Len(Rosu) <= 1 Then Rosu = "0" & Rosu
If Len(Verde) <= 1 Then Verde = "0" & Verde
If Len(Albastru) <= 1 Then Albastru = "0" & Albastru
Text1.Text = Rosu & Verde & Albastru
End Sub




