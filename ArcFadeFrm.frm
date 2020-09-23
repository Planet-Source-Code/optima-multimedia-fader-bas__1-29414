VERSION 5.00
Begin VB.Form ArcFadeFrm 
   Caption         =   "Fade It"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "FrmBackFade"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FadeBackOval"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SquareBackFade"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TextBackFade"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FadeForeLabel"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FadeBackPicture"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Fader"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1800
      Shape           =   5  'Rounded Square
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   3240
      Shape           =   2  'Oval
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "ArcFadeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
PulseBackFade 255, 234, 56, 76, 205, 200, Picture1, 0.1, 10
End Sub

Private Sub Command2_Click()
PulseForeFade 34, 65, 234, 200, 156, 23, Label2, 0.01, 30

End Sub

Private Sub Command3_Click()
PulseBackFade 255, 134, 56, 76, 205, 100, Text1, 0.1, 10
End Sub

Private Sub Command4_Click()
PulseBackFade 155, 134, 56, 76, 105, 100, Shape3, 0.2, 20
End Sub

Private Sub Command5_Click()
PulseBackFade 255, 255, 255, 0, 0, 0, Shape2, 0.001, 1
End Sub

Private Sub Command6_Click()
PulseBackFade 255, 0, 56, 255, 205, 50, ArcFadeFrm, 0.1, 10
End Sub


