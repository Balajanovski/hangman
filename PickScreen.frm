VERSION 5.00
Begin VB.Form PickScreen 
   BackColor       =   &H00000000&
   Caption         =   "Pick a Category..."
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Mathematics 
      Height          =   1800
      Left            =   8280
      Picture         =   "PickScreen.frx":0000
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   3
      Top             =   2760
      Width           =   1800
   End
   Begin VB.PictureBox Chemistry 
      Height          =   1800
      Left            =   2760
      Picture         =   "PickScreen.frx":2419
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   2
      Top             =   2760
      Width           =   1800
   End
   Begin VB.PictureBox ComputerScience 
      Height          =   1800
      Left            =   5640
      Picture         =   "PickScreen.frx":3ACB
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   1
      Top             =   2760
      Width           =   1800
   End
   Begin VB.Label Message 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pick a Category"
      BeginProperty Font 
         Name            =   "Clarendon Blk BT"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "PickScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture3_Click()

End Sub

Private Sub Button_Click(Index As Integer)

End Sub
