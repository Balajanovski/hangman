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
   Begin VB.PictureBox HardButton 
      Enabled         =   0   'False
      Height          =   1800
      Left            =   8280
      Picture         =   "PickScreen.frx":0000
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox NormalButton 
      Enabled         =   0   'False
      Height          =   1800
      Left            =   5640
      Picture         =   "PickScreen.frx":1F40
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox EasyButton 
      Enabled         =   0   'False
      Height          =   1800
      Left            =   2760
      Picture         =   "PickScreen.frx":36C5
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox ComputerScienceButton 
      Height          =   1800
      Left            =   8280
      Picture         =   "PickScreen.frx":4FBF
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   3
      Top             =   2760
      Width           =   1800
   End
   Begin VB.PictureBox MathButton 
      Height          =   1800
      Left            =   2760
      Picture         =   "PickScreen.frx":73D8
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   2
      Top             =   2760
      Width           =   1800
   End
   Begin VB.PictureBox ChemistryButton 
      Height          =   1800
      Left            =   5640
      Picture         =   "PickScreen.frx":8A8A
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
Private Sub SwitchToDifficultySelect()
    ' Change the text and the window caption
    Message.Caption = "Pick a Difficulty"
    Me.Caption = "Pick a Difficulty"
    
    ' Hide the old buttons
    MathButton.Visible = False
    ChemistryButton.Visible = False
    ComputerScienceButton.Visible = False
    
    ' Show new ones
    EasyButton.Visible = True
    NormalButton.Visible = True
    HardButton.Visible = True
    
    ' Enable the new buttons
    EasyButton.Enabled = True
    NormalButton.Enabled = True
    HardButton.Enabled = True
End Sub

' Event handlers for clicking the topic selection buttons
Private Sub ChemistryButton_Click()
    TopicState = Chemistry
    Call SwitchToDifficultySelect
End Sub

Private Sub ComputerScienceButton_Click()
    TopicState = CS
    Call SwitchToDifficultySelect
End Sub


Private Sub MathButton_Click()
    TopicState = Math
    Call SwitchToDifficultySelect
End Sub

Private Sub EasyButton_Click()
    DifficultyState = Easy
End Sub

Private Sub HardButton_Click()
    DifficultyState = Hard
End Sub

Private Sub NormalButton_Click()
    DifficultyState = Normal
End Sub
