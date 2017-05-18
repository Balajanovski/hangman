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
   Begin VB.PictureBox NotTimed 
      Enabled         =   0   'False
      Height          =   1800
      Left            =   6840
      Picture         =   "PickScreen.frx":0000
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox HardButton 
      Enabled         =   0   'False
      Height          =   1800
      Left            =   8280
      Picture         =   "PickScreen.frx":23A9
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox ComputerScienceButton 
      Height          =   1800
      Left            =   8280
      Picture         =   "PickScreen.frx":42E9
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   3
      Top             =   2760
      Width           =   1800
   End
   Begin VB.PictureBox ChemistryButton 
      Height          =   1800
      Left            =   5640
      Picture         =   "PickScreen.frx":6702
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   1
      Top             =   2760
      Width           =   1800
   End
   Begin VB.PictureBox MathButton 
      Height          =   1800
      Left            =   2760
      Picture         =   "PickScreen.frx":8E7A
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   2
      Top             =   2760
      Width           =   1800
   End
   Begin VB.PictureBox EasyButton 
      Enabled         =   0   'False
      Height          =   1800
      Left            =   2760
      Picture         =   "PickScreen.frx":A52C
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox Timed 
      Enabled         =   0   'False
      Height          =   1800
      Left            =   4080
      Picture         =   "PickScreen.frx":BE26
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox NormalButton 
      Enabled         =   0   'False
      Height          =   1800
      Left            =   5640
      Picture         =   "PickScreen.frx":D900
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Instruction 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   1560
      TabIndex        =   12
      Top             =   960
      Width           =   9615
   End
   Begin VB.Label ShowInstruct 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Show Instructions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10440
      TabIndex        =   11
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label LabelTimed 
      Caption         =   "Label1"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label LabelDifficulty 
      Caption         =   "Label1"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Message 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pick a Category"
      BeginProperty Font 
         Name            =   "Arial"
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
Dim TopicState As String
Dim Difficulty As Integer
Dim IsTimed As Boolean

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

Private Sub SwitchToTimedSelect()
    Message.Caption = "Timed?"
    Me.Caption = "Timed?"
    
    EasyButton.Visible = False
    NormalButton.Visible = False
    HardButton.Visible = False
    
    Timed.Visible = True
    NotTimed.Visible = True
    
    Timed.Enabled = True
    NotTimed.Enabled = True
End Sub

' Event handlers for clicking the topic selection buttons
Private Sub ChemistryButton_Click()
    TopicState = "Chemistry"
    Call SwitchToDifficultySelect
End Sub

Private Sub ComputerScienceButton_Click()
    TopicState = "CS"
    Call SwitchToDifficultySelect
End Sub

Private Sub Form_Load()
    Instruction.Visible = False
    Instruction.Caption = "Instructions:" + vbCrLf + "You have to guess the individual letters in the word by either using your keyboard or the keys provided to you." + vbCrLf + "You have to get four words, from the Categories of either Math, Chemistry or Computer Science, correct in order to win the game." + vbCrLf + "You have a choice of three difficulties." + vbCrLf + "-On Easy difficulty you will be provided with 2 extra turns after every correct word and have 120 seconds in the Timed Mode, along with a hint." + vbCrLf + "-On Normal difficulty you will be provided with 1 extra turn after every correct word and have 90 seconds in the Timed Mode." + vbCrLf + "-On Hard difficulty you will not given any extra turns after a correct word and have 60 seconds in the Timed Mode."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Instruction.Visible = False
    If Message.Caption = "Pick a Category" Then
        MathButton.Visible = True
        ChemistryButton.Visible = True
        ComputerScienceButton.Visible = True
    ElseIf Message.Caption = "Pick a Difficulty" Then
        EasyButton.Visible = True
        NormalButton.Visible = True
        HardButton.Visible = True
    Else
        Timed.Visible = True
        NotTimed.Visible = True
    End If
End Sub


Private Sub MathButton_Click()
    TopicState = "Math"
    Call SwitchToDifficultySelect
End Sub

Private Sub EasyButton_Click()
    LabelDifficulty.Caption = -1
    Call SwitchToTimedSelect
End Sub

Private Sub HardButton_Click()
    LabelDifficulty.Caption = 1
    Call SwitchToTimedSelect
End Sub

Private Sub NormalButton_Click()
    LabelDifficulty.Caption = 0
    Call SwitchToTimedSelect
End Sub

Private Sub LoadForm()
    If TopicState = "Math" Then
        CategoryMath.Show
        Unload Me
    ElseIf TopicState = "Chemistry" Then
        CategoryChem.Show
        Unload Me
    ElseIf TopicState = "CS" Then
        CategoryCS.Show
        Unload Me
    End If
End Sub

Private Sub ShowInstruct_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Instruction.Visible = True
    MathButton.Visible = False
    ChemistryButton.Visible = False
    ComputerScienceButton.Visible = False
    EasyButton.Visible = False
    NormalButton.Visible = False
    HardButton.Visible = False
    Timed.Visible = False
    NotTimed.Visible = False
End Sub


Private Sub Timed_Click()
    LabelTimed.Caption = True
    Call LoadForm
End Sub

Private Sub NotTimed_Click()
    LabelTimed.Caption = False
    Call LoadForm
End Sub
