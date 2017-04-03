VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton mathButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Math"
      BeginProperty Font 
         Name            =   "Proxy 3"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton csButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Computer Science"
      BeginProperty Font 
         Name            =   "Proxy 5"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton play 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PLAY"
      BeginProperty Font 
         Name            =   "Proxy 3"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Timer backgroundTimer 
      Interval        =   20
      Left            =   120
      Top             =   120
   End
   Begin VB.Label title 
      BackColor       =   &H00000000&
      Caption         =   "HANGMAN"
      BeginProperty Font 
         Name            =   "Proxy 5"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HasTakenBeginTime As Boolean
Dim MoveFlag As Boolean

Dim TimeDifference As Integer

Const STARTING_SHADE As Integer = 10
Const ENDING_SHADE As Integer = 100

Private Sub Form_Load()
    backgroundTimer.Enabled = True
    Me.BackColor = "&H00" & "00" & Hex$(STARTING_SHADE)
    TimeDifference = STARTING_SHADE
    MoveFlag = False
End Sub

Private Sub play_Click()
    title.Caption = " Topics"
    csButton.Visible = True
    mathButton.Visible = True
    play.Visible = False
End Sub

' This method is responsible for the changing colour of the background
Private Sub backgroundTimer_Timer()
    Dim NewColor As String
    If Not MoveFlag Then
        TimeDifference = TimeDifference + 1
        If TimeDifference >= ENDING_SHADE Then
            MoveFlag = True
        Else
            NewColor = "&H00" & Hex$(TimeDifference) & "00"
            Me.BackColor = NewColor
            title.BackColor = NewColor
        End If
    Else
        TimeDifference = TimeDifference - 1
        If TimeDifference <= STARTING_SHADE Then
            MoveFlag = False
        Else
            NewColor = "&H00" & Hex$(TimeDifference) & "00"
            Me.BackColor = NewColor
            title.BackColor = NewColor
        End If
    End If
End Sub
