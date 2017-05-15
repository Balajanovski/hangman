VERSION 5.00
Begin VB.Form LoadingForm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Loading Screen.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   12360
      Top             =   120
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "HANGMAN"
      BeginProperty Font 
         Name            =   "Clarendon Lt BT"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   6840
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label PlayButton 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PLAY"
      BeginProperty Font 
         Name            =   "Clarendon BT"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7920
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Shape LoadingBar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   495
      Index           =   5
      Left            =   10440
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape LoadingBar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   495
      Index           =   4
      Left            =   9960
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape LoadingBar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   495
      Index           =   3
      Left            =   9480
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape LoadingBar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   495
      Index           =   2
      Left            =   9000
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape LoadingBar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   8520
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape LoadingBar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   8040
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Presents 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PRESENTS"
      BeginProperty Font 
         Name            =   "Clarendon Blk BT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7440
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   1095
      Left            =   7800
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Label Loading 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Clarendon Blk BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8520
      TabIndex        =   0
      Top             =   4320
      Width           =   2055
   End
End
Attribute VB_Name = "LoadingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TimerIndex As Integer
Dim Index As Integer

Private Sub Form_Load()
    
    ' Initialise various properties
    Presents.Top = -500
    TimerIndex = 0
    Index = 0
    PlayButton.Enabled = False
    
End Sub


Private Sub PlayButton_Click()
    Call SwapWindow(LoadingForm, PickScreen)
End Sub

Private Sub Timer1_Timer()

    ' Move the presents label down
    If Presents.Top < 1320 Then
        Presents.Top = Presents.Top + 15
    End If
    
    ' Create an index from the timer's iterations for the loading bar
    Index = TimerIndex \ 35 Mod 6
    
    If Index >= 5 Then
        ' Stop iterating the timer
        Timer1.Enabled = False
        
        ' Make the play button usable and visible
        PlayButton.Enabled = True
        PlayButton.Visible = True
        
        ' Set the loading text to say Press to Play!
        Loading.Caption = "Press to Play!"
        Loading.Top = Loading.Top - 300
        Loading.Height = Loading.Height + 300
        
        ' Set the title to be visible
        Title.Visible = True
        
        ' Change title bar of window
        LoadingForm.Caption = "Hangman - By James, Rishabh & Dibaloak"
    End If
    
    ' Turn on segments of the loading bar
    LoadingBar(Index).Visible = True
    
    ' Increment
    TimerIndex = TimerIndex + 1
   
End Sub
