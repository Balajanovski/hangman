VERSION 5.00
Begin VB.Form LoadingForm 
   BackColor       =   &H00000000&
   Caption         =   "Loading..."
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   Picture         =   "Loading Screen.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   12360
      Top             =   120
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
    Presents.Top = -500
    TimerIndex = 0
    Index = 0
End Sub


Private Sub Timer1_Timer()
    If Presents.Top < 1320 Then
        Presents.Top = Presents.Top + 15
    End If
    
    Index = TimerIndex \ 35 Mod 6
    If Index >= 6 Then
        TimerIndex = 0
        For i = 0 To 5
            LoadingBar(i).Visible = False
        Next
    End If
    LoadingBar(Index).Visible = True
    
    TimerIndex = TimerIndex + 1
    
    
End Sub
