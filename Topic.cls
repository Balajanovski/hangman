VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Topic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private EasyWords(4) As String
Private NormalWords(4) As String
Private DifficultWords(4) As String

Private EasyWordsIndex As Integer
Private NormalWordsIndex As Integer
Private DifficultWordsIndex As Integer

Private RecentlyPickedWord(4) As String
Private RecentlyPickedWordIndex As Integer


Public Sub Init(ByRef Words() As String)
    EasyWordsIndex = 0
    NormalWordsIndex = 0
    DifficultWordsIndex = 0
    RecentlyPickedWordIndex = 0

    For i = 0 To UBound(Words)
        Dim Length As Integer
        Length = Len(Words(i))
        
        
        If Length >= 7 Then
            DifficultWords(DifficultWordsIndex) = Words(i)
            DifficultWordsIndex = DifficultWordsIndex + 1
        ElseIf Length < 7 And Length >= 5 Then
            NormalWords(NormalWordsIndex) = Words(i)
            NormalWordsIndex = NormalWordsIndex + 1
        ElseIf Length < 5 Then
            EasyWords(EasyWordsIndex) = Words(i)
            EasyWordsIndex = EasyWordsIndex + 1
        End If
        
    Next
    
End Sub

Public Function ReturnWords(ByVal WordDifficulty As Difficulty) As String
    Dim PickedWord As String
    
    If RecentlyPickedWordsIndex >= 3 Then
        Call Randomize
        RecentlyPickedWordsIndex = 0
    Else
        If WordDifficulty = Easy Then
            PickedWord = EasyWords(Rnd())
            RecentlyPickedWords(RecentlyPickedWordsIndex) = PickedWord
        ElseIf WordDifficulty = Normal Then
            PickedWord = NormalWords(Rnd())
            RecentlyPickedWords(RecentlyPickedWordsIndex) = PickedWord
        ElseIf WordDifficulty = Difficult Then
            PickedWord = DifficultWords(Rnd())
            RecentlyPickedWords(RecentlyPickedWordsIndex) = PickedWord
        End If
        RecentlyPickedWordsIndex = RecentlyPickedWordsIndex + 1
    End If
    
    ReturnWords = PickedWord
End Function

