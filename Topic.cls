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

Private RecentlyPickedWord(4) As String
Private RecentlyPickedWordIndex As Integer


' Put an array of 12 in, any more are just ignored
Public Sub Init(ByRef Words() As String)

    Dim EasyWordsIndex As Integer
    Dim NormalWordsIndex As Integer
    Dim DifficultWordsIndex As Integer
    
    EasyWordsIndex = 0
    NormalWordsIndex = 0
    DifficultWordsIndex = 0
    RecentlyPickedWordIndex = 0
    
    ' Sort the words based on their lengths into their respective categories
    For i = 0 To 11
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

Public Function ReturnWords(ByVal WordDifficulty As DifficultyEnum) As String
    Dim PickedWord As String
    
    If RecentlyPickedWordsIndex >= 3 Then
        Call Randomize
        RecentlyPickedWordsIndex = 0
    Else
        Dim FoundWord As Boolean
        FoundWord = False
        If WordDifficulty = Easy Then
        
            ' Loop until a word is found that was not recently picked
            While Not FoundWord
                PickedWord = EasyWords(Rnd())
                
                Dim WordInRecent As Boolean
                WordInRecent = False
                
                For i = 0 To 4
                    If RecentlyPickedWords(i) = PickedWord Then
                        WordInRecent = True
                    End If
                Next
                
                If Not WordInRecent Then
                    FoundWord = True
                End If
            Loop
            RecentlyPickedWords(RecentlyPickedWordsIndex) = PickedWord
            
        ElseIf WordDifficulty = Normal Then
        
            ' Loop until a word is found that was not recently picked
            While Not FoundWord
                PickedWord = NormalWords(Rnd())
                
                Dim WordInRecent As Boolean
                WordInRecent = False
                
                For i = 0 To 4
                    If RecentlyPickedWords(i) = PickedWord Then
                        WordInRecent = True
                    End If
                Next
                
                If Not WordInRecent Then
                    FoundWord = True
                End If
            Loop
            RecentlyPickedWords(RecentlyPickedWordsIndex) = PickedWord
            
        ElseIf WordDifficulty = Difficult Then
        
            ' Loop until a word is found that was not recently picked
            While Not FoundWord
                PickedWord = DifficultWords(Rnd())
                
                Dim WordInRecent As Boolean
                WordInRecent = False
                
                For i = 0 To 4
                    If RecentlyPickedWords(i) = PickedWord Then
                        WordInRecent = True
                    End If
                Next
                
                If Not WordInRecent Then
                    FoundWord = True
                End If
            Loop
            RecentlyPickedWords(RecentlyPickedWordsIndex) = PickedWord
        End If
        RecentlyPickedWordsIndex = RecentlyPickedWordsIndex + 1
        
    End If
    
    ReturnWords = PickedWord
End Function

