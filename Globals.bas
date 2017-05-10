Attribute VB_Name = "Module1"
' Enum to represent the game state
Enum GameStateType
    Title
    Normal
    Difficult
    EndScreen
End Enum

' Pass into topic class to get word of that difficulty
Enum Difficulty
    Easy
    Normal
    Difficult
End Enum

' Enum to represent which topic
Enum TopicType
    None
    CS
    Math
    Chemistry
End Enum


' Global variables which represent the game state and topic respectively
Public GameState As GameStateType
Public Topic As TopicType


' RUN THIS
Public Sub Init()
    GameState = Title
End Sub


