Attribute VB_Name = "Module1"
' Pass into topic class to get word of that difficulty
Enum DifficultyEnum
    Easy
    Normal
    Hard
End Enum

' Enum to represent which topic
Enum TopicEnum
    None
    CS
    Math
    Chemistry
End Enum


' Global variables which represent the game state and topic respectively
Public DifficultyState As DifficultyEnum
Public TopicState As TopicEnum

' Class which handles word management
Public Topic As Topic


' Loads Form2 and unloads Form1
Public Sub SwapWindow(ByRef Form1 As Form, ByRef Form2 As Form)
    Load Form2
    Form1.Hide
    Form2.Show
    Unload Form1
End Sub


