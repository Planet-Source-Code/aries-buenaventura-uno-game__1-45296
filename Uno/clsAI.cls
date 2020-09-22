VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Easy As New clsEasy
Private Normal As New clsNormal
Private Difficult As New clsDifficult

Public Function ComputerMove(Player As Object, ByVal CardNum As Integer, _
    ByVal Color As Integer)
    
    Dim SelectedCard As Integer
    
    If UCase$(Difficulty) = "EASY" Then
        ' The computer will play in easy mode
        SelectedCard = Easy.Move(Player, CardNum, Color)
    ElseIf UCase$(Difficulty = "NORMAL") Then
        ' The computer will play in normal mode
        SelectedCard = Normal.Move(Player, CardNum, Color)
    ElseIf UCase$(Difficulty = "DIFFICULT") Then
        ' The computer will play in difficult mode
        SelectedCard = Difficult.Move(Player, CardNum, Color)
    End If
    
    ComputerMove = SelectedCard ' get computer move
End Function
