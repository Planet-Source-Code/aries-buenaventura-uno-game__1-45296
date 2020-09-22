VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Move(Player As Object, ByVal CardNum As Integer, ByVal Color As Integer) As Integer
    Dim mov As Integer
    
    mov = -1
    mov = SearchWildDraw(Player, CardNum, Color)
    If mov = -1 Then mov = SearchDrawTwo(Player, CardNum, Color)
    If mov = -1 Then mov = SearchMove(Player, CardNum, Color)
    If mov = -1 Then mov = SearchWild(Player, CardNum, Color)
    
    If mov <> -1 Then
        If (Player(mov).CardType = Wild) Or (Player(mov).CardType = [Wild Draw Four]) Then
            Dim cc As Integer
    
            If Player.Count - 1 = 1 Then
                cc = Int(4 * Rnd)
            Else
                cc = GetLargestColor(Player)
            End If
    
            Player(mov).CardColor = cc
        End If
    End If
    
    Move = mov
End Function

