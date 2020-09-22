VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEasy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Move(Player As Object, ByVal CardNum As Integer, ByVal Color As Integer) As Integer
    Dim mov As Integer
        
    mov = -1
    mov = SearchMove(Player, CardNum, Color)
    If mov = -1 Then mov = SearchWild(Player, CardNum, Color)
    If mov = -1 Then mov = SearchWildDraw(Player, CardNum, Color)
    
    Move = mov
End Function

