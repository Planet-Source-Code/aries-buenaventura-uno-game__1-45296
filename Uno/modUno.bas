Attribute VB_Name = "modUno"
Option Explicit

Public Const SND_PICK = 101
Public Const SND_THROW = 102

Public Const TOTAL_CARDS = 108
Public Const TOTAL_COLORS = 4
Public Const CARDS_PER_PLAYER = 7
Public Const MAX_PLAYERS = 4
    
Public AnimationType As Integer
Public AutoSort As Boolean
Public CardDeck As Integer
Public Difficulty As String
Public Opponents As Integer
Public PlayerName(4) As Variant
Public PlayerScore(4) As Integer
Public WaveExpression As String

Public NewGame As Boolean

Public Function CountCard(Player As Object, ByVal Compare As Integer, ByVal opt As Integer) As Integer
    Dim Card As Object, sum As Integer
    
    For Each Card In Player
        If Card.Index <> 0 Then
            If (Card.CardType <> [Wild]) And (Card.CardType <> [Wild Draw Four]) Then
                If opt = 0 Then
                    If Compare = Card.CardType Then sum = sum + 1
                Else
                    If Compare = Card.CardColor Then sum = sum + 1
                End If
            End If
        End If
    Next Card
    
    CountCard = sum
End Function

Public Function GetPlayer(Player As String) As Integer
    Dim nPlayer As Integer
    
    If Player = "PlayerOne" Then
        nPlayer = 0
    ElseIf Player = "PlayerTwo" Then
        nPlayer = 1
    ElseIf Player = "PlayerThree" Then
        nPlayer = 2
    ElseIf Player = "PlayerFour" Then
        nPlayer = 3
    End If
    GetPlayer = nPlayer
End Function

Public Function GetWinner() As Integer
    Dim i As Integer, Least As Integer
    
    Least = PlayerScore(0)
    For i = 0 To Opponents
        Least = IIf(Least > PlayerScore(i), PlayerScore(i), Least)
    Next i
    
    Dim nPlayer As Integer
    For i = 0 To Opponents
        If Least = PlayerScore(i) Then
            nPlayer = i
            Exit For
        End If
    Next i
    
        
    GetWinner = nPlayer
End Function

Public Function TotalCards(F As Form)
    Dim Card As Control, sum As Integer

    sum = sum + 0
    For Each Card In F.Controls
        If TypeName(Card) = "UnoCard" Then
            If Card.Visible And Card.Tag = "AnimCard" Then
                sum = sum + 1
            End If
        End If
    Next Card
    
    TotalCards = sum
End Function

Public Function SearchCard(Cards As Object, _
    CardNum As Integer, Color As Integer, WhatCard As Integer)
    
    Dim Card As Object, j As Integer
    
    j = -1

    For Each Card In Cards
        If Card.Index <> 0 Then
            If Card.CardColor = Color Then
                If Card.CardType = WhatCard Then
                    j = Card.Index
                End If
            End If
        End If
        If j <> -1 Then Exit For
    Next Card
    
    If j = -1 Then
        If CardNum = WhatCard Then
            For Each Card In Cards
                If Card.Index <> 0 Then
                    If Card.CardType = WhatCard Then
                        j = Card.Index
                    End If
                End If
            
                If j <> -1 Then Exit For
            Next Card
        End If
    End If
    
    SearchCard = j
End Function

Public Function SearchMove(Cards As Object, CardNum As Integer, Color As Integer) As Integer
    Dim Card As Object, j As Integer
    
    j = -1
    
    For Each Card In Cards
        If Card.Index <> 0 Then
            If Card.CardType <> Wild And Card.CardType <> [Wild Draw Four] Then
                If Card.CardType = CardNum Then
                    j = Card.Index
                ElseIf Card.CardColor = Color Then
                    j = Card.Index
                End If
            End If
        End If
        
        If j <> -1 Then Exit For
    Next Card
     
    SearchMove = j
End Function

Public Function SearchWild(Cards As Object, CardNum As Integer, Color As Integer) As Integer
    Dim Card As Object, j As Integer
    
    j = -1
    
    For Each Card In Cards
        If Card.Index <> 0 Then
            If Card.CardType = Wild Then
                j = Card.Index
            End If
            If j <> -1 Then Exit For
        End If
    Next Card
    
    SearchWild = j
End Function

Public Function SearchDrawTwo(Cards As Object, CardNum As Integer, Color As Integer) As Integer
    Dim Card As Object, j As Integer
    
    j = -1

    For Each Card In Cards
        If Card.Index <> 0 Then
            If Card.CardType = [Draw Two] Then
                If Card.CardType = CardNum Or Card.CardColor = Color Then
                    j = Card.Index
                End If
            End If
        End If
        If j <> -1 Then Exit For
    Next Card
    
    SearchDrawTwo = j
End Function

Public Function SearchWildDraw(Cards As Object, CardNum As Integer, Color As Integer) As Integer
    Dim Card As Object, j As Integer
    
    j = -1
    
    If CountCard(Cards, Color, 1) = 0 Then
        For Each Card In Cards
            If Card.Index <> 0 Then
                If Card.CardType = [Wild Draw Four] Then
                    j = Card.Index
                    Exit For
                End If
            End If
                
            If j <> -1 Then Exit For
        Next Card
    End If
    
    SearchWildDraw = j
End Function

Public Function GetLargestColor(Player As Object) As Integer
    Dim cb As Integer, cr As Integer
    Dim cg As Integer, cy As Integer
    
    cb = CountCard(Player, [blue], 1)
    cr = CountCard(Player, [Red], 1)
    cg = CountCard(Player, [Green], 1)
    cy = CountCard(Player, [Yellow], 1)
    
    Dim br As Integer, gy As Integer
    Dim lbr As Integer, lgy As Integer
    
    If cb < cr Then
        br = cr
        lbr = [Red]
    Else
        br = cb
        lbr = [blue]
    End If
    
    If cg < cy Then
        gy = cy
        lgy = [Yellow]
    Else
        gy = cg
        lgy = [Green]
    End If
    
    GetLargestColor = IIf(br < gy, lgy, lbr)
End Function

Public Function GetLargestNum(Cards As Object, Color)
    Dim Card As Object, j As Integer, l As Integer
    
    j = -1: l = -1
    
    For Each Card In Cards
        If Card.Index <> 0 Then
            If (Card.CardType >= 0) And Card.CardType <= 9 Then
                If Card.CardColor = Color Then
                    If l = -1 Then
                        l = Card.CardType
                    ElseIf l < Card.CardType Then
                        l = Card.CardType
                    End If
                End If
            End If
        End If
    Next Card
    
    For Each Card In Cards
        If Card.Index <> 0 Then
            If (Card.CardType = l) And (Card.CardColor = Color) Then
                j = Card.Index
            End If
        End If
            
        If j <> -1 Then Exit For
    Next Card
    
    GetLargestNum = j
End Function

Public Sub OpenCards(Cards As Object, ByVal opt As Boolean)
    Dim Card As Object
    
    For Each Card In Cards
        Card.CardHide = Not opt
    Next Card
End Sub

Public Function ChkParen(ByVal Expression As String) As Boolean
    Dim i As Integer, c As String
    Dim nOpenParen As Integer, nCloseParen As Integer
    
    Expression = Trim$(Expression)
    
    For i = 1 To Len(Expression)
        c = Mid$(Expression, i, 1)
        
        If c = "(" Then
            nOpenParen = nOpenParen + 1
        ElseIf c = ")" Then
            nCloseParen = nCloseParen + 1
        End If
    Next i
    
    If nOpenParen <> nCloseParen Then
        Call MsgBox("Unbalanced Parenthesis")
    End If
    
    ChkParen = Not (nOpenParen <> nCloseParen)
End Function





