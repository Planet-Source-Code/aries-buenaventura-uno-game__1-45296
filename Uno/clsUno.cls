VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type CardInfo
    UnoCardType As Integer
    UnoCardColor As Integer
    UnoPicture As StdPicture
End Type

Public AI As New clsAI
Public DrawPileCards As Collection
Public DiscardPileCards As New Collection

Public Sub Initialize()
    Dim i As Integer, Card As Object, dd As New Collection
        
    Set DrawPileCards = Nothing
    Set DiscardPileCards = Nothing
    
    For i = 1 To TOTAL_CARDS
        Set Card = New clsCard
        Card.UnoCardType = i
        Card.UnoCardColor = i Mod TOTAL_COLORS
        Call dd.Add(Card)
    Next i
        
    Dim temp As New Collection, CardNumber As Integer
    
    Set temp = Shuffle(dd)
    For i = 1 To temp.Count
        CardNumber = temp(i).UnoCardType Mod 27
        If CardNumber > 18 Then
            Select Case CardNumber
            Case 19 To 20 ' Draw Two
                CardNumber = 10
            Case 21 To 22 ' Reverse
                CardNumber = 11
            Case 23 To 24 ' Skip
                CardNumber = 12
            Case 25       ' Wild
                CardNumber = 13
            Case 26       ' Wild Draw Four
                CardNumber = 14
            End Select
        Else
            CardNumber = CardNumber Mod 10
        End If
        
        dd(i).UnoCardType = CardNumber
        dd(i).UnoCardColor = temp(i).UnoCardColor
    Next i
    
    Set DrawPileCards = dd
End Sub

Public Function ChkMove(ByRef Player As Object, ByVal Index As Integer, _
    previd As Integer, prevcc As Integer) As Boolean
    
    Dim id As Integer, cc As Integer, IsOk As Boolean
        
    IsOk = False
    id = Player(Index).CardType
    cc = Player(Index).CardColor
    
    If (id <> [Wild]) And (id <> [Wild Draw Four]) Then
        If Not IsOk Then IsOk = IIf(id = previd, True, False)
        If Not IsOk Then IsOk = IIf(cc = prevcc, True, False)
    Else
        If Not IsOk Then IsOk = IIf(id = [Wild], True, False)
        If Not IsOk Then IsOk = IIf(CountCard(Player, prevcc, 1) = 0, True, False)
    End If
    
    ChkMove = IsOk
End Function

Public Sub Pick(ByRef Player As Object, ByVal Count As Integer)
    If DrawPileCards.Count > 0 Then
        Dim i As Integer
        
        For i = 1 To Count
            Call CreateObject(Player, 1)
            Player(Player.Count - 1).Visible = False
            Player(Player.Count - 1).CardType = _
                DrawPileCards(DrawPileCards.Count).UnoCardType
            Player(Player.Count - 1).CardColor = _
                DrawPileCards(DrawPileCards.Count).UnoCardColor
            DrawPileCards.Remove DrawPileCards.Count
            Player(Player.Count - 1).Visible = True
        Next i
    End If
End Sub

Public Function Shuffle(ByRef DiscardPileCards As Collection) As Collection
    Dim temp As Object, dw As Collection
    Dim dd As Collection, IsMatch As Boolean
    Dim i As Integer, n As Integer
    
    Set dw = New Collection
    Set dd = DiscardPileCards
    
    While dd.Count <> dw.Count
        Randomize
        n = Int((dd.Count * Rnd) + 1)
        
        If dw.Count = 0 Then
            Set temp = New clsCard
            temp.UnoCardType = dd(n).UnoCardType
            temp.UnoCardColor = dd(n).UnoCardColor
            Call dw.Add(temp)
        Else
            IsMatch = False
            
            For i = 1 To dw.Count
                If n = dw(i).UnoCardType Then
                    IsMatch = True
                    Exit For
                End If
            Next i
            
            If Not IsMatch Then
                Set temp = New clsCard
                temp.UnoCardType = dd(n).UnoCardType
                temp.UnoCardColor = dd(n).UnoCardColor
                Call dw.Add(temp)
            End If
        End If
    Wend

    Set Shuffle = dw
End Function

Public Sub Throw(ByRef Player As Object, ByVal Index As Integer)
    Dim dd As Object
    
    On Error Resume Next
    Set dd = New clsCard
    
    dd.UnoCardType = Player(Index).CardType
    dd.UnoCardColor = Player(Index).CardColor
    Call DiscardPileCards.Add(dd)
    
    Dim i As Integer
    
    Player(Player.Count - 1).Visible = False
    
    For i = Index To Player.Count - 2
        Player(i).Repaint = False
        Player(i).CardType = Player(i + 1).CardType
        Player(i).CardColor = Player(i + 1).CardColor
        Set Player(i).Picture = Player(i + 1).Picture
        Player(i).Repaint = True
    Next i
    
    DestroyObject Player, 1
End Sub

Public Function TotalPoints(ByRef Player As Object) As Integer
    Dim Card As Object, Total As Integer
    
    Total = 0
    For Each Card In Player
        If Card.Visible Then
            Total = Total + Card.Points
        End If
    Next Card
    
    TotalPoints = Total
End Function

Public Sub SortCards(Cards As Object)
    Dim Card As Object
    
    Dim oBlue As Object, cBlue As New Collection
    For Each Card In Cards
        If Card.Index <> 0 Then
            If (Card.CardType <> 13) And (Card.CardType <> 14) Then
                If Card.CardColor = [blue] Then
                    Set oBlue = New clsCard
                    oBlue.UnoCardType = Card.CardType
                    oBlue.UnoCardColor = Card.CardColor
                    cBlue.Add oBlue
                End If
            End If
        End If
    Next Card
            
    Dim oRed As Object, cRed As New Collection
    For Each Card In Cards
        If Card.Index <> 0 Then
            If (Card.CardType <> 13) And (Card.CardType <> 14) Then
                If Card.CardColor = [Red] Then
                    Set oRed = New clsCard
                    oRed.UnoCardType = Card.CardType
                    oRed.UnoCardColor = Card.CardColor
                    cRed.Add oRed
                End If
            End If
        End If
    Next Card
    
    Dim oGreen As Object, cGreen As New Collection
    For Each Card In Cards
        If Card.Index <> 0 Then
            If (Card.CardType <> 13) And (Card.CardType <> 14) Then
                If Card.CardColor = [Green] Then
                    Set oGreen = New clsCard
                    oGreen.UnoCardType = Card.CardType
                    oGreen.UnoCardColor = Card.CardColor
                    cGreen.Add oGreen
                End If
            End If
        End If
    Next Card
    
    Dim oYellow As Object, cYellow As New Collection
    For Each Card In Cards
        If Card.Index <> 0 Then
            If (Card.CardType <> 13) And (Card.CardType <> 14) Then
                If Card.CardColor = [Yellow] Then
                    Set oYellow = New clsCard
                    oYellow.UnoCardType = Card.CardType
                    oYellow.UnoCardColor = Card.CardColor
                    cYellow.Add oYellow
                End If
            End If
        End If
    Next Card
    
    Dim oWildCard As Object, cWildCard As New Collection
    For Each Card In Cards
        If Card.Index <> 0 Then
            If (Card.CardType = 13) Or (Card.CardType = 14) Then
                Set oWildCard = New clsCard
                oWildCard.UnoCardType = Card.CardType
                oWildCard.UnoCardColor = Card.CardColor
                cWildCard.Add oWildCard
            End If
        End If
    Next Card
    
    Call Sort(cBlue)
    Call Sort(cRed)
    Call Sort(cGreen)
    Call Sort(cYellow)
    Call Sort(cWildCard)
    
    Dim temp As Object, allCards As New Collection
    
    For Each Card In cBlue
        Set temp = New clsCard
        temp.UnoCardType = Card.UnoCardType
        temp.UnoCardColor = Card.UnoCardColor
        allCards.Add temp
    Next Card
    
    For Each Card In cRed
        Set temp = New clsCard
        temp.UnoCardType = Card.UnoCardType
        temp.UnoCardColor = Card.UnoCardColor
        allCards.Add temp
    Next Card
    
    For Each Card In cGreen
        Set temp = New clsCard
        temp.UnoCardType = Card.UnoCardType
        temp.UnoCardColor = Card.UnoCardColor
        allCards.Add temp
    Next Card
    
    For Each Card In cYellow
        Set temp = New clsCard
        temp.UnoCardType = Card.UnoCardType
        temp.UnoCardColor = Card.UnoCardColor
        allCards.Add temp
    Next Card
    
    For Each Card In cWildCard
        Set temp = New clsCard
        temp.UnoCardType = Card.UnoCardType
        temp.UnoCardColor = Card.UnoCardColor
        allCards.Add temp
    Next Card
    
    Dim ci() As CardInfo
    Dim i As Integer, step As Integer
    ReDim ci(Cards.Count - 1) As CardInfo
    For i = LBound(ci()) To UBound(ci())
        ci(i).UnoCardType = Cards(i).CardType
        ci(i).UnoCardColor = Cards(i).CardColor
        Set ci(i).UnoPicture = Cards(i).Picture
    Next i
    
    For i = 1 To Cards.Count - 1
        If Cards(i).CardType <> allCards(i).UnoCardType Or _
            Cards(i).CardColor <> allCards(i).UnoCardColor Then
                For step = 0 To allCards.Count
                    If allCards(i).UnoCardType = ci(step).UnoCardType And _
                        allCards(i).UnoCardColor = ci(step).UnoCardColor Then
                            Cards(i).Repaint = False
                            Cards(i).CardType = ci(step).UnoCardType
                            Cards(i).CardColor = ci(step).UnoCardColor
                            Set Cards(i).Picture = ci(step).UnoPicture
                            Cards(i).Repaint = True
                            Exit For
                    End If
                Next step
        End If
    Next i
End Sub

Private Sub Sort(Cards As Object)
    Dim i As Integer, j As Integer, z As Integer
    
    For i = 1 To Cards.Count
        For j = 1 To Cards.Count - 1
            If Cards(j).UnoCardType > Cards(j + 1).UnoCardType Then
                z = Cards(j).UnoCardType
                Cards(j).UnoCardType = Cards(j + 1).UnoCardType
                Cards(j + 1).UnoCardType = z
            End If
        Next j
    Next i
End Sub
