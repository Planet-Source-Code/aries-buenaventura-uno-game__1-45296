VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DEFAULT_SPEED = 4

Public BounceRightEdge As Integer
Public BounceBottomEdge As Integer

Public RotXradius As Integer
Public RotYradius As Integer

Public WaveRightEdge As Integer
Public WaveBottomEdge As Integer

Public Speed As Integer, Play As Boolean

Public Sub Bounce(F As Form)
    Dim PicWidth As Integer, PicHeight As Integer
    Dim NewX As Single, NewY As Single, temp As Single
    Dim dx As Integer, dy As Integer, X As Integer, Y As Integer
    Dim Card As Control, switch As Boolean
    
    On Error Resume Next
    
    switch = False
    dx = 1 * (Speed - 9): dy = 1 * (Speed - 9)
    dx = dx + 10: dy = dy + 10
    Do While Play
        For Each Card In F.Controls
            If Not Play Then Exit For
            If TypeName(Card) = "UnoCard" Then
                If Card.Visible And Card.Tag = "AnimCard" Then
                    If Card.CardHide Then Card.CardHide = False
                    PicWidth = Card.Width
                    PicHeight = Card.Height
                    temp = X + dx
                    If temp + PicWidth > BounceRightEdge Then
                        dx = -Abs(dx)
                    ElseIf temp < 0 Then
                        dx = Abs(dx)
                    End If
    
                    NewX = X + dx
        
                    temp = Y + dy
                    If temp + PicHeight > BounceBottomEdge Then
                        dy = -Abs(dy)
                    ElseIf temp < 0 Then
                        dy = Abs(dy)
                    End If
                    NewY = Y + dy
                
                    X = NewX
                    Y = NewY
                    
                    If Not switch Then
                        Call Linear(Card, Card.Left, Card.Top, X, Y, _
                            IIf(Speed < 5, 5, Speed))
                        Card.Move X, Y
                    Else
                        Card.Move X, Y
                    End If
                    Card.ZOrder 0
                    F.Refresh
                End If
            End If
        Next Card
        DoEvents
        If Not switch Then switch = True
    Loop
End Sub

Private Sub Class_Initialize()
    Play = False
    Speed = DEFAULT_SPEED
End Sub

Public Sub Linear(cBox As Control, ByVal X1 As Integer, ByVal Y1 As Integer, _
    ByVal X2 As Integer, ByVal Y2 As Integer, Optional ByVal Rate As Integer = 0)
        
    On Error Resume Next
    
    If Speed + Rate = 0 Then Exit Sub
    Dim i As Single, X As Single, Y As Single
    Dim px As Single, py As Single
    Dim dx As Single, dy As Single
    Dim sdx As Single, sdy As Single
    Dim dxabs As Single, dyabs As Single
    
    X = 0: Y = 0
    px = X1: py = Y1
    dx = X2 - X1: dy = Y2 - Y1
    sdx = Sgn(dx): sdy = Sgn(dy)
    dxabs = Abs(dx): dyabs = Abs(dy)
    Rate = Abs(Speed + Rate)

    If dxabs >= dyabs Then
        For i = 0 To dxabs Step Rate
            Y = Y + dyabs
            If Y >= dxabs Then
                Y = Y - dxabs
                py = py + sdy * Rate
            End If
            cBox.ZOrder 0
            cBox.Move px, py
            px = px + sdx * Rate
            DoEvents
        Next i
    Else
        For i = 0 To dyabs Step Rate
            X = X + dxabs
            If X >= dyabs Then
                X = X - dyabs
                px = px + sdx * Rate
            End If
            cBox.ZOrder 0
            cBox.Move px, py
            py = py + sdy * Rate
            DoEvents
        Next i
    End If
End Sub

Public Sub Rotation(F As Form)
    Dim switch As Boolean
    Dim Card As Control, deg As Single
    Dim xloc As Single, yloc As Single

    On Error Resume Next
    
    deg = 1: switch = False
    
    Do While Play
        For Each Card In F.Controls
            If Not Play Then Exit For
            If TypeName(Card) = "UnoCard" Then
                If Card.Visible And Card.Tag = "AnimCard" Then
                    If Card.CardHide Then Card.CardHide = False
        
                    xloc = RotXradius + Cos(Rads(deg)) * RotXradius
                    yloc = RotYradius + Sin(Rads(deg)) * RotYradius
                    
                    If Not switch Then
                        Call Linear(Card, Card.Left, Card.Top, xloc, yloc, _
                            IIf(Speed < 5, 5, Speed))
                        Call Card.Move(xloc, yloc)
                    Else
                        Call Card.Move(xloc, yloc)
                    End If
                    Card.ZOrder 0
                    F.Refresh
                    deg = deg + Speed * 0.5
                    If deg > 360 Then deg = 0
                End If
            End If
        Next Card
        DoEvents
        If Not switch Then switch = True
    Loop
End Sub

Public Sub Wave(F As Form, ByVal formula As String, _
    CtrlWidth As Integer, CtrlHeight As Integer)
    
    Dim w As Integer, h As Integer
    Dim X As Integer, Y As Integer
    Dim xmid As Integer, ymid As Integer
    
    Dim Card As Control, Crest As Single
    Dim switch As Boolean, switch1 As Boolean, deg As Single
    
    On Error Resume Next
    w = WaveRightEdge
    h = WaveBottomEdge
    xmid = (w - CtrlWidth) / 2
    ymid = (h - CtrlHeight) / 2
    deg = -xmid + CtrlWidth / 2
    Crest = ymid / 10
    Do While Play
        For Each Card In F.Controls
            If Not Play Then Exit For
            If TypeName(Card) = "UnoCard" Then
                If Card.Visible And Card.Tag = "AnimCard" Then
                    If Card.CardHide Then Card.CardHide = False
                    If w <> WaveRightEdge Or h <> WaveBottomEdge Then
                        w = WaveRightEdge
                        h = WaveBottomEdge
                        xmid = (w - Card.Width) / 2
                        ymid = (h - Card.Height) / 2
                    End If
                    
                    Call Script.ExecuteStatement("x=" & Rads(deg))
                    X = deg: Y = Crest * Script.Eval(formula)
                    
                    If Not switch1 Then
                        Call Linear(Card, Card.Left, Card.Top, xmid + X, ymid - Y, _
                            IIf(Speed < 5, 5, Speed))
                        Card.Move xmid + X, ymid - Y
                    Else
                        Call Card.Move(xmid + X, ymid - Y)
                    End If
                    
                    Call Card.ZOrder(0)
                    F.Refresh
                    If Not switch Then
                        deg = deg + Speed
                        switch = IIf(Card.Left + Card.Width > w, True, False)
                    Else
                        deg = deg - Speed
                        switch = IIf(Card.Left < 0, False, True)
                    End If
                End If
            End If
        Next Card
        DoEvents
        If Not switch1 Then switch1 = True
    Loop
End Sub


