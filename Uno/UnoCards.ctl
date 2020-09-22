VERSION 5.00
Begin VB.UserControl UnoCard 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   ScaleHeight     =   1425
   ScaleWidth      =   1095
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "UnoCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Const m_def_CardColor = 0
Const m_def_CardHide = False
Const m_def_CardType = 0
Const m_def_Deck = 0
Const m_def_Points = 0
Const m_def_Repaint = True

Dim m_CardColor As CardColorConstants
Dim m_CardHide As Boolean
Dim m_CardType As CardTypeConstants
Dim m_Deck As DeckConstants
Dim m_Points As Integer
Dim m_Repaint As Boolean

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Enum CardColorConstants
    [Blue]
    [Red]
    [Green]
    [Yellow]
End Enum

Public Enum CardTypeConstants
    [Zero]
    [One]
    [Two]
    [Three]
    [Four]
    [Five]
    [Six]
    [Seven]
    [Eight]
    [Nine]
    [Draw Two]
    [Reverse]
    [Skip]
    [Wild]
    [Wild Draw Four]
End Enum

Public Enum DeckConstants
    [Default]
    [Deck 1]
    [Deck 2]
    [Deck 3]
End Enum

Event Click()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Dim PicSave As StdPicture
Dim twpX As Integer, twpY As Integer
Dim PicWidth As Integer, PicHeight As Integer

Public Property Get CardColor() As CardColorConstants
    CardColor = m_CardColor
End Property

Public Property Let CardColor(ByVal New_CardColor As CardColorConstants)
    m_CardColor = New_CardColor
    PropertyChanged "CardColor"
    If Not CardHide Then Call DrawCard
End Property

Public Property Get CardHide() As Boolean
    CardHide = m_CardHide
End Property

Public Property Let CardHide(ByVal New_CardHide As Boolean)
    m_CardHide = New_CardHide
    PropertyChanged "CardHide"
    Call DrawCard
End Property

Public Property Get CardType() As CardTypeConstants
    CardType = m_CardType
End Property

Public Property Let CardType(ByVal New_CardType As CardTypeConstants)
    m_CardType = New_CardType
    PropertyChanged "CardType"
    If Not CardHide Then Call DrawCard
End Property

Public Property Get Deck() As DeckConstants
    Deck = m_Deck
End Property

Public Property Let Deck(ByVal New_Deck As DeckConstants)
    m_Deck = New_Deck
    PropertyChanged "Deck"
    If CardHide Then Call DrawCard
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
    Set PicSave = UserControl.Picture
End Property

Public Property Get Repaint() As Boolean
    Repaint = m_Repaint
End Property

Public Property Let Repaint(ByVal New_Repaint As Boolean)
    m_Repaint = New_Repaint
    PropertyChanged "Enabled"
End Property

Public Property Get Points() As Integer
    Points = m_Points
End Property

Public Property Let Points(ByVal New_Points As Integer)
    m_Points = New_Points
    PropertyChanged "Points"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    Dim BM As BITMAP
    
    ' Get picture information
    Call GetObject(LoadResPicture(300, vbResBitmap), Len(BM), BM)
        
    twpX = Screen.TwipsPerPixelX
    twpY = Screen.TwipsPerPixelY
    PicWidth = BM.bmWidth: PicHeight = BM.bmHeight
    picTemp.Width = PicWidth * twpX
    picTemp.Height = PicHeight * twpY
    Set PicSave = Nothing
End Sub

Private Sub UserControl_InitProperties()
    m_CardColor = m_def_CardColor
    m_CardHide = m_def_CardHide
    m_CardType = m_def_CardType
    m_Points = m_def_Points
    m_Repaint = m_def_Repaint
End Sub

Private Sub UserControl_LostFocus()
    Set UserControl.Picture = PicSave
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set PicSave = UserControl.Picture
    Call CardPress
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static flutterUp As Boolean
    Static flutterDn As Boolean
    
    If Button And vbLeftButton Then
        If X < 0 Or X > UserControl.Width Or Y < 0 Or Y > UserControl.Height Then
            If Not flutterUp Then
                flutterUp = True
                Set UserControl.Picture = PicSave
            End If
            flutterDn = False
        Else
            If Not flutterDn Then
                flutterDn = True
                Call CardPress
            End If
            flutterUp = False
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set UserControl.Picture = PicSave
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_CardColor = PropBag.ReadProperty("CardColor", m_def_CardColor)
    m_CardHide = PropBag.ReadProperty("CardHide", m_def_CardHide)
    m_CardType = PropBag.ReadProperty("CardType", m_def_CardType)
    m_Deck = PropBag.ReadProperty("Deck", m_def_Deck)
    m_Points = PropBag.ReadProperty("Points", m_def_Points)
    m_Repaint = PropBag.ReadProperty("Repaint", m_def_Repaint)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = PicWidth * twpX
    UserControl.Height = PicHeight * twpY
End Sub

Private Sub UserControl_Show()
    Call DrawCard
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("CardColor", m_CardColor, m_def_CardColor)
    Call PropBag.WriteProperty("CardHide", m_CardHide, m_def_CardHide)
    Call PropBag.WriteProperty("CardType", m_CardType, m_def_CardHide)
    Call PropBag.WriteProperty("Deck", m_Deck, m_def_Deck)
    Call PropBag.WriteProperty("Points", m_Points, m_def_Points)
    Call PropBag.WriteProperty("Repaint", m_Repaint, m_def_Repaint)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

Private Sub CardPress()
    Const SRCAND = &H8800C6
    
    If PicSave Then
        picTemp.Cls
        Set picTemp.Picture = Nothing
        picTemp.BackColor = &HE0E0E0
        picTemp.PaintPicture PicSave, 0, 0, , , , , , , SRCAND
        Set picTemp.Picture = picTemp.Image
        Set UserControl.Picture = picTemp.Picture
    End If
End Sub

Private Sub DrawCard()
    If Repaint Then
        If Not CardHide Then
            If CardType < 13 Then
                Dim resID As Integer, s As String
                Dim xmid As Integer, ymid As Integer
                
                ' Draw the card
                resID = IIf(CardHide, 999, CardColor + 100)
                
                UserControl.AutoRedraw = True
                Set UserControl.Picture = LoadResPicture(resID, vbResBitmap)
                s = GetCardType(CardType)
                UserControl.FontName = "Arial Narrow"
                UserControl.FontSize = 8
                Call PutText(5 * twpX, 1.5 * twpY, s, IIf(CardColor = Blue, vbWhite, vbBlack))
                UserControl.FontName = "Arial Black"
                UserControl.FontSize = 30
                s = Left$(s, 1)
                xmid = (UserControl.ScaleWidth - UserControl.TextWidth(s)) / 2
                ymid = (UserControl.ScaleHeight - UserControl.TextHeight(s)) / 2
                Call OutlineText(xmid, ymid, s, 1, GetCardColor(CardColor))
                Set UserControl.Picture = UserControl.Image
                UserControl.AutoRedraw = False
            Else
                Set UserControl.Picture = LoadResPicture(CardType Mod 13 + 200, vbResBitmap)
            End If
        Else
            Set UserControl.Picture = LoadResPicture(300 + Deck, vbResBitmap)
        End If
    End If
    
    ' Card score
    Select Case CardType
    Case [Zero] To [Nine]
        Points = CardType
    Case [Draw Two], [Reverse], [Skip]
        Points = 20
    Case Else   ' Wild and Wild Draw Four
        Points = 50
    End Select
End Sub

Private Function GetCardColor(Color As Integer)
    Dim c As Long
    
    Select Case Color
    Case Is = 0: c = vbBlue
    Case Is = 1: c = vbRed
    Case Is = 2: c = vbGreen
    Case Is = 3: c = vbYellow
    End Select
        
    GetCardColor = c
End Function

Private Function GetCardType(id As Integer)
    Dim s As String
    
    Select Case id
    Case 0 To 9: s = id
    Case Is = 10: s = "Draw Two"
    Case Is = 11: s = "Reverse"
    Case Is = 12: s = "Skip"
    End Select
    
    GetCardType = s
End Function

Private Sub OutlineText(X As Integer, Y As Integer, text As String, weight As Integer, Optional Color As Long = vbBlack)
    Dim curx As Integer, cury As Integer, i As Integer
    
    ' Create outline text
    For i = -weight To weight
        curx = X - i * twpX: cury = Y
        Call PutText(curx, cury, text)
        curx = X: cury = Y - i * twpY
        Call PutText(curx, cury, text)
        curx = X - weight * twpX: cury = Y - i * twpY
        Call PutText(curx, cury, text)
        curx = X + weight * twpX: cury = Y - i * twpY
        Call PutText(curx, cury, text)
    Next i
    
    Call PutText(X, Y, text, Color)
End Sub

Private Sub PutText(X As Integer, Y As Integer, text As String, Optional Color As Long = vbBlack)
    UserControl.ForeColor = Color
    UserControl.CurrentX = X
    UserControl.CurrentY = Y
    UserControl.Print text
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub


