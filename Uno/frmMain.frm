VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Uno Game"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9465
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6090
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "6/1/2003"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "6:20 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTray 
      BorderStyle     =   0  'None
      Height          =   6075
      Left            =   7140
      TabIndex        =   17
      Top             =   0
      Width           =   2295
      Begin MSComctlLib.ListView lstPlayers 
         Height          =   1515
         Left            =   0
         TabIndex        =   19
         Top             =   4500
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2672
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstHistory 
         Height          =   4155
         Left            =   0
         TabIndex        =   18
         Top             =   60
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   7329
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Round : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   20
         Top             =   4320
         Width           =   750
      End
      Begin VB.Label lblRound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   720
         TabIndex        =   21
         Top             =   4320
         Width           =   120
      End
   End
   Begin VB.PictureBox picTable 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Height          =   6075
      Left            =   0
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   472
      TabIndex        =   1
      Top             =   0
      Width           =   7140
      Begin Uno.UnoCard PlayerFour 
         Height          =   1350
         Index           =   0
         Left            =   2880
         TabIndex        =   6
         Tag             =   "AnimCard"
         Top             =   600
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   2381
         Enabled         =   0   'False
         CardHide        =   -1  'True
         Picture         =   "frmMain.frx":030A
      End
      Begin VB.PictureBox PicColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   900
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   27
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   60
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   27
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   480
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   27
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin Uno.UnoCard PlayerThree 
         Height          =   1350
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         Tag             =   "AnimCard"
         Top             =   600
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   2381
         Enabled         =   0   'False
         CardHide        =   -1  'True
         Picture         =   "frmMain.frx":42A4
      End
      Begin Uno.UnoCard PlayerTwo 
         Height          =   1350
         Index           =   0
         Left            =   960
         TabIndex        =   7
         Tag             =   "AnimCard"
         Top             =   600
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   2381
         Enabled         =   0   'False
         CardHide        =   -1  'True
         Picture         =   "frmMain.frx":823E
      End
      Begin Uno.UnoCard PlayerOne 
         Height          =   1350
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Tag             =   "AnimCard"
         Top             =   600
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   2381
         MousePointer    =   99
         Picture         =   "frmMain.frx":C1D8
      End
      Begin Uno.UnoCard DrawPile 
         Height          =   1350
         Left            =   0
         TabIndex        =   9
         Top             =   1980
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   2381
         CardHide        =   -1  'True
         MousePointer    =   99
         Picture         =   "frmMain.frx":10172
      End
      Begin Uno.UnoCard DiscardPile 
         Height          =   1350
         Left            =   960
         TabIndex        =   10
         Top             =   1980
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   2381
         Enabled         =   0   'False
         CardHide        =   -1  'True
         Picture         =   "frmMain.frx":1410C
      End
      Begin Uno.UnoCard Dummy 
         Height          =   1350
         Index           =   0
         Left            =   1920
         TabIndex        =   11
         Top             =   1980
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   2381
         Enabled         =   0   'False
         CardHide        =   -1  'True
         Picture         =   "frmMain.frx":180A6
      End
      Begin Uno.UnoCard Dummy 
         Height          =   1350
         Index           =   1
         Left            =   2880
         TabIndex        =   12
         Top             =   1980
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   2381
         Enabled         =   0   'False
         CardHide        =   -1  'True
         Picture         =   "frmMain.frx":1C040
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   16
         Top             =   3660
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   1980
         TabIndex        =   15
         Top             =   3420
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1980
         TabIndex        =   14
         Top             =   3900
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   1980
         TabIndex        =   13
         Top             =   4140
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Shape shpCard 
         FillStyle       =   7  'Diagonal Cross
         Height          =   1215
         Index           =   1
         Left            =   1080
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   6
         FillColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         Shape           =   2  'Oval
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape shpCard 
         FillStyle       =   7  'Diagonal Cross
         Height          =   1215
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuGameReset 
         Caption         =   "Reset"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGameDeal 
         Caption         =   "&Deal"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuGameBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameDemo 
         Caption         =   "Dem&o"
      End
      Begin VB.Menu mnuGameCheat 
         Caption         =   "&Cheat"
      End
      Begin VB.Menu mnuGameBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Uno Game"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Uno As New clsUno

Dim Done As Boolean
Dim Reset As Boolean, Deal As Boolean
Dim Rotation As Integer, StopGame As Boolean
Dim NextPlayer As Object, SplitCards As Integer
Dim SelectedAnimation As Integer

Private Sub AlignCards(Player As Object, Optional Animate As Boolean = False)
    Dim cW As Long, cH As Long
    Dim sw As Long, sH As Long
    Dim i As Integer, nC As Integer
    Dim xpos As Single, ypos As Single
    
    nC = Player.Count - 1
    cW = Player(0).Width: cH = Player(0).Height
    sw = picTable.ScaleWidth: sH = picTable.ScaleHeight
    
    For i = 0 To nC
        If Opponents < 3 Then
            Select Case Player(0).Name
            Case Is = "PlayerOne"
                xpos = (sw - (cW + cW * 0.3 * nC)) / 2 + i * cW * 0.3 - 10
                ypos = sH - cH - cH * 0.05
            Case Is = "PlayerTwo"
                xpos = (sw - (cW + cW * 0.3 * nC)) / 2 + i * cW * 0.3 - 10
                ypos = cH * 0.05
            Case Is = "PlayerThree"
                xpos = (i - 1) * cW * 0.1
                ypos = (sH - cH) / 2
            End Select
        Else
            Select Case Player(0).Name
            Case Is = "PlayerOne"
                xpos = (sw - (cW + cW * 0.3 * nC)) / 2 + i * cW * 0.3 - 10
                ypos = sH - cH - cH * 0.05
            Case Is = "PlayerTwo"
                xpos = sw - cW - (i - 1) * cW * 0.1
                ypos = (sH - cH) / 2
            Case Is = "PlayerThree"
                xpos = (sw - (cW + cW * 0.3 * nC)) / 2 + i * cW * 0.3 - 10
                ypos = cH * 0.05
            Case Is = "PlayerFour"
                xpos = (i - 1) * cW * 0.1
                ypos = (sH - cH) / 2
            End Select
        End If
        If Animate Then
            If i <> 0 Then
                Call BeginPlaySound(SND_PICK, False)
                Call Animated.Linear(Player(i), DrawPile.Left, _
                    DrawPile.Top, xpos, ypos)
                Call EndPlaySound
            End If
        End If
        
        Call Player(i).Move(xpos, ypos)
        Call Player(i).ZOrder(0)
    Next i
End Sub

Private Sub AlignControls()
    Dim cW As Integer, cH As Integer
    Dim sw As Integer, sH As Integer
    Dim xpos As Integer, ypos As Integer
    
    sw = picTable.ScaleWidth: sH = picTable.ScaleHeight
    cW = DrawPile.Width: cH = DrawPile.Height
    
    xpos = (sw - cW) / 2 - cW * 0.55 - 4 - 0.153 * DrawPile.Width
    ypos = (sH - cH) / 2
    
    Call Dummy(0).Move(xpos, ypos)
    Call Dummy(1).Move(xpos + 2, ypos + 1)
    Call DrawPile.Move(xpos + 4, ypos + 2)
    Call Dummy(0).ZOrder(0)
    Call Dummy(1).ZOrder(0)
    Call DrawPile.ZOrder(0)
    
    If Not Deal Then
        If Not Dummy(1).Visible Then
            Call DrawPile.Move(xpos + 2, ypos + 1)
        ElseIf Not Dummy(1).Visible Then
            Call DrawPile.Move(xpos, ypos)
        Else
            Call DrawPile.Move(xpos + 4, ypos + 2)
        End If
    End If
    
    Call shpCard(0).Move(xpos, ypos, DrawPile.Width, DrawPile.Height)
    Call shpCircle.Move(xpos + (shpCard(0).Width - shpCircle.Width) / 2, _
        ypos + (shpCard(0).Height - shpCircle.Height) / 2)
    cW = DiscardPile.Width
    cH = DiscardPile.Height
    
    xpos = (sw - cW) / 2 + cW * 0.55 - 0.153 * DiscardPile.Width
    ypos = (sH - cH) / 2
    
    If Deal Then
        Call DiscardPile.Move(DrawPile.Left, DrawPile.Top)
    Else
        Call DiscardPile.Move(xpos, ypos)
    End If
    Call shpCard(1).Move(xpos, ypos, DrawPile.Width, DrawPile.Height)
    Call PicColor.Move(shpCard(1).Left + shpCard(1).Width, shpCard(1).Top, _
        shpCard(1).Width * 0.3056, shpCard(1).Height)
End Sub

Private Sub AlignNames()
    Dim cW As Integer, cH As Integer
    
    cW = DrawPile.Width: cH = DrawPile.Height
    
    If Opponents < 3 Then
        Call lblName(0).Move((picTable.ScaleWidth - lblName(0).Width) / 2, _
            picTable.ScaleHeight - cH - cH * 0.05 - lblName(0).Height)
        Call lblName(1).Move((picTable.ScaleWidth - lblName(1).Width) / 2, _
            cH + cH * 0.05)
        Call lblName(2).Move(0, picTable.ScaleHeight / 2 - cH / 2 - lblName(0).Height)
    Else
        Call lblName(0).Move((picTable.ScaleWidth - lblName(0).Width) / 2, _
            picTable.ScaleHeight - cH - cH * 0.05 - lblName(0).Height)
        Call lblName(1).Move(picTable.ScaleWidth - lblName(1).Width, _
            picTable.ScaleHeight / 2 - cH / 2 - lblName(1).Height)
        Call lblName(2).Move((picTable.ScaleWidth - lblName(2).Width) / 2, _
            cH + cH * 0.05)
        Call lblName(3).Move(0, picTable.ScaleHeight / 2 - cH / 2 - lblName(0).Height)
    End If
End Sub

Private Sub CheckWordCard(Player As Object)
    Dim s As Object, i As Integer, n As Integer
    
    If DiscardPile.CardType = [Draw Two] Then
        n = 2
    ElseIf DiscardPile.CardType = [Wild Draw Four] Then
        n = 4
    End If
    
    If DiscardPile.CardType = [Draw Two] Or DiscardPile.CardType = [Wild Draw Four] Then
        Set s = GetNextPlayer(Player, Rotation)
        n = IIf(Uno.DrawPileCards.Count >= n, n, _
            Uno.DrawPileCards.Count)
        For i = 1 To n
            Call Uno.Pick(s, 1)
            Call BeginPlaySound(SND_PICK, False)
            Call PickThrowAnimated(s)
            Call AlignCards(s)
            If AutoSort Then
                If s(0).Name = "PlayerOne" Then Call Uno.SortCards(PlayerOne)
            End If
            Call UpdateDrawPile
            DoEvents
        Next i
        Set NextPlayer = GetNextPlayer(Player, Rotation)
        Set NextPlayer = GetNextPlayer(NextPlayer, Rotation)
    ElseIf (DiscardPile.CardType = Reverse) Then
        If Opponents = 1 Then
        Else
            Rotation = IIf(Rotation = 0, 1, 0)
            Set NextPlayer = GetNextPlayer(Player, Rotation)
        End If
    ElseIf (DiscardPile.CardType = Skip) Then
        If Opponents = 1 Then
        Else
            Set NextPlayer = GetNextPlayer(Player, Rotation)
            Set NextPlayer = GetNextPlayer(NextPlayer, Rotation)
        End If
    Else
        Set NextPlayer = GetNextPlayer(Player, Rotation)
    End If
End Sub

Private Sub Computer()
    Dim ret As Integer, Player As Object
    
    Set Player = NextPlayer
    ret = Uno.AI.ComputerMove(Player, DiscardPile.CardType, _
        DiscardPile.CardColor)
        
    If ret <> -1 Then
        Call PlayerMove(Player, ret)
    Else
        DrawPile_Click
    End If
End Sub

Private Sub DrawPile_Click()
    If Uno.DrawPileCards.Count > 0 Then
        picTable.Enabled = False
        Call Uno.Pick(NextPlayer, 1)
        Call UpdateDrawPile
        Call BeginPlaySound(SND_PICK, False)
        Call PickThrowAnimated(NextPlayer)
        Call AlignCards(NextPlayer)
        If AutoSort Then
            If NextPlayer(0).Name = "PlayerOne" Then Call Uno.SortCards(PlayerOne)
        End If
        Call History("[" & PlayerName(GetPlayer(NextPlayer(0).Name)) _
            & "] Draws a card", -1, 0, NextPlayer(0).Name)
        Set NextPlayer = GetNextPlayer(NextPlayer, Rotation)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call Form_Unload(True)
End Sub

Private Sub Form_Load()
    Call lstHistory.ColumnHeaders.Add(, , "History", _
        lstHistory.Width - Screen.TwipsPerPixelX * 5)
    Call lstPlayers.ColumnHeaders.Add(, , "Player", _
        lstHistory.Width * 0.4)
    Call lstPlayers.ColumnHeaders.Add(, , "Wins", _
        lstHistory.Width * 0.25, lvwColumnCenter)
    Call lstPlayers.ColumnHeaders.Add(, , "Losses", _
        lstHistory.Width * 0.32, lvwColumnCenter)
    Set Me.Icon = LoadResPicture(101, vbResIcon)
    Call OpenSettings
    Call InitLogo
    
    App.HelpFile = App.Path & "\UNO.HLP"
End Sub

Private Sub Form_Resize()
    Call SizeControls
    If picTable.Picture Then Call RepaintLogo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Animated.Play Then Animated.Play = False
    Call SaveSettings
    Call EndPlaySound
    Call WinHelp(0, App.HelpFile, HELP_QUIT, 0)
    End
End Sub

Private Sub SetupGame()
    Dim Card As Control
        
    PicColor.Visible = False
    DiscardPile.ZOrder 0
    DiscardPile.CardHide = True
    If Animated.Speed <> 0 Then
        Call Animated.Linear(DiscardPile, shpCard(1).Left, _
            shpCard(1).Top, DrawPile.Left, DrawPile.Top, _
            (-1 * Animated.Speed + 1))
    End If
    Call AlignControls
    Call ShowControls(True)

    If Animated.Speed <> 0 Then
        DrawPile.CardHide = False
        For Each Card In Me.Controls
            If TypeName(Card) = "UnoCard" Then
                If Card.Visible And Card.Tag = "AnimCard" Then
                    Card.CardHide = False
                    Call Animated.Linear(Card, Card.Left, _
                        Card.Top, DrawPile.Left, DrawPile.Top, _
                        IIf(Animated.Speed < 5, 5, Animated.Speed))
                    Call Card.Move(DrawPile.Left, DrawPile.Top)
                    DrawPile.Repaint = False
                    DrawPile.CardType = Card.CardType
                    DrawPile.CardColor = Card.CardColor
                    Set DrawPile.Picture = Card.Picture
                    DrawPile.Repaint = True
                    Card.Visible = False
                    Card.CardHide = True
                End If
            End If
        Next Card
        DrawPile.CardHide = True
    End If

    If PlayerOne.Count > 1 Then
        Call DestroyObject(PlayerOne, PlayerOne.Count)
    End If
    If PlayerTwo.Count > 1 Then
        Call DestroyObject(PlayerTwo, PlayerTwo.Count)
    End If
    If PlayerThree.Count > 1 Then
        Call DestroyObject(PlayerThree, PlayerThree.Count)
    End If
    If PlayerFour.Count > 1 Then
        Call DestroyObject(PlayerFour, PlayerFour.Count)
    End If
    
    Dim i As Integer, Player(MAX_PLAYERS) As Object
    
    DoEvents
    
    Set Player(0) = PlayerOne
    Set Player(1) = PlayerTwo
    Set Player(2) = PlayerThree
    Set Player(3) = PlayerFour
        
    PicColor.Visible = False
    
    DrawPile.CardHide = True
    Call Uno.Initialize
    For i = 0 To Opponents
        Call Uno.Pick(Player(i), CARDS_PER_PLAYER)
        Call AlignCards(Player(i), True)
        DoEvents
    Next i
    If AutoSort Then
        Call Uno.SortCards(PlayerOne)
    End If
    For Each Card In PlayerOne
        Card.CardHide = False
        If Animated.Speed <> 0 Then DoEvents
    Next Card
    
    Call DiscardPile.ZOrder(0)
    DiscardPile.CardType = Int(10 * Rnd)
    DiscardPile.CardColor = Int(TOTAL_COLORS * Rnd)
    If Animated.Speed <> 0 Then
        Call BeginPlaySound(SND_PICK, False)
        Call Animated.Linear(DiscardPile, DrawPile.Left, DrawPile.Top, _
            shpCard(1).Left, shpCard(1).Top)
        Call EndPlaySound
    End If
        
    Call DiscardPile.Move(shpCard(1).Left, shpCard(1).Top)
    DiscardPile.CardHide = False
    Call Gradient(PicColor, DiscardPile.CardColor)
    PicColor.Visible = True
    SplitCards = Uno.DrawPileCards.Count / 3
    lblRound.Caption = Val(lblRound.Caption) + 1
    Dim n As Integer
    n = DiscardPile.CardType Mod (Opponents + 1)

    Select Case n
    Case Is = 0
        Set NextPlayer = PlayerOne
    Case Is = 1
        Set NextPlayer = PlayerTwo
    Case Is = 2
        Set NextPlayer = PlayerThree
    Case Is = 3
        Set NextPlayer = PlayerFour
    End Select
    lblName(n).ForeColor = vbYellow
    If AutoSort Then
        Call Uno.SortCards(PlayerOne)
    End If
    Call MsgBox(PlayerName(n) & " starts!", vbOKOnly)
End Sub

Private Sub GameInit()
    If Animated.Play Then Animated.Play = False
    
    Deal = True
    StopGame = False
    mnuGameDeal.Enabled = False
    Call LockForm(True)
    lblName(0).ForeColor = vbWhite
    lblName(1).ForeColor = vbWhite
    lblName(2).ForeColor = vbWhite
    lblName(3).ForeColor = vbWhite
    Call ShowPlayers(True)
    lstHistory.ListItems.Clear
    Set picTable.Picture = Nothing
    Call SetupGame
    Call LockForm(False)
    mnuGameDeal.Enabled = True
    Done = False
    Deal = False
End Sub

Private Sub GameStart()
    Do While (Uno.DrawPileCards.Count > 0) And _
        (Not StopGame) And (Not Reset)
        
        If Not mnuGameDemo.Checked Then
            If NextPlayer(0).Name = "PlayerOne" Then
                If Not picTable.Enabled Then
                    picTable.Enabled = True
                    Set DrawPile.MouseIcon = LoadResPicture(101, vbResCursor)
                End If
            Else
                picTable.Enabled = False
                Call Computer
            End If
        Else
            picTable.Enabled = False
            Call Computer
        End If
        DoEvents
    Loop
    
    If Not Reset Then Call GameEnd
End Sub

Private Sub GameEnd()
    If Not Done Then
        If Uno.DiscardPileCards.Count = 0 Then _
            DrawPile.Visible = False
        PlayerScore(0) = Uno.TotalPoints(PlayerOne)
        PlayerScore(1) = Uno.TotalPoints(PlayerTwo)
        PlayerScore(2) = Uno.TotalPoints(PlayerThree)
        PlayerScore(3) = Uno.TotalPoints(PlayerFour)
        Dim i As Integer, n As Integer, Winner As Integer
        
        Winner = -1
        If Opponents = 1 Then
            If PlayerOne.Count < 2 Then
                Winner = 0
            ElseIf PlayerTwo.Count < 2 Then
                Winner = 1
            End If
        ElseIf Opponents = 2 Then
            If PlayerOne.Count < 2 Then
                Winner = 0
            ElseIf PlayerTwo.Count < 2 Then
                Winner = 1
            ElseIf PlayerThree.Count < 2 Then
                Winner = 2
            End If
        ElseIf Opponents = 3 Then
            If PlayerOne.Count < 2 Then
                Winner = 0
            ElseIf PlayerTwo.Count < 2 Then
                Winner = 1
            ElseIf PlayerThree.Count < 2 Then
                Winner = 2
            ElseIf PlayerFour.Count < 2 Then
                Winner = 3
            End If
        End If
        
        If Winner = -1 Then Winner = GetWinner
        frmScore.Winner = Winner
        Call frmScore.Show(vbModal, Me)
        
        For i = 0 To Opponents
            If Winner = i Then
                n = Val(lstPlayers.ListItems(i + 1).SubItems(1))
                    lstPlayers.ListItems(i + 1).SubItems(1) = n + 1
                lstPlayers.ListItems(i + 1).ListSubItems(1).ForeColor = vbRed
                lstPlayers.ListItems(i + 1).ListSubItems(2).ForeColor = vbRed
            Else
                n = Val(lstPlayers.ListItems(i + 1).SubItems(2))
                lstPlayers.ListItems(i + 1).SubItems(2) = n + 1
            End If
        Next i
    
        If Winner = 0 And Not mnuGameDemo.Checked Then
            Call ShowPlayers(False)
            n = TotalCards(Me)
            If n < 15 Then
                n = Abs(15 - n)
                Call CreateObject(PlayerOne, n)
            End If
            Dim Card As Object
            For Each Card In PlayerOne
                Card.CardType = Int(15 * Rnd)
                Card.CardColor = Int(4 * Rnd)
            Next Card
            Call Victory
        End If
        Done = True
    End If
End Sub

Private Function GetNextPlayer(Player As Object, Rotation As Integer) As Object
    Dim temp As Object
    
    If Opponents = 1 Then
        Select Case Player(0).Name
        Case Is = "PlayerOne"
            Set temp = PlayerTwo
        Case Is = "PlayerTwo"
            Set temp = PlayerOne
        End Select
    ElseIf Opponents = 2 Then
        Select Case Player(0).Name
        Case Is = "PlayerOne"
            Set temp = IIf(Rotation = 0, PlayerTwo, PlayerThree)
        Case Is = "PlayerTwo"
            Set temp = IIf(Rotation = 0, PlayerThree, PlayerOne)
        Case Is = "PlayerThree"
            Set temp = IIf(Rotation = 0, PlayerOne, PlayerTwo)
        End Select
    ElseIf Opponents = 3 Then
        Select Case Player(0).Name
        Case Is = "PlayerOne"
            Set temp = IIf(Rotation = 0, PlayerTwo, PlayerFour)
        Case Is = "PlayerTwo"
            Set temp = IIf(Rotation = 0, PlayerThree, PlayerOne)
        Case Is = "PlayerThree"
            Set temp = IIf(Rotation = 0, PlayerFour, PlayerTwo)
        Case Is = "PlayerFour"
            Set temp = IIf(Rotation = 0, PlayerOne, PlayerThree)
        End Select
    End If
    
    lblName(0).ForeColor = vbWhite
    lblName(1).ForeColor = vbWhite
    lblName(2).ForeColor = vbWhite
    lblName(3).ForeColor = vbWhite
    
    If temp(0).Name = "PlayerOne" Then
        lblName(0).ForeColor = vbYellow
    ElseIf temp(0).Name = "PlayerTwo" Then
        lblName(1).ForeColor = vbYellow
    ElseIf temp(0).Name = "PlayerThree" Then
        lblName(2).ForeColor = vbYellow
    Else
        lblName(3).ForeColor = vbYellow
    End If
    Set GetNextPlayer = temp
End Function

Private Sub History(prompt As String, id As Integer, cc As Integer, Player As String)
    Dim s As String
    
    If id <> -1 Then
        If id <> [Wild Draw Four] Then
            Select Case cc
            Case Is = 0: s = "Blue"
            Case Is = 1: s = "Red"
            Case Is = 2: s = "Green"
            Case Is = 3: s = "Yellow"
            End Select
        
            Select Case id
            Case 0 To 9: s = s & " " & id
            Case Is = 10: s = s & " " & "Draw 2"
            Case Is = 11: s = s & " " & "Reverse"
            Case Is = 12: s = s & " " & "Skip"
            Case Is = 13: s = s & " " & "Wild Card"
            End Select
        Else
            s = "Wild Draw 4"
        End If
    End If
    
    s = prompt & " " & s & "."
    
    Dim itm As ListItem
    Set itm = lstHistory.ListItems.Add(, , s)
    itm.ToolTipText = s
    If Player = "PlayerOne" Then
        itm.ForeColor = vbRed
    Else
        itm.ForeColor = vbBlack
    End If
    Call lstHistory.SetFocus
    Call SendKeys("{End}", True)
    Call SendKeys("{End}", True)
End Sub

Private Sub InitLogo()
    If Dir$(PathLogo) <> "" Then
        Dim BM As BITMAP
        If GetObject(LoadPicture(PathLogo), Len(BM), BM) Then
            picLogo.Width = BM.bmWidth
            picLogo.Height = BM.bmHeight
            picLogo.BackColor = picTable.BackColor
            Call Logo(picLogo, picTemp)
            Set picTemp.Picture = Nothing
            Call RepaintLogo
        End If
    End If
End Sub

Private Sub LockForm(bVal As Boolean)
    bVal = Not bVal
    Me.Enabled = bVal
    mnuGameDeal.Enabled = bVal
    mnuGameOptions.Enabled = bVal
End Sub

Private Sub mnuGameCheat_Click()
    mnuGameCheat.Checked = Not mnuGameCheat.Checked
    Call OpenCards(PlayerTwo, mnuGameCheat.Checked)
    Call OpenCards(PlayerThree, mnuGameCheat.Checked)
    Call OpenCards(PlayerFour, mnuGameCheat.Checked)
End Sub

Private Sub mnuGameDeal_Click()
    Call GameInit
    Call GameStart
End Sub

Private Sub mnuGameDemo_Click()
    mnuGameDemo.Checked = Not mnuGameDemo.Checked
End Sub

Private Sub mnuGameExit_Click()
    Call Unload(Me)
End Sub

Private Sub mnuGameNew_Click()
    Dim i As Integer
    
    mnuGameNew.Enabled = False
    mnuGameDeal.Enabled = True
    mnuGameReset.Enabled = True
    Call lstPlayers.ListItems.Clear
    For i = 0 To Opponents
        Call lstPlayers.ListItems.Add(i + 1, , PlayerName(i))
        lstPlayers.ListItems(i + 1).SubItems(1) = 0
        lstPlayers.ListItems(i + 1).SubItems(2) = 0
        If i = 0 Then
            lstPlayers.ListItems(i + 1).ForeColor = vbRed
            lstPlayers.ListItems(i + 1).ListSubItems(1).ForeColor = vbRed
            lstPlayers.ListItems(i + 1).ListSubItems(2).ForeColor = vbRed
        End If
    Next i
    lblRound.Caption = 0
    Call ShowPlayers(True)
    Reset = False
    NewGame = False
    sbStatusBar.Panels(2).text = "Difficulty : " & UCase$(Difficulty)
    Call mnuGameDeal_Click
End Sub

Private Sub mnuGameOptions_Click()
    Call frmOptions.Show(vbModal, Me)
End Sub

Private Sub mnuGameReset_Click()
    Dim s As String
    
    s = "Are you sure you want to start a new game?"
    If MsgBox(s, vbYesNo Or vbQuestion) = vbYes Then
        mnuGameNew.Enabled = True
        mnuGameReset.Enabled = False
        mnuGameDeal.Enabled = False
        picTable.Enabled = False
        Reset = True
        NewGame = True
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    Call frmAbout.Show(vbModal, Me)
End Sub

Private Sub RepaintLogo()
    If picLogo.Picture Then
        Set picTable.Picture = Nothing
        picTable.AutoRedraw = True
        picTable.PaintPicture picLogo.Picture, _
            (picTable.ScaleWidth - picLogo.ScaleWidth) / 2, _
            (picTable.ScaleHeight - picLogo.ScaleHeight) / 2
        Set picTable.Picture = picTable.Image
        picTable.AutoRedraw = False
    End If
End Sub

Private Sub SizeControls()
    On Error Resume Next
    
    Call fraTray.Move(Me.ScaleWidth - fraTray.Width - _
        Screen.TwipsPerPixelX, fraTray.Top, _
        fraTray.Width, Me.ScaleHeight - sbStatusBar.Height)
    picTable.Width = Me.ScaleWidth - fraTray.Width
    picTable.Height = Me.ScaleHeight - sbStatusBar.Height
    Call lstPlayers.Move(lstPlayers.Left, _
        fraTray.Height - lstPlayers.Height)
    Call lblPrompt.Move(0, lstPlayers.Top - _
        lblPrompt.Height)
    Call lblRound.Move(lblPrompt.Width, lblPrompt.Top)
    Call lstHistory.Move(lstHistory.Left, lstHistory.Top, _
        lstHistory.Width, lblPrompt.Top - lblPrompt.Height)
    Call AlignControls
    If Not Deal Then DoEvents
    
    If Not Animated.Play Then
        Call AlignCards(PlayerOne)
        If Not Deal Then DoEvents
        Call AlignCards(PlayerTwo)
        If Not Deal Then DoEvents
        Call AlignCards(PlayerThree)
        If Not Deal Then DoEvents
        Call AlignCards(PlayerFour)
        If Not Deal Then DoEvents
        Call AlignNames
        If Not Deal Then DoEvents
    End If
    
    If Animated.Play Then
        If SelectedAnimation = 0 Then
            Animated.BounceRightEdge = picTable.ScaleWidth
            Animated.BounceBottomEdge = picTable.ScaleHeight
        ElseIf SelectedAnimation = 1 Then
            Animated.RotXradius = (picTable.ScaleWidth - PlayerOne(0).Width) / 2
            Animated.RotYradius = (picTable.ScaleHeight - PlayerOne(0).Height) / 2
        Else
            Animated.WaveRightEdge = picTable.ScaleWidth
            Animated.WaveBottomEdge = picTable.ScaleHeight
        End If
    End If
End Sub

Private Sub ShowControls(bVal As Boolean)
    DiscardPile.Visible = bVal
    DrawPile.Visible = bVal
    Dummy(0).Visible = bVal
    Dummy(1).Visible = bVal
    shpCard(0).Visible = bVal
    shpCard(1).Visible = bVal
    shpCircle.Visible = bVal
End Sub

Private Sub ShowPlayers(bVal As Boolean)
    Dim i As Integer
    
    For i = 0 To lblName.Count - 1
        lblName(i).Visible = False
    Next i
    
    If Not bVal Then Exit Sub

    For i = 0 To Opponents
        lblName(i).Caption = PlayerName(i)
    Next i
    Call AlignNames
    For i = 0 To Opponents
        lblName(i).Visible = bVal
    Next i
End Sub

Private Sub mnuHelpContents_Click()
    Call WinHelp(0, App.HelpFile, HELP_CONTEXT, 100)
End Sub

Private Sub PlayerOne_Click(Index As Integer)
    picTable.Enabled = False
    Call PlayerMove(PlayerOne, Index)
End Sub

Private Sub PlayerMove(Player As Object, Index As Integer)
    Dim id As Integer, cc As Integer
        
    id = DiscardPile.CardType
    cc = DiscardPile.CardColor
    
    If Uno.ChkMove(Player, Index, id, cc) Then
        picTable.Enabled = False
        Call BeginPlaySound(SND_THROW, False)
        Dim X As Integer, Y As Integer
        ' save player one position
        X = Player(Index).Left: Y = Player(Index).Top
        If Player(0).Name = "PlayerOne" Then Me.Enabled = False
        Call Animated.Linear(Player(Index), Player(Index).Left, _
            Player(Index).Top, DiscardPile.Left, DiscardPile.Top)
        If Player(0).Name = "PlayerOne" Then Me.Enabled = True
        If Not mnuGameDemo.Checked Then
            If Player(0).Name = "PlayerOne" Then
                If (Player(Index).CardType = Wild) Or (Player(Index).CardType = [Wild Draw Four]) Then
                    Call frmColor.Show(vbModal)
                    If frmColor.Color = -1 Then
                        Call Animated.Linear(Player(Index), Player(Index).Left, _
                            Player(Index).Top, X, Y)
                        Call AlignCards(Player)
                        Exit Sub
                    Else
                        cc = frmColor.Color
                    End If
                Else
                    cc = Player(Index).CardColor
                End If
            Else
                cc = Player(Index).CardColor
            End If
        Else
            cc = Player(Index).CardColor
        End If
        
        DiscardPile.CardType = Player(Index).CardType
        DiscardPile.CardColor = cc
        Call DiscardPile.ZOrder(0)
        Call Gradient(PicColor, DiscardPile.CardColor)
        Call CheckWordCard(Player)
        Call History("[" & PlayerName(GetPlayer(Player(0).Name)) & "] plays", _
            DiscardPile.CardType, DiscardPile.CardColor, Player(0).Name)
        Call Uno.Throw(Player, Index)
        Call AlignCards(Player)
        StopGame = IIf(Player.Count - 1 = 0, True, False)
    Else
        Call Beep
    End If
    Exit Sub
    
ErrHandler:
End Sub

Private Sub PickThrowAnimated(Player As Object)
    Dim xmid As Single, ymid As Single
    
    xmid = (picTable.ScaleWidth - Player(0).Width) / 2
    If Opponents < 2 Then
        If Player(0).Name = "PlayerThree" Then
            Call Animated.Linear(Player(Player.Count - 1), DrawPile.Left, _
                DrawPile.Top, Player(Player.Count - 2).Left + Player(0).Width * 0.1, Player(0).Top)
        Else
            Call Animated.Linear(Player(Player.Count - 1), DrawPile.Left, _
                DrawPile.Top, xmid, Player(0).Top)
        End If
    Else
        If Player(0).Name = "PlayerFour" Then
            Call Animated.Linear(Player(Player.Count - 1), DrawPile.Left, DrawPile.Top, _
                Player(Player.Count - 2).Left - Player(0).Width * 0.1, Player(0).Top)
        ElseIf Player(0).Name = "PlayerTwo" Then
            Call Animated.Linear(Player(Player.Count - 1), DrawPile.Left, _
                DrawPile.Top, Player(Player.Count - 2).Left + Player(0).Width * 0.1, Player(0).Top)
        Else
            Call Animated.Linear(Player(Player.Count - 1), DrawPile.Left, _
                DrawPile.Top, xmid, Player(0).Top)
        End If
    End If
End Sub

Private Sub UpdateDrawPile()
    If (SplitCards * 2) = Uno.DrawPileCards.Count Then
        Call DrawPile.Move(Dummy(1).Left, Dummy(1).Top)
        Call DrawPile.ZOrder(0)
        Dummy(1).Visible = False
    ElseIf SplitCards = Uno.DrawPileCards.Count Then
        Call DrawPile.Move(Dummy(0).Left, Dummy(0).Top)
        Call DrawPile.ZOrder(0)
        Dummy(0).Visible = False
    ElseIf Uno.DrawPileCards.Count < 1 Then
        DrawPile.Visible = False
    End If
    Call DrawPile.ZOrder(0)
End Sub

Private Sub Victory()
    PicColor.Visible = False
    Call ShowControls(False)
    If AnimationType = 3 Then
        SelectedAnimation = Second(Time) Mod 3
    Else
        SelectedAnimation = AnimationType
    End If
    
    Call RepaintLogo
    Animated.Play = True
    If SelectedAnimation = 0 Then
        Animated.BounceRightEdge = picTable.ScaleWidth
        Animated.BounceBottomEdge = picTable.ScaleHeight
        Call Animated.Bounce(Me)
    ElseIf SelectedAnimation = 1 Then
        Animated.RotXradius = (picTable.ScaleWidth - PlayerOne(0).Width) / 2
        Animated.RotYradius = (picTable.ScaleHeight - PlayerOne(0).Height) / 2
        Call Animated.Rotation(Me)
    Else
        Animated.WaveRightEdge = picTable.ScaleWidth
        Animated.WaveBottomEdge = picTable.ScaleHeight
        Call Animated.Wave(Me, WaveExpression, _
            DrawPile.Width, DrawPile.Height)
    End If
End Sub

Private Sub PlayerOne_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static OldIndex As Integer
    
    If Index = OldIndex Then Exit Sub
    If Uno.ChkMove(PlayerOne, Index, DiscardPile.CardType, DiscardPile.CardColor) Then
        Set PlayerOne(Index).MouseIcon = LoadResPicture(101, vbResCursor)
        PlayerOne(Index).MousePointer = vbCustom
    Else
        PlayerOne(Index).MousePointer = vbArrow
    End If
    OldIndex = Index
End Sub


