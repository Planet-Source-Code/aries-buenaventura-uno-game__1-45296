VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5160
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   315
      Left            =   60
      TabIndex        =   40
      Top             =   3420
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   3420
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3060
      TabIndex        =   2
      Top             =   3420
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   3420
      Width           =   975
   End
   Begin VB.Frame fraGeneral 
      Height          =   2835
      Left            =   120
      TabIndex        =   4
      Top             =   420
      Width           =   4935
      Begin VB.Frame fraSort 
         Caption         =   "Sort"
         Height          =   795
         Left            =   3120
         TabIndex        =   42
         Top             =   1860
         Width           =   1695
         Begin VB.CheckBox chkAutoSort 
            Caption         =   "Auto Sort"
            Height          =   255
            Left            =   360
            TabIndex        =   43
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame fraSpeed 
         Caption         =   "Animation speed"
         Height          =   795
         Left            =   120
         TabIndex        =   25
         Top             =   1860
         Width           =   2955
         Begin MSComctlLib.Slider sldSpeed 
            Height          =   315
            Left            =   60
            TabIndex        =   29
            Top             =   180
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
         End
         Begin VB.Label lblAnimation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "10"
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   28
            Top             =   540
            Width           =   180
         End
         Begin VB.Label lblAnimation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            Height          =   195
            Index           =   1
            Left            =   1335
            TabIndex        =   27
            Top             =   540
            Width           =   210
         End
         Begin VB.Label lblAnimation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   540
            Width           =   90
         End
      End
      Begin VB.Frame fraName 
         Caption         =   "Player names"
         Height          =   1695
         Left            =   2340
         TabIndex        =   12
         Top             =   180
         Width           =   2475
         Begin VB.TextBox txtName 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   20
            Text            =   "Plue"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtName 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   19
            Text            =   "Musica"
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Index           =   1
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   18
            Text            =   "Elie"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Index           =   0
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   13
            Text            =   "Haru"
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Computer III"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   17
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Computer II"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   16
            Top             =   960
            Width           =   810
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Computer I"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   15
            Top             =   660
            Width           =   765
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "You"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   14
            Top             =   360
            Width           =   285
         End
      End
      Begin VB.Frame fraOpponents 
         Caption         =   "Opponents"
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   2115
         Begin VB.OptionButton optOpponent 
            Caption         =   "One"
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   8
            Top             =   360
            Width           =   1515
         End
         Begin VB.OptionButton optOpponent 
            Caption         =   "Two"
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   7
            Top             =   660
            Width           =   1515
         End
         Begin VB.OptionButton optOpponent 
            Caption         =   "Three"
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   6
            Top             =   1020
            Width           =   1515
         End
      End
   End
   Begin VB.Frame fraDifficulty 
      Caption         =   "Select Difficulty"
      Height          =   2835
      Left            =   120
      TabIndex        =   21
      Top             =   420
      Visible         =   0   'False
      Width           =   4935
      Begin VB.OptionButton optLevel 
         Caption         =   "Difficult"
         Height          =   435
         Index           =   2
         Left            =   420
         TabIndex        =   24
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Easy"
         Height          =   435
         Index           =   0
         Left            =   420
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Normal"
         Height          =   435
         Index           =   1
         Left            =   420
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame fraDeck 
      Caption         =   "Select Card Back"
      Height          =   2835
      Left            =   120
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox picTray 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         DrawWidth       =   4
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   60
         ScaleHeight     =   615
         ScaleWidth      =   675
         TabIndex        =   10
         Top             =   180
         Width           =   675
         Begin Uno.UnoCard crdDeck 
            Height          =   1350
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   2381
            CardHide        =   -1  'True
            Picture         =   "frmOptions.frx":0000
         End
      End
   End
   Begin VB.Frame fraAnimation 
      Height          =   2835
      Left            =   120
      TabIndex        =   30
      Top             =   420
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdAdvance 
         Caption         =   "Advance"
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   41
         Top             =   2160
         Width           =   3915
      End
      Begin VB.TextBox txtExpression 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton cmdCustomWave 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   38
         Top             =   1800
         Width           =   315
      End
      Begin VB.OptionButton optAnimation 
         Caption         =   "Wave"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   36
         Top             =   1380
         Width           =   1275
      End
      Begin VB.CheckBox chkAnimation 
         Caption         =   "Random animation"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   35
         Top             =   2580
         Value           =   2  'Grayed
         Width           =   1695
      End
      Begin VB.CheckBox chkAnimation 
         Caption         =   "Custom animation"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   34
         Top             =   660
         Width           =   1695
      End
      Begin VB.OptionButton optAnimation 
         Caption         =   "Rotation"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   33
         Top             =   1140
         Width           =   1275
      End
      Begin VB.OptionButton optAnimation 
         Caption         =   "Bounce"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   32
         Top             =   900
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.Label lblExpression 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expression:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1020
         TabIndex        =   37
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label lblPrompt 
         Caption         =   "Select the type of animation would you like to play when user wins the game."
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   180
         Width           =   4335
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   3315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5847
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Deck"
            Key             =   "Deck"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Difficulty"
            Key             =   "Difficulty"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Animation"
            Key             =   "Animation"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AnimationChanged As Boolean
Dim AutoSortChanged As Boolean
Dim DeckChanged As Boolean
Dim LevelChanged As Boolean
Dim NameChanged As Boolean
Dim SpeedChanged As Boolean

Dim Level As Integer
Dim nAnimation As Integer

Private Sub chkAnimation_Click(Index As Integer)
    Select Case Index
    Case Is = 0
        If chkAnimation(0).Value = vbChecked Then
            chkAnimation(1).Enabled = False
            chkAnimation(1).Value = vbUnchecked
            optAnimation(0).Enabled = True
            optAnimation(1).Enabled = True
            optAnimation(2).Enabled = True
            cmdAdvance.Enabled = True
        ElseIf chkAnimation(0).Value = vbUnchecked Then
            chkAnimation(1).Enabled = True
            chkAnimation(1).Value = vbChecked
            optAnimation(0).Enabled = False
            optAnimation(1).Enabled = False
            optAnimation(2).Enabled = False
            cmdAdvance.Enabled = False
        End If
    Case Is = 1
        If chkAnimation(1).Value = vbChecked Then
            chkAnimation(0).Enabled = False
            chkAnimation(0).Value = vbUnchecked
            nAnimation = 3
        ElseIf chkAnimation(1).Value = vbUnchecked Then
            chkAnimation(0).Enabled = True
            chkAnimation(0).Value = vbChecked
        End If
    End Select
    
    If optAnimation(2).Value Then
        If chkAnimation(0).Value = vbChecked Then
            cmdCustomWave.Enabled = True
            lblExpression.Enabled = True
            txtExpression.Enabled = True
        ElseIf chkAnimation(0).Value = vbUnchecked Then
            cmdCustomWave.Enabled = False
            lblExpression.Enabled = False
            txtExpression.Enabled = False
        End If
    End If
    AnimationChanged = True
    cmdApply.Enabled = True
End Sub

Private Sub chkAutoSort_Click()
    AutoSortChanged = True
    cmdApply.Enabled = True
End Sub

Private Sub cmdAdvance_Click()
    Call frmCustom.Show(vbModal, Me)
    If frmCustom.retExpression <> "" Then
        txtExpression.text = frmCustom.retExpression
    End If
End Sub

Private Sub cmdApply_Click()
    If AnimationChanged Then
        AnimationType = nAnimation
        AnimationChanged = False
    End If
    
    If AutoSortChanged Then
        If chkAutoSort.Value = vbChecked Then
            Dim temp As New clsUno
            Call temp.SortCards(frmMain.PlayerOne)
        End If
        AutoSort = chkAutoSort
        AutoSortChanged = False
    End If
    
    If DeckChanged Then
        Call UpdateDeck(CardDeck)
        DeckChanged = False
    End If
    
    If LevelChanged Then
        Select Case Level
        Case Is = 0
            Difficulty = "EASY"
        Case Is = 1
            Difficulty = "NORMAL"
        Case Is = 2
            Difficulty = "DIFFICULT"
        End Select
        LevelChanged = False
    End If

    If NameChanged Then
        PlayerName(0) = Trim$(txtName(0).text)
        PlayerName(1) = Trim$(txtName(1).text)
        PlayerName(2) = Trim$(txtName(2).text)
        PlayerName(3) = Trim$(txtName(3).text)
        NameChanged = False
    End If
    
    If SpeedChanged Then
        Animated.Speed = sldSpeed.Value
        SpeedChanged = False
    End If
      
    cmdApply.Enabled = False
End Sub

Private Sub cmdPlay_Click()
    Call Unload(Me)
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdCustomWave_Click()
    Call frmList.Show(vbModal, Me)
    If frmList.retExpression <> "" Then
        txtExpression.text = frmList.retExpression
    End If
End Sub

Private Sub cmdDefault_Click()
    Call DefaultSettings
    Call DestroyObject(crdDeck, crdDeck.Count)
    Call Form_Load
    cmdApply.Enabled = True
End Sub

Private Sub cmdOk_Click()
    Call cmdApply_Click
    Call Unload(Me)
End Sub
    
Private Sub crdDeck_Click(Index As Integer)
    Dim BorderWidth As Integer, Color As Long
    
    picTray.Cls
    picTray.Line (crdDeck(Index).Left, crdDeck(Index).Top)- _
        (crdDeck(Index).Left + crdDeck(Index).Width, _
        crdDeck(Index).Top + crdDeck(Index).Height), vbBlue, B
    CardDeck = Index
    DeckChanged = True
    cmdApply.Enabled = True
End Sub

Private Sub crdDeck_GotFocus(Index As Integer)
    DeckChanged = True
    Call crdDeck_Click(Index)
End Sub

Private Sub Form_Load()
    sldSpeed.Value = Animated.Speed
    
    Call CreateObject(crdDeck, 3)
    Dim x As Integer, step As Long, Card As Object
    Dim CardWidth As Integer, CardHeight As Integer
    CardWidth = crdDeck(0).Width
    CardHeight = crdDeck(0).Height
    Dim twpX As Integer, twpY As Integer
    twpX = Screen.TwipsPerPixelX: twpY = Screen.TwipsPerPixelY
    picTray.Width = CardWidth * crdDeck.Count + CardWidth * 0.33 + _
        picTray.DrawWidth * twpX * 2
    picTray.Height = CardHeight + picTray.DrawWidth * twpY * 2
    Call picTray.Move((fraDeck.Width - picTray.Width) / 2, _
        (fraDeck.Height - picTray.Height) / 2)
    x = CardWidth + (CardWidth * 0.33) / (crdDeck.Count - 1)
    For Each Card In crdDeck
        Card.Deck = Card.Index
        Call Card.Move(step + twpX * picTray.DrawWidth, _
            twpY * picTray.DrawWidth)
        step = step + x
    Next Card
    
    txtName(0).text = PlayerName(0)
    txtName(1).text = PlayerName(1)
    txtName(2).text = PlayerName(2)
    txtName(3).text = PlayerName(3)
    
    Call optOpponent_Click(Opponents - 1)
    Call crdDeck_Click(CardDeck)
    Select Case UCase$(Difficulty)
    Case Is = "EASY"
        optLevel(0).Value = True
    Case Is = "NORMAL"
        optLevel(1).Value = True
    Case Is = "DIFFICULT"
        optLevel(2).Value = True
    End Select
    
    If Not NewGame Then
        optOpponent(0).Enabled = False
        optOpponent(1).Enabled = False
        optOpponent(2).Enabled = False
        
        lblName(0).Enabled = False
        lblName(1).Enabled = False
        lblName(2).Enabled = False
        lblName(3).Enabled = False
        
        txtName(0).Enabled = False
        txtName(1).Enabled = False
        txtName(2).Enabled = False
        txtName(3).Enabled = False
        
        optLevel(0).Enabled = False
        optLevel(1).Enabled = False
        optLevel(2).Enabled = False
        cmdDefault.Enabled = False
    End If
    
    If AnimationType < 3 Then
        chkAnimation(0).Value = vbChecked
        Call chkAnimation_Click(0)
        optAnimation(AnimationType).Value = True
        Call optAnimation_Click(AnimationType)
    Else
        chkAnimation(1).Value = vbChecked
        Call chkAnimation_Click(1)
    End If
    txtExpression.text = WaveExpression
    chkAutoSort.Value = IIf(AutoSort, vbChecked, vbUnchecked)
    AnimationChanged = False
    cmdApply.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyObject(crdDeck, crdDeck.Count)
End Sub

Private Sub optAnimation_Click(Index As Integer)
    Select Case Index
    Case 0 To 1
        cmdCustomWave.Enabled = False
        lblExpression.Enabled = False
        txtExpression.Enabled = False
    Case Is = 2
        cmdCustomWave.Enabled = True
        lblExpression.Enabled = True
        txtExpression.Enabled = True
    End Select
    
    nAnimation = Index
    AnimationChanged = True
    cmdApply.Enabled = True
End Sub

Private Sub optLevel_Click(Index As Integer)
    Level = Index
    LevelChanged = True
    cmdApply.Enabled = True
End Sub

Private Sub optOpponent_Click(Index As Integer)
    Dim i As Integer
    
    optOpponent(Index).Value = True
    
    Opponents = Index + 1
    lblName(2).Enabled = False
    txtName(2).Enabled = False
    lblName(3).Enabled = False
    txtName(3).Enabled = False
    
    If Opponents > 1 Then
        For i = 2 To Opponents
            lblName(i).Enabled = True
            txtName(i).Enabled = True
        Next i
    End If
    cmdApply.Enabled = True
End Sub

Private Sub sldSpeed_Change()
    SpeedChanged = True
    cmdApply.Enabled = True
End Sub

Private Sub tbsOptions_Click()
    Select Case tbsOptions.SelectedItem.Key
    Case Is = "General"
        fraGeneral.Visible = True
        fraDeck.Visible = False
        fraDifficulty.Visible = False
        fraAnimation.Visible = False
    Case Is = "Deck"
        fraGeneral.Visible = False
        fraDeck.Visible = True
        fraDifficulty.Visible = False
        fraAnimation.Visible = False
    Case Is = "Difficulty"
        fraGeneral.Visible = False
        fraDeck.Visible = False
        fraDifficulty.Visible = True
        fraAnimation.Visible = False
    Case Is = "Animation"
        fraGeneral.Visible = False
        fraDeck.Visible = False
        fraDifficulty.Visible = False
        fraAnimation.Visible = True
    End Select
End Sub

Private Sub txtExpression_Change()
    WaveExpression = Trim$(txtExpression.text)
    AnimationChanged = True
    cmdApply.Enabled = True
End Sub

Private Sub txtName_Change(Index As Integer)
    NameChanged = True
    cmdApply.Enabled = True
End Sub

Private Sub txtName_LostFocus(Index As Integer)
    If Trim$(txtName(Index).text) = vbNullString Then
        txtName(Index).text = Trim$(PlayerName(Index))
    End If
End Sub
