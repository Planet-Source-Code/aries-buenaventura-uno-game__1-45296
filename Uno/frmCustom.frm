VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCustom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFrame 
      Height          =   4815
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   4455
         Left            =   120
         MousePointer    =   99  'Custom
         ScaleHeight     =   295
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   567
         TabIndex        =   1
         Top             =   240
         Width           =   8535
         Begin VB.PictureBox picLogo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   960
            ScaleHeight     =   24
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   27
            TabIndex        =   3
            Top             =   60
            Visible         =   0   'False
            Width           =   405
         End
         Begin Uno.UnoCard crdTemp 
            Height          =   1350
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Tag             =   "AnimCard"
            Top             =   0
            Visible         =   0   'False
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   2381
            Picture         =   "frmCustom.frx":0000
         End
      End
      Begin VB.PictureBox picGraph 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4455
         Left            =   120
         ScaleHeight     =   295
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   567
         TabIndex        =   2
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.PictureBox picTray 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   8910
      TabIndex        =   4
      Top             =   4740
      Width           =   8910
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   7860
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Default         =   -1  'True
         Height          =   375
         Left            =   7860
         TabIndex        =   19
         Top             =   180
         Width           =   975
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Custom wave animation"
         Height          =   1455
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   3075
         Begin VB.TextBox txtExpression 
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   420
            Width           =   2535
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
            Height          =   315
            Left            =   2160
            TabIndex        =   17
            Top             =   1020
            Width           =   795
         End
         Begin VB.CommandButton cmdGraph 
            Caption         =   "&Graph"
            Height          =   315
            Left            =   1320
            TabIndex        =   16
            Top             =   1020
            Width           =   795
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "&Play"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   1020
            Width           =   795
         End
         Begin VB.CommandButton cmdExpression 
            Caption         =   "..."
            Height          =   315
            Left            =   2640
            TabIndex        =   9
            Top             =   405
            Width           =   315
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expression:"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   11
            Top             =   240
            Width           =   810
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "note: x is constant"
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   2
            Left            =   1620
            TabIndex        =   10
            Top             =   720
            Width           =   1290
         End
      End
      Begin VB.Frame fraFrame 
         Height          =   1455
         Index           =   3
         Left            =   3180
         TabIndex        =   5
         Top             =   60
         Width           =   2415
         Begin VB.CommandButton cmdRotation 
            Caption         =   "Rotation"
            Height          =   375
            Left            =   1260
            TabIndex        =   14
            Top             =   180
            Width           =   975
         End
         Begin VB.CommandButton cmdBounce 
            Caption         =   "Bounce"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   180
            Width           =   975
         End
         Begin MSComctlLib.Slider sldSpeed 
            Height          =   195
            Left            =   660
            TabIndex        =   6
            Top             =   1140
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   344
            _Version        =   393216
         End
         Begin VB.Label lblPrompt 
            AutoSize        =   -1  'True
            Caption         =   "Speed:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   7
            Top             =   1140
            Width           =   510
         End
      End
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_CARDS = 15

Public retExpression As String

Dim Animated As New clsAnimation

Dim step As Integer
Dim Crest As Single, i As Integer
Dim X As Integer, Y As Integer
Dim w As Integer, h As Integer
Dim xmid As Integer, ymid As Integer

Private Sub cmdBounce_Click()
    Call picPreview.ZOrder(0)
    
    Animated.Play = True
    Animated.BounceRightEdge = picPreview.ScaleWidth
    Animated.BounceBottomEdge = picPreview.ScaleHeight
    Call Animated.Bounce(Me)
End Sub

Private Sub cmdCancel_Click()
    Animated.Play = False
    Me.Hide
End Sub

Private Sub cmdClear_Click()
    Call picGraph.Cls
End Sub

Private Sub cmdExpression_Click()
    Call frmList.Show(vbModal, Me)
    If frmList.retExpression <> "" Then
        txtExpression.text = frmList.retExpression
    End If
End Sub

Private Sub cmdGraph_Click()
    If ChkExpression(txtExpression.text) Then
        If Animated.Play Then Animated.Play = False
        Call PlotGraph
        Call picGraph.ZOrder(0)
    End If
End Sub

Private Sub cmdOk_Click()
    If ChkExpression(txtExpression.text) Then
        retExpression = Trim$(txtExpression.text)
        Animated.Play = False
        Me.Hide
    End If
End Sub

Private Sub cmdPlay_Click()
    If ChkExpression(txtExpression.text) Then
        Call PlotGraph
        Call picPreview.ZOrder(0)
    
        Animated.Play = True
        Animated.WaveRightEdge = picPreview.ScaleWidth
        Animated.WaveBottomEdge = picPreview.ScaleHeight
        Call Animated.Wave(Me, Trim$(txtExpression.text), _
            crdTemp(0).Width, crdTemp(0).Height)
    End If
End Sub

Private Sub cmdRotation_Click()
    Call picPreview.ZOrder(0)
    
    Animated.Play = True
    Animated.RotXradius = (picPreview.ScaleWidth - crdTemp(0).Width) / 2
    Animated.RotYradius = (picPreview.ScaleHeight - crdTemp(0).Height) / 2
    Call Animated.Rotation(Me)
End Sub

Private Sub Form_Load()
    Dim Card As Object
    
    Call CreateObject(crdTemp, NUM_CARDS)
    For Each Card In crdTemp
        Card.CardType = Int(15 * Rnd)
        Card.CardColor = Int(4 * Rnd)
    Next Card
    
    sldSpeed.Value = Animated.Speed
    txtExpression.text = WaveExpression
    
    Call CoordinateSystem
    Call Logo(picPreview, picLogo)
End Sub

Private Sub sldSpeed_Change()
    Animated.Speed = sldSpeed.Value
End Sub

Private Sub CoordinateSystem()
    If Animated.Play Then Animated.Play = False
    
    step = 10
    w = picGraph.ScaleWidth
    h = picGraph.ScaleHeight
    xmid = w / 2: ymid = h / 2
    Crest = ymid / step
    
    For i = -step To step
        picGraph.Line (0, ymid - i * Crest)-(w, ymid - i * Crest), &H8000000F
        picGraph.CurrentX = xmid + 5
        picGraph.CurrentY = ymid - i * Crest - picGraph.TextHeight(i) / 2
        picGraph.Print i
    Next
    
    picGraph.Line (0, ymid)-(w, ymid), vbBlack
    picGraph.Line (xmid, 0)-(xmid, h), vbBlack
    Set picGraph.Picture = picGraph.Image
End Sub

Private Sub PlotGraph()
    Dim c As Long
    On Error GoTo EvalError
    
    c = RGB(Int(125 * Rnd) + 50, Int(125 * Rnd) + 50, Int(125 * Rnd) + 50)
    
    For i = -xmid To xmid
        Call Script.ExecuteStatement("x=" & Rads(CSng(i)))
        X = i: Y = Crest * Script.Eval(Trim$(txtExpression.text))
        picGraph.PSet (xmid + X, ymid - Y), c
    Next i

EvalError:
    If Err.Number = 1002 Then
        ' Syntax Error
        Call MsgBox(Err.Description)
    Else
        Resume Next
    End If
End Sub

 


