VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Uno"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRotXY 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   4020
      Width           =   495
   End
   Begin VB.CommandButton cmdRotXY 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   4020
      Width           =   495
   End
   Begin VB.CommandButton cmdArrow 
      Height          =   375
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4020
      Width           =   495
   End
   Begin VB.CommandButton cmdArrow 
      Height          =   375
      Index           =   0
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4020
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   4020
      Width           =   1455
   End
   Begin VB.PictureBox picViewer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3375
      Left            =   60
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   6
      Top             =   600
      Width           =   5880
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programmed by: Aris Buenaventura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   300
         TabIndex        =   10
         Top             =   1260
         Width           =   4980
      End
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Timer tmrAnimation 
      Interval        =   10
      Left            =   600
      Top             =   600
   End
   Begin VB.Label lblPromp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "or"
      Height          =   195
      Index           =   2
      Left            =   3480
      TabIndex        =   12
      Top             =   300
      Width           =   165
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ajb2001lg@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   3720
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   300
      Width           =   1665
   End
   Begin VB.Label lblPromp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uno Game version 1.0"
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   9
      Top             =   60
      Width           =   1605
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ravemasterharuglory@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1080
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   300
      Width           =   2385
   End
   Begin VB.Label lblPromp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "email:"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   300
      Width           =   435
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FIG_MAXPTS = 14
Const UNO_MAXPTS = 30

Private Type FIG_PTS
    X As Single
    Y As Single
    z As Single
End Type

Private Type UNO_PTS
    X1 As Single: Y1 As Single: Z1 As Single
    X2 As Single: Y2 As Single: Z2 As Single
End Type

Dim Rotation As Integer
Dim FigPoints() As FIG_PTS
Dim UnoPoints() As UNO_PTS
Dim h As Integer, w As Integer
Dim xmid As Integer, ymid As Integer

Dim ThreeD As New clsThreeD

Private Sub cmdArrow_Click(Index As Integer)
    ThreeD.angle = IIf(Index = 0, Rads(10), Rads(-10))
End Sub

Private Sub cmdClose_Click()
    Call Unload(Me)
End Sub

Private Sub cmdRotXY_Click(Index As Integer)
    Select Case Index
    Case Is = 0
        Rotation = 0
        Set cmdArrow(0).Picture = LoadResPicture("DNWRD", vbResBitmap)
        Set cmdArrow(1).Picture = LoadResPicture("UPWRD", vbResBitmap)
    Case Is = 1
        Rotation = 1
        Set cmdArrow(0).Picture = LoadResPicture("BWRD", vbResBitmap)
        Set cmdArrow(1).Picture = LoadResPicture("FWRD", vbResBitmap)
    End Select
End Sub

Private Sub Form_Load()
    tmrAnimation.Enabled = False
    Set Me.Icon = LoadResPicture(101, vbResIcon)
    Set lblEmail(0).MouseIcon = LoadResPicture(101, vbResCursor)
    Set lblEmail(1).MouseIcon = LoadResPicture(101, vbResCursor)
    Call Logo(picViewer, picLogo)
    h = picViewer.ScaleHeight - picViewer.ScaleTop
    w = picViewer.ScaleWidth - picViewer.ScaleLeft
    xmid = w / 2: ymid = h / 2
    Call InitFigure
    Call InitUnoLines
    Call cmdRotXY_Click(1)
    tmrAnimation.Enabled = True
End Sub

Private Sub InitFigure()
    Dim i As Integer, angle As Single
    Dim radius As Single, step As Single
    radius = (h - 30) / 2
    step = 360 / FIG_MAXPTS
    ReDim FigPoints(FIG_MAXPTS) As FIG_PTS
    For i = LBound(FigPoints) To UBound(FigPoints)
        FigPoints(i).X = Cos(Rads(angle)) * radius
        FigPoints(i).Y = Sin(Rads(angle)) * radius
        FigPoints(i).z = 0
        angle = angle + step
    Next i
    Call DrawFigure(xmid, ymid)
End Sub

Private Sub InitUnoLines()
    Dim i As Integer, step As Integer
    ReDim UnoPoints(-UNO_MAXPTS To UNO_MAXPTS) As UNO_PTS
    
    step = xmid / UNO_MAXPTS
    For i = LBound(UnoPoints) To UBound(UnoPoints)
        UnoPoints(i).X1 = i * step
        UnoPoints(i).Y1 = 0
        UnoPoints(i).Z1 = i
        
        UnoPoints(i).X2 = i * step
        UnoPoints(i).Y2 = ymid
        UnoPoints(i).Z2 = i * UNO_MAXPTS
    Next i
    Call DrawUnoLines(xmid, ymid)
End Sub

Private Sub DrawFigure(ByVal X As Single, ByVal Y As Single)
    Dim i As Integer, j As Integer
        
    ' draw the new shape
    For i = 0 To FIG_MAXPTS - 1
        For j = i To FIG_MAXPTS - 1
            picViewer.Line (X + FigPoints(i).X, Y + FigPoints(i).Y)- _
                (X + FigPoints(j).X, Y + FigPoints(j).Y), vbWhite
        Next j
    Next i
End Sub

Private Sub DrawUnoLines(ByVal X As Single, ByVal Y As Single)
    Dim i As Integer, Color As Long
    
    ' draw the new shape
    For i = LBound(UnoPoints) To UBound(UnoPoints)
        Color = IIf(Sgn(i) = 1, vbYellow, vbGreen)
        picViewer.Line (X + UnoPoints(i).X1, Y + UnoPoints(i).Y1)- _
            (X + UnoPoints(i).X2, Y + UnoPoints(i).Y2), Color
        Color = IIf(Sgn(i) = 1, vbRed, vbBlue)
        picViewer.Line (X + UnoPoints(i).X1, Y + -1 * UnoPoints(i).Y1)- _
            (X + UnoPoints(i).X2, Y + -1 * UnoPoints(i).Y2), Color
    Next i
End Sub

Private Sub PlayAnimation()
    Dim i As Integer
    
    picViewer.Cls
    For i = LBound(FigPoints) To UBound(FigPoints)
        If Rotation = 0 Then ' x - axis
            Call ThreeD.rotaboutx(FigPoints(i).X, FigPoints(i).Y, FigPoints(i).z)
        Else                 ' y - axis
            Call ThreeD.rotabouty(FigPoints(i).X, FigPoints(i).Y, FigPoints(i).z)
        End If
    Next i
    For i = LBound(UnoPoints) To UBound(UnoPoints)
        If Rotation = 0 Then
            ' rotate about x axes
            Call ThreeD.rotaboutx(UnoPoints(i).X1, UnoPoints(i).Y1, UnoPoints(i).Z1)
            Call ThreeD.rotaboutx(UnoPoints(i).X2, UnoPoints(i).Y2, UnoPoints(i).Z2)
        Else
            ' rotate about y axes
            Call ThreeD.rotabouty(UnoPoints(i).X1, UnoPoints(i).Y1, UnoPoints(i).Z1)
            Call ThreeD.rotabouty(UnoPoints(i).X2, UnoPoints(i).Y2, UnoPoints(i).Z2)
        End If
    Next i
    Call DrawFigure(xmid, ymid)
    Call DrawUnoLines(xmid, ymid)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrAnimation.Enabled = False
End Sub

Private Sub lblEmail_Click(Index As Integer)
    On Error Resume Next
    Call ShellExecute(0, "open", "mailto:" & lblEmail(Index).Caption, 0, 0, 0)
End Sub

Private Sub tmrAnimation_Timer()
    Call PlayAnimation
End Sub
