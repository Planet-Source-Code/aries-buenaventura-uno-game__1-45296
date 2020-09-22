VERSION 5.00
Begin VB.Form frmScore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Score"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Frame frmScore 
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   3975
      Begin VB.Frame fraTotalScore 
         Caption         =   "Total Score"
         Height          =   1515
         Left            =   2520
         TabIndex        =   3
         Top             =   180
         Width           =   1275
         Begin VB.Label lblTotalScore 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   15
            Top             =   1140
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblTotalScore 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   14
            Top             =   900
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblTotalScore 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   13
            Top             =   660
            Width           =   915
         End
         Begin VB.Label lblTotalScore 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   12
            Top             =   420
            Width           =   915
         End
      End
      Begin VB.Frame fraHandTotal 
         Caption         =   "Hand Total"
         Height          =   1515
         Left            =   1320
         TabIndex        =   2
         Top             =   180
         Width           =   1275
         Begin VB.Label lblHandTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   1140
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblHandTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   900
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblHandTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   660
            Width           =   915
         End
         Begin VB.Label lblHandTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   420
            Width           =   915
         End
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   840
         Width           =   105
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   600
         Width           =   105
      End
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Winner As Integer

Private Sub cmdOk_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To Opponents
        lblName(i).Caption = PlayerName(i)
        lblHandTotal(i).Caption = PlayerScore(i)
    Next i
    
    lblName(0).Caption = lblName(0).Caption & " (You)"
    If Opponents = 2 Then
        lblName(2).Visible = True
        lblHandTotal(2).Visible = True
        lblTotalScore(2).Visible = True
    ElseIf Opponents = 3 Then
        lblName(2).Visible = True
        lblName(3).Visible = True
        lblHandTotal(2).Visible = True
        lblTotalScore(2).Visible = True
        lblHandTotal(3).Visible = True
        lblTotalScore(3).Visible = True
    End If
        
    Dim s As String, sum As Integer
    
    sum = 0
    For i = 0 To Opponents
        sum = sum + CInt(lblHandTotal(i).Caption)
    Next i
    
    If PlayerScore(0) = 0 Then
        lblTotalScore(0).Caption = sum
    ElseIf PlayerScore(1) = 0 Then
        lblTotalScore(1).Caption = sum
    ElseIf PlayerScore(2) = 0 Then
        lblTotalScore(2).Caption = sum
    ElseIf PlayerScore(3) = 0 Then
        lblTotalScore(3).Caption = sum
    End If
    
    If Winner = 0 Then
        s = "** Congratulations. You're the man!!! **"
    Else
        s = "** " & PlayerName(Winner) & " WIN!!! **"
    End If
    
    lblPrompt.Caption = s
    lblName(Winner).ForeColor = vbRed
    lblHandTotal(Winner).ForeColor = vbRed
    lblTotalScore(Winner).ForeColor = vbRed
End Sub
