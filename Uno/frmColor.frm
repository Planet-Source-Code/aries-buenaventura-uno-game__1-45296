VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Chooser"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      Top             =   1080
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   915
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Color"
      Height          =   1395
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   1755
      Begin VB.OptionButton optColor 
         Caption         =   "Yellow"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   1155
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Green"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   2
         Top             =   780
         Width           =   1155
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Red"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1155
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Blue"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Top             =   180
         Width           =   1155
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   315
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   780
         Width           =   315
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   315
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   315
      End
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a color for the wild card you just played, then click OK button."
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   180
      Width           =   1920
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Color As Integer

Private Sub cmdCancel_Click()
    Color = -1
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
    Call Unload(Me)
End Sub

Private Sub optColor_Click(Index As Integer)
    Color = Index
End Sub
