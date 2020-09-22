VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expression"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3540
      TabIndex        =   6
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3540
      TabIndex        =   5
      Top             =   60
      Width           =   1095
   End
   Begin VB.Frame fraFrame 
      Height          =   2775
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   3435
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2580
         TabIndex        =   8
         Top             =   2280
         Width           =   795
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   1740
         TabIndex        =   4
         Top             =   2280
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         Cancel          =   -1  'True
         Caption         =   "&Edit"
         Height          =   375
         Left            =   900
         TabIndex        =   3
         Top             =   2280
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   2280
         Width           =   795
      End
      Begin VB.ListBox lstExpression 
         Height          =   1425
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   3315
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the expression would you like to used, then click OK."
         Height          =   495
         Index           =   1
         Left            =   60
         TabIndex        =   0
         Top             =   180
         Width           =   3285
      End
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public retExpression As String

Dim Filename As String
Dim ListChanged As Boolean

Private Sub cmdAdd_Click()
    Dim s As String
    
    s = InputBox("Expression: ", "Add", "")
    If s <> "" Then
        If ChkExpression(s) Then
            lstExpression.AddItem (s)
            ListChanged = True
            cmdSave.Enabled = True
        End If
    End If
    lstExpression.SetFocus
End Sub

Private Sub cmdPlay_Click()
    Call Unload(Me)
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdEdit_Click()
    With lstExpression
        If .ListIndex <> -1 Then
            Dim s As String
            s = Trim$(InputBox("New Expression:", "Edit", .List(.ListIndex)))
            If s <> "" And s <> .List(.ListIndex) Then
                If ChkExpression(s) Then
                    Dim OldIndex As Integer
                    OldIndex = .ListIndex
                    Call .RemoveItem(OldIndex)
                    Call .AddItem(s, OldIndex)
                    .ListIndex = OldIndex
                    ListChanged = True
                    cmdSave.Enabled = True
                End If
            End If
        End If
    End With
    lstExpression.SetFocus
End Sub

Private Sub cmdOk_Click()
    With lstExpression
        If .ListIndex <> -1 Then
            retExpression = .List(.ListIndex)
        End If
    End With
    Call Unload(Me)
End Sub

Private Sub cmdRemove_Click()
    With lstExpression
        If .ListIndex <> -1 Then
            Call .RemoveItem(.ListIndex)
            ListChanged = True
            cmdSave.Enabled = True
        End If
        
        .SetFocus
    End With
End Sub

Private Sub cmdSave_Click()
    If ListChanged Then
        Call FileSave
        cmdSave.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    ListChanged = False
    Filename = App.Path & "\List.dat"
    retExpression = ""
    Call FileOpen
End Sub

Private Sub FileOpen()
    On Error GoTo OpenErr
    
    If Dir$(Filename) <> "" Then
        Dim InFile As Integer, s As String
        
        Call lstExpression.Clear
        
        InFile = FreeFile
        Open Filename For Input As InFile
            Do While Not EOF(InFile)
                Input #InFile, s
                Call lstExpression.AddItem(s)
            Loop
        Close InFile
    End If
    Exit Sub

OpenErr:
    Call MsgBox(Err.Description, vbOKOnly & vbCritical)
End Sub

Private Sub FileSave()
    On Error GoTo SaveErr
    
    Dim i As Integer, InFile As Integer
    
    InFile = FreeFile
    Open Filename For Output As InFile
        For i = 0 To lstExpression.ListCount - 1
            Write #InFile, lstExpression.List(i)
        Next i
    Close InFile
    Exit Sub
    
SaveErr:
    Call MsgBox(Err.Description, vbOKOnly Or vbCritical)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdSave.Enabled Then
        Dim ret As Integer
        
        ret = MsgBox("Do you want to save the changes?", _
            vbExclamation Or vbYesNoCancel, "Uno")
        If ret = vbYes Then
            Call FileSave
        ElseIf ret = vbNo Then
        Else
            Cancel = True
        End If
    End If
End Sub
