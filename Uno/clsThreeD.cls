VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThreeD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public angle As Single

Dim oldx As Single, oldy As Single, oldz As Single

Public Sub rotabouty(ByRef X As Single, ByRef Y As Single, ByRef z As Single)
    oldx = X
    oldy = Y
    oldz = z
    
    X = (oldx * Cos(angle)) - (oldz * Sin(angle))
    Y = Y
    z = (oldx * Sin(angle)) + (oldz * Cos(angle))
End Sub

Public Sub rotaboutx(ByRef X As Single, ByRef Y As Single, ByRef z As Single)
    oldx = X
    oldy = Y
    oldz = z
    
    X = X
    Y = (oldz * Sin(angle)) + (oldy * Cos(angle))
    z = (oldz * Cos(angle)) - (oldy * Sin(angle))
End Sub

Public Sub rotaboutz(ByRef X As Single, ByRef Y As Single, ByRef z As Single)
    oldx = X
    oldy = Y
    oldz = z
    
    X = (oldx * Sin(angle)) + (oldy * Cos(angle))
    Y = (oldx * Cos(angle)) - (oldy * Sin(angle))
    z = z
End Sub

Private Sub Class_Initialize()
    angle = Rads(10) ' start at 10 degs.
End Sub
