Attribute VB_Name = "modMain"
Option Explicit

Public Animated As New clsAnimation
Public Script As New ScriptControl

Public PathLogo As String
Public SoundBuffer() As Byte

Public Sub Main()
    Dim Math As New clsMath
    
    NewGame = True
    Script.Language = "VBScript"
    PathLogo = App.Path & "\LOGO.BMP"

    Call Script.AddObject("Math", Math, True)
    Call frmMain.Show
End Sub

Public Sub BeginPlaySound(ByVal ResourceId As Integer, Optional ByVal SoundLoop As Boolean = False)
    SoundBuffer = LoadResData(ResourceId, "SOUND")
    
    If SoundLoop Then
        Call sndPlaySound(SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_LOOP)
    Else
        Call sndPlaySound(SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
    End If
End Sub

Public Sub CreateObject(objType As Object, ByVal Count As Integer)
    ' Create new objects (control box)
    
    Dim i As Integer
    ' I used this code to create new cards.
    For i = 1 To Count
        If objType.Count > 0 Then
            Call Load(objType(objType.Count))
            objType(objType.Count - 1).Visible = True
            objType(objType.Count - 1).Left = _
                -objType(objType.Count - 1).Width
            objType(objType.Count - 1).Top = _
                -objType(objType.Count - 1).Height
        End If
    Next i
End Sub

Public Sub DestroyObject(objType As Object, ByVal Count As Integer)
    ' Destroy created objects (control box)
    ' using CreateObject Function.
    Dim i As Integer
    
    ' I used this code to remove the cards that
    ' I've created using CreateObject.
    
    For i = 1 To Count
        If objType.Count > 1 Then
            Call Unload(objType(objType.Count - 1))
        End If
    Next i
End Sub

Public Sub EndPlaySound()
    Call sndPlaySound(ByVal vbNullString, 0&)
End Sub

Public Sub Gradient(pBox As PictureBox, Color As Long)
    Dim i As Integer, X As Integer
    
    ' horizontal gradient
    X = 255 / pBox.Height
    For i = 0 To pBox.Height
        Select Case Color
        Case Is = 0
            pBox.Line (0, i)-(pBox.Width, i), RGB(0, 0, i * X)
            frmMain.sbStatusBar.Panels(1).text = "Now playing BLUE..."
        Case Is = 1
            pBox.Line (0, i)-(pBox.Width, i), RGB(i * X, 0, 0)
            frmMain.sbStatusBar.Panels(1).text = "Now playing RED..."
        Case Is = 2
            pBox.Line (0, i)-(pBox.Width, i), RGB(0, i * X, 0)
            frmMain.sbStatusBar.Panels(1).text = "Now playing GREEN..."
        Case Is = 3
            pBox.Line (0, i)-(pBox.Width, i), RGB(i * X, i * X, 0)
            frmMain.sbStatusBar.Panels(1).text = "Now playing YELLOW..."
        End Select
    Next i
End Sub

Public Sub Logo(DstPicBox As PictureBox, SrcPicBox As PictureBox)
    Dim BM As BITMAP
    
    If Dir$(PathLogo) <> "" Then
        With SrcPicBox
            Set .Picture = LoadPicture(PathLogo)
            
            If .Picture Then
                Dim OldAutoRedraw As Integer, OldScalemode As Integer
                OldAutoRedraw = DstPicBox.AutoRedraw
                OldScalemode = DstPicBox.ScaleMode
                DstPicBox.AutoRedraw = True
                DstPicBox.ScaleMode = vbPixels
                Call GetObject(.Picture, Len(BM), BM)
                Dim xmid As Integer, ymid As Integer
                xmid = (DstPicBox.ScaleWidth - BM.bmWidth) / 2
                ymid = (DstPicBox.ScaleHeight - BM.bmHeight) / 2
                Call TransBMP(DstPicBox.hdc, xmid, ymid, _
                    .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HFF00FF)
                Set DstPicBox.Picture = DstPicBox.Image
                DstPicBox.AutoRedraw = OldAutoRedraw
                DstPicBox.ScaleMode = OldScalemode
            End If
        End With
    End If
End Sub

Public Function Pi() As Single
    Pi = 4 * Atn(1)
End Function

Public Function Rads(deg As Single) As Single
    Rads = deg * Pi / 180 ' convert the angle into radians
End Function

Public Function TransBMP(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal TransColor As Long) As Long
    ' create transparent bitmap
    If DstW = 0 Or DstH = 0 Then Exit Function
    
    Dim B As Long, h As Long, F As Long, i As Long
    Dim TmpDC As Long, tmpBMP As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As Long, Data2() As Long
    Dim Info As BITMAPINFO
    
    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    tmpBMP = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, tmpBMP)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 4 - 1)
    ReDim Data2(DstW * DstH * 4 - 1)
    Info.bmiHeader.biSize = Len(Info.bmiHeader)
    Info.bmiHeader.biWidth = DstW
    Info.bmiHeader.biHeight = DstH
    Info.bmiHeader.biPlanes = 1
    Info.bmiHeader.biBitCount = 32
    Info.bmiHeader.biCompression = 0

    Call BitBlt(TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy)
    Call BitBlt(Sr2DC, 0, 0, DstW, DstH, SrcDC, SrcX, SrcY, vbSrcCopy)
    Call GetDIBits(TmpDC, tmpBMP, 0, DstH, Data1(0), Info, 0)
    Call GetDIBits(Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0)
    
    For h = 0 To DstH - 1
        F = h * DstW
        For B = 0 To DstW - 1
            i = F + B
            If (Data2(i) And &HFFFFFF) = TransColor Then
            Else
                Data1(i) = Data2(i)
            End If
        Next B
    Next h

    Call SetDIBitsToDevice(DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0)

    Erase Data1
    Erase Data2
    Call DeleteObject(SelectObject(TmpDC, TmpObj))
    Call DeleteObject(SelectObject(Sr2DC, Sr2Obj))
    Call DeleteDC(TmpDC)
    Call DeleteDC(Sr2DC)
End Function

Public Sub OpenSettings()
    On Error GoTo OpenErr
    Dim Filename As String
    
    Filename = App.Path & "\Setting.dat"
    
    If Dir$(Filename) <> "" Then
        Dim InFile As Integer, s As String
        Dim Speed As Integer
        
        InFile = FreeFile
        Open Filename For Input As InFile
            Input #InFile, Opponents
            Input #InFile, PlayerName(0)
            Input #InFile, PlayerName(1)
            Input #InFile, PlayerName(2)
            Input #InFile, PlayerName(3)
            Input #InFile, Speed
            Input #InFile, AutoSort
            Input #InFile, CardDeck
            Input #InFile, Difficulty
            Input #InFile, AnimationType
            Input #InFile, WaveExpression
        Close InFile
        Animated.Speed = Speed
    Else
        Call DefaultSettings
    End If
    Call UpdateDeck(CardDeck)
    Exit Sub

OpenErr:
    Call MsgBox(Err.Description, vbOKOnly & vbCritical)
End Sub

Public Sub SaveSettings()
    On Error GoTo SaveErr
    Dim Filename As String
    
    Filename = App.Path & "\Setting.dat"
    
    Dim i As Integer, InFile As Integer
    
    InFile = FreeFile
    Open Filename For Output As InFile
        Write #InFile, Opponents
        Write #InFile, PlayerName(0)
        Write #InFile, PlayerName(1)
        Write #InFile, PlayerName(2)
        Write #InFile, PlayerName(3)
        Write #InFile, Animated.Speed
        Write #InFile, AutoSort
        Write #InFile, CardDeck
        Write #InFile, Difficulty
        Write #InFile, AnimationType
        Write #InFile, WaveExpression
    Close InFile
    Exit Sub
    
SaveErr:
    Call MsgBox(Err.Description, vbOKOnly Or vbCritical)
End Sub

Public Sub DefaultSettings()
    Opponents = 1
    PlayerName(0) = "Haru"
    PlayerName(1) = "Elie"
    PlayerName(2) = "Musica"
    PlayerName(3) = "Plue"
    Animated.Speed = 4
    CardDeck = 0
    AutoSort = False
    Difficulty = "NORMAL"
    AnimationType = 0
    WaveExpression = "10*cos(x)"
End Sub

Public Sub UpdateDeck(Deck As Integer)
    Dim cBox As Control
    
    For Each cBox In frmMain.Controls ' read all controls in frmMain
        If TypeName(cBox) = "UnoCard" Then
            ' if control name is equal to UnoCard then change deck
            If cBox.Deck <> Deck Then cBox.Deck = Deck
        End If
    Next cBox
End Sub

Public Function ChkExpression(Expression) As Boolean
    On Error GoTo EvalError
        
    Expression = Trim$(Expression)
    If ChkParen(Expression) Then
        Call Script.Eval(Expression)
        ChkExpression = True
    Else
        ChkExpression = False
    End If
    Exit Function
    
EvalError:
    Call MsgBox(Script.Error.Description, _
        vbOKOnly Or vbInformation, "Invalid Expression!")
End Function

