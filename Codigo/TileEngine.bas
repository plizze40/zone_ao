Attribute VB_Name = "modTileEngine"
Option Explicit

Public OffsetCounterX As Single
Public OffsetCounterY As Single
Public Movement_Speed As Single
Public Engine_BaseSpeed As Single
Private EndTime As Long

Private Const INFINITE_LOOPS As Integer = -1

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Dim timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Dim engineBaseSpeed As Single

Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521
Public Const SRCCOPY = &HCC0020

Public Type position
    X As Integer
    Y As Integer
End Type

Public Type Position2
    X As Double
    Y As Double
End Type

Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single

    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
End Type

Public Type Grh
    GrhIndex As Integer
    FrameCounter As Double
    SpeedCounter As Byte
    Started As Byte
    Loops As Integer
    Speed As Single
End Type

Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As position
End Type

Public Type HeadData
    Head(1 To 4) As Grh
End Type

Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
End Type

Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type

Public Type FxData
    FxIndex As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

Public Type Char
    Active As Byte
    Heading As Byte
    Pos As position

    Body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean

    Fx As Grh
    FxIndex As Integer

    Criminal As Byte
    Navegando As Byte

    Nombre As String
    GM As Integer

    haciendoataque As Byte


    scrollDirectionX As Long
    scrollDirectionY As Long

    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    ServerIndex As Integer

    pie As Boolean
    muerto As Boolean
    invisible As Boolean

    Movement As Boolean
End Type

Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh

    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte

    Trigger As Integer
End Type

Public Type MapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer


    Changed As Byte
End Type

Public IniPath As String
Public MapPath As String

Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public CurMap As Integer
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As position
Public AddtoUserPos As position
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Long
Public FPSLastCheck As Long
Public FramesPerSecCounter As Long

Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

Public MainViewTop As Integer
Public MainViewLeft As Integer

Public Const TileBufferSize As Long = 6

Public Const TilePixelHeight As Long = 32
Public Const TilePixelWidth As Long = 32

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Public LastTime As Long

Public MainViewWidth As Integer
Public MainViewHeight As Integer

Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh

Public MapData() As MapBlock
Public MapInfo As MapInfo

Public CharList(1 To 10000) As Char

Public bRain As Boolean
Public bRainST As Boolean
Public bTecho As Boolean
Public brstTick As Long

Private RLluvia(7) As RECT
Private iFrameIndex As Byte
Private llTick As Long
Private LTLluvia(4) As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer


Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type


Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
    plFogata = 3
End Enum

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long



Sub CargarCabezas()
    Dim n As Integer, i As Integer, Numheads As Integer, Index As Integer

    Dim Miscabezas() As tIndiceCabeza

    n = FreeFile
    Open App.path & "/Content/Init/Cabezas.ind" For Binary Access Read As #n


    Get #n, , MiCabecera


    Get #n, , Numheads


    ReDim HeadData(0 To Numheads + 1) As HeadData
    ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

    For i = 1 To Numheads
        Get #n, , Miscabezas(i)
        InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
        InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
        InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
        InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
    Next i

    Close #n

End Sub

Sub CargarCascos()
    Dim n As Integer, i As Integer, NumCascos As Integer, Index As Integer

    Dim Miscabezas() As tIndiceCabeza

    n = FreeFile
    Open App.path & "/Content/Init/Cascos.ind" For Binary Access Read As #n


    Get #n, , MiCabecera


    Get #n, , NumCascos


    ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
    ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

    For i = 1 To NumCascos
        Get #n, , Miscabezas(i)
        InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
        InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
        InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
        InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
    Next i

    Close #n

End Sub

Sub CargarCuerpos()
    Dim n As Integer, i As Integer
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo

    n = FreeFile
    Open App.path & "/Content/Init/Personajes.ind" For Binary Access Read As #n


    Get #n, , MiCabecera


    Get #n, , NumCuerpos


    ReDim BodyData(0 To NumCuerpos + 1) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

    For i = 1 To NumCuerpos
        Get #n, , MisCuerpos(i)
        InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
        InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
        InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
        InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
        BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
        BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
    Next i

    Close #n

End Sub
Sub CargarFxs()
    Dim n As Long, i As Long
    Dim NumFxs As Integer
    Dim MisFxs() As tIndiceFx

    n = FreeFile
    Open App.path & "/Content/Init/Fxs.ind" For Binary Access Read As #n


    Get #n, , MiCabecera


    Get #n, , NumFxs

    ReDim FxData(0 To NumFxs + 1) As FxData

    For i = 1 To NumFxs
        Get #n, , FxData(i)
    Next i

    Close #n
End Sub
Sub CargarArrayLluvia()
    Dim n As Integer, i As Integer
    Dim Nu As Integer

    n = FreeFile
    Open App.path & "/Content/init/fk.ind" For Binary Access Read As #n


    Get #n, , MiCabecera


    Get #n, , Nu


    ReDim bLluvia(1 To Nu) As Byte

    For i = 1 To Nu
        Get #n, , bLluvia(i)
    Next i

    Close #n

End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tx As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tx = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub
Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer)
    On Error Resume Next


    If CharIndex > LastChar Then LastChar = CharIndex

    NumChars = NumChars + 1

    If arma = 0 Then arma = 2
    If escudo = 0 Then escudo = 2
    If casco = 0 Then casco = 2

    CharList(CharIndex).Head = HeadData(Head)

    CharList(CharIndex).Body = BodyData(Body)

    If Body > 83 And Body < 88 Then
        CharList(CharIndex).Navegando = 1
    Else: CharList(CharIndex).Navegando = 0
    End If

    CharList(CharIndex).arma = WeaponAnimData(arma)

    CharList(CharIndex).escudo = ShieldAnimData(escudo)
    CharList(CharIndex).casco = CascoAnimData(casco)

    CharList(CharIndex).Heading = Heading


    CharList(CharIndex).Moving = 0
    CharList(CharIndex).MoveOffsetX = 0
    CharList(CharIndex).MoveOffsetX = 0
    CharList(CharIndex).Movement = False


    CharList(CharIndex).Pos.X = X
    CharList(CharIndex).Pos.Y = Y


    CharList(CharIndex).Active = 1


    MapData(X, Y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    CharList(CharIndex).Active = 0
    CharList(CharIndex).Criminal = 0
    CharList(CharIndex).FxIndex = 0
    CharList(CharIndex).invisible = False
    CharList(CharIndex).Moving = 0
    CharList(CharIndex).muerto = False
    CharList(CharIndex).Nombre = ""
    CharList(CharIndex).pie = False
    CharList(CharIndex).Pos.X = 0
    CharList(CharIndex).Pos.Y = 0
    CharList(CharIndex).UsandoArma = False

End Sub

Sub EraseChar(ByVal CharIndex As Integer)
    On Error Resume Next





    CharList(CharIndex).Active = 0


    If CharIndex = LastChar Then
        Do Until CharList(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If


    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

    Call ResetCharInfo(CharIndex)


    NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    If GrhIndex <= 0 Then Exit Sub

    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If


    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If

    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharByHead(CharIndex As Integer, nheading As Byte)
    Dim addX As Integer
    Dim addY As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer

    X = CharList(CharIndex).Pos.X
    Y = CharList(CharIndex).Pos.Y


    Select Case nheading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1

    Case WEST
        addX = -1

    End Select

    nX = X + addX
    nY = Y + addY

    MapData(nX, nY).CharIndex = CharIndex
    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY
    MapData(X, Y).CharIndex = 0

    CharList(CharIndex).MoveOffsetX = -1 * (TilePixelWidth * addX)
    CharList(CharIndex).MoveOffsetY = -1 * (TilePixelHeight * addY)
    CharList(CharIndex).Movement = False

    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nheading

    CharList(CharIndex).scrollDirectionX = addX
    CharList(CharIndex).scrollDirectionY = addY
    If UserEstado <> 1 Then Call DoPasosFx(CharIndex)
End Sub


Public Sub DoFogataFx()
'  If FX = 0 Then 'FIXME
'     If bFogata Then
'        bFogata = HayFogata()
'       If Not bFogata Then frmPrincipal.StopSound
'  Else
'     bFogata = HayFogata()
'    If bFogata Then frmPrincipal.Play "fuego.wav", True
' End If
' End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

    Dim X As Integer, Y As Integer

    For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
        For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1

            If MapData(X, Y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If

        Next X
    Next Y

    EstaPCarea = False

End Function
Public Function TickON(Cual As Integer, Cont As Integer) As Boolean
    Static TickCount(200) As Integer
    If Cont = 999 Then Exit Function
    TickCount(Cual) = TickCount(Cual) + 1
    If TickCount(Cual) < Cont Then
        TickON = False
    Else
        TickCount(Cual) = 0
        TickON = True
    End If
End Function
Sub DoPasosFx(ByVal CharIndex As Integer)
    Static pie As Boolean

    If CharList(CharIndex).Navegando = 0 Then
        If UserMontando And EstaPCarea(CharIndex) And CharIndex = UserCharIndex Then
            If TickON(0, 4) Then Call PlayWaveDS(SND_MONTANDO)
        Else
            If CharList(CharIndex).Criminal = 1 Then Exit Sub
            If Not CharList(CharIndex).muerto And EstaPCarea(CharIndex) Then
                CharList(CharIndex).pie = Not CharList(CharIndex).pie
                If CharList(CharIndex).pie Then
                    Call PlayWaveDS(SND_PASOS1)
                Else
                    Call PlayWaveDS(SND_PASOS2)
                End If
            End If
        End If
    Else: Call PlayWaveDS(SND_NAVEGANDO)
    End If

End Sub
Sub MoveCharByPosAndHead(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)

    On Error Resume Next

    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer



    X = CharList(CharIndex).Pos.X
    Y = CharList(CharIndex).Pos.Y

    MapData(X, Y).CharIndex = 0

    addX = nX - X
    addY = nY - Y


    MapData(nX, nY).CharIndex = CharIndex


    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY

    CharList(CharIndex).MoveOffsetX = -1 * (TilePixelWidth * addX)
    CharList(CharIndex).MoveOffsetY = -1 * (TilePixelHeight * addY)

    CharList(CharIndex).scrollDirectionX = Sgn(addX)
    CharList(CharIndex).scrollDirectionY = Sgn(addY)
    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nheading
End Sub
Sub MoveCharByPos(CharIndex As Integer, nX As Integer, nY As Integer)
    On Error Resume Next

    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nheading As Byte

    X = CharList(CharIndex).Pos.X
    Y = CharList(CharIndex).Pos.Y

    MapData(X, Y).CharIndex = 0

    addX = nX - X
    addY = nY - Y


    If Sgn(addX) = 1 Then nheading = EAST
    If Sgn(addX) = -1 Then nheading = WEST
    If Sgn(addY) = -1 Then nheading = NORTH
    If Sgn(addY) = 1 Then nheading = SOUTH

    MapData(nX, nY).CharIndex = CharIndex

    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY

    CharList(CharIndex).MoveOffsetX = -1 * (TilePixelWidth * addX)
    CharList(CharIndex).MoveOffsetY = -1 * (TilePixelHeight * addY)

    CharList(CharIndex).scrollDirectionX = Sgn(addX)
    CharList(CharIndex).scrollDirectionY = Sgn(addY)
    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nheading

End Sub
Sub MoveCharByPosConHeading(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)
    On Error Resume Next

    If InMapBounds(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y) Then MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

    MapData(nX, nY).CharIndex = CharIndex

    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY

    CharList(CharIndex).Moving = 0
    CharList(CharIndex).MoveOffsetX = 0
    CharList(CharIndex).MoveOffsetY = 0

    CharList(CharIndex).Heading = nheading

End Sub

Sub MoveScreen(Heading As Byte)



    Dim X As Integer
    Dim Y As Integer
    Dim tx As Integer
    Dim tY As Integer


    Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1

    Case WEST
        X = -1

    End Select


    tx = UserPos.X + X
    tY = UserPos.Y + Y


    If tx < MinXBorder Or tx > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else

        AddtoUserPos.X = X
        UserPos.X = tx
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1

    End If




End Sub


Function HayFogata() As Boolean
    Dim j As Integer, k As Integer
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function
Private Function AmigoClan(ByVal CharIndex As Integer) As Boolean
    Dim Nombre1 As String
    Dim Nombre2 As String

    Nombre1 = CharList(UserCharIndex).Nombre
    Nombre2 = CharList(CharIndex).Nombre

    If InStr(Nombre1, "<") > 0 And InStr(Nombre2, "<") > 0 Then

        AmigoClan = Trim$(mid$(Nombre2, InStr(Nombre2, "<"))) = _
                    Trim$(mid$(Nombre1, InStr(Nombre1, "<")))
    End If
End Function

Function NextOpenChar() As Integer
    Dim loopc As Integer

    loopc = 1
    Do While CharList(loopc).Active
        loopc = loopc + 1
    Loop

    NextOpenChar = loopc

End Function
Sub LoadGrhData()
    On Error GoTo ErrorHandler

    Dim Grh As Integer
    Dim Frame As Integer
    Dim tempint As Integer


    ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

    Open IniPath & "Graficos.ind" For Binary Access Read As #1
    Seek #1, 1

    Get #1, , MiCabecera
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint
    Get #1, , tempint

    Get #1, , Grh

    Do Until Grh <= 0


        Get #1, , GrhData(Grh).NumFrames
        If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler

        If GrhData(Grh).NumFrames > 1 Then


            For Frame = 1 To GrhData(Grh).NumFrames

                Get #1, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                    GoTo ErrorHandler
                End If

            Next Frame

            Get #1, , GrhData(Grh).Speed
            If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler


            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler

            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
            If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler

        Else


            Get #1, , GrhData(Grh).FileNum
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler

            Get #1, , GrhData(Grh).sX
            If GrhData(Grh).sX < 0 Then GoTo ErrorHandler

            Get #1, , GrhData(Grh).sY
            If GrhData(Grh).sY < 0 Then GoTo ErrorHandler

            Get #1, , GrhData(Grh).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            Get #1, , GrhData(Grh).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler


            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth

            GrhData(Grh).Frames(1) = Grh

        End If


        Get #1, , Grh

    Loop


    Close #1

    Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(X As Integer, Y As Integer) As Boolean





    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        LegalPos = False
        Exit Function
    End If


    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If


    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If

    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If

    LegalPos = True

End Function

Function LegalPosMuerto(X As Integer, Y As Integer) As Boolean





    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        LegalPosMuerto = False
        Exit Function
    End If


    If MapData(X, Y).Blocked = 1 Then
        LegalPosMuerto = False
        Exit Function
    End If


    If MapData(X, Y).CharIndex > 0 Then
        If CharList(MapData(X, Y).CharIndex).muerto = True Then
            LegalPosMuerto = False
            Exit Function
        End If
    End If

    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    End If

    LegalPosMuerto = True

End Function




Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean





    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapLegalBounds = False
        Exit Function
    End If

    InMapLegalBounds = True

End Function

Function InMapBounds(ByVal X As Long, ByVal Y As Long) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        InMapBounds = False
        Exit Function
    End If

    InMapBounds = True

End Function

Sub DDrawGrhtoSurface(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, Center As Byte, Animate As Byte)

    Dim CurrentGrh As Grh
    Dim destRect As RECT
    Dim SourceRect As RECT

    If Animate Then
        If Grh.Started = 1 Then
            If Grh.SpeedCounter > 0 Then
                Grh.SpeedCounter = Grh.SpeedCounter - 1
                If Grh.SpeedCounter = 0 Then
                    Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                    Grh.FrameCounter = Grh.FrameCounter + (1 / (8 / Velocidad))
                    If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                        Grh.FrameCounter = 1
                    End If
                End If
            End If
        End If
    End If

    CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    If Center Then
        If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16
        End If
        If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32
        End If
    End If

End Sub

Sub DDrawTransGrhIndextoSurface(Grh As Integer, ByVal X As Integer, ByVal Y As Integer, Center As Byte, Animate As Byte)
    Dim CurrentGrh As Grh
    Dim destRect As RECT
    Dim SourceRect As RECT

    With destRect
        .left = X
        .top = Y
        .right = .left + GrhData(Grh).pixelWidth
        .bottom = .top + GrhData(Grh).pixelHeight
    End With

    ' If destRect.left >= 0 And destRect.top >= 0 And destRect.right <= SurfaceDesc.lWidth And destRect.bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .left = GrhData(Grh).sX
        .top = GrhData(Grh).sY
        .right = .left + GrhData(Grh).pixelWidth
        .bottom = .top + GrhData(Grh).pixelHeight
    End With
    ' End If
End Sub

Sub DrawGrh222(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, Center As Boolean, Animate As Boolean, Optional ByVal color As Long = -1)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'
'
'*****************************************************************
    On Error GoTo ErrHandler
    Dim CurrentGrhIndex As Long
    Dim FrameDuration As Single
    Dim TextureID As Long

    If Grh.GrhIndex = 0 Then Exit Sub

    If False Then
        If Grh.Started = 1 Then
            If Grh.SpeedCounter > 0 Then
                Grh.SpeedCounter = Grh.SpeedCounter - 1
                If Grh.SpeedCounter = 0 Then
                    Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                    Grh.FrameCounter = Grh.FrameCounter + (1 / (8 / Velocidad))
                    If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                        Grh.FrameCounter = 1
                    End If
                End If
            End If
        End If
    End If

    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    If CurrentGrhIndex = 0 Then Exit Sub

    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * 16) + 16
            End If

            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * 32) + 32
            End If
        End If

        'draw
        TextureID = SurfaceDB.GetTexture(.FileNum)
        Call SpriteBatchDrawTexture(TextureID, MakeVec2(X, Y), MakeRect(.sX, .sY, .pixelWidth, .pixelHeight), color)
    End With
    Exit Sub
ErrHandler:
    Debug.Print "DrawGrh: " & Err.Description
End Sub

Sub DrawGrh(ByRef Grh As Grh, ByVal X As Single, ByVal Y As Single, ByVal Center As Boolean, ByVal Animate As Boolean, Optional ByVal color As Long = -1)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'*****************************************************************
    Dim CurrentGrhIndex As Long
    Dim FrameDuration As Single
    Dim Texture As Long

    If Grh.GrhIndex = 0 Then Exit Sub

    On Error GoTo error
    If Animate Then
        If Grh.Started = 1 Then
            FrameDuration = Grh.Speed / GrhData(Grh.GrhIndex).NumFrames
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime / FrameDuration) * Movement_Speed

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1

                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        ElseIf Grh.FrameCounter > 1 Then
            FrameDuration = Grh.Speed / GrhData(Grh.GrhIndex).NumFrames
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime / FrameDuration) * Movement_Speed

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = 1
            End If
        End If
    End If


    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth - TilePixelWidth) \ 2
            End If

            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        Texture = SurfaceDB.GetTexture(.FileNum)
        Call SpriteBatchDrawTexture(Texture, MakeVec2(X, Y), MakeRect(.sX, .sY, .pixelWidth, .pixelHeight), color)
    End With

    Exit Sub

error:

    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    ElseIf Err.Number = 9 Then
        Debug.Print "Posible Sub Indice fuera del intervalo en DrawGrh id " & Grh.GrhIndex
        Grh.GrhIndex = 0
    Else
        'Call Log_Engine("Error in Draw_Grh, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical, Err.Number
        Call CloseClient
    End If
End Sub

Sub DrawGrhtoHdc(hwnd As Long, hDC As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
'FIXME
' If Grh <= 0 Then Exit Sub

' SecundaryClipper.SetHWnd hWnd
'   SurfaceDB.GetBMP(GrhData(Grh).FileNum).BltToDC Hdc, SourceRect, destRect

End Sub
Sub PlayWaveAPI(File As String)
    Dim rc As Integer

    rc = sndPlaySound(File, SND_ASYNC)

End Sub
Public Sub RenderScreen(ByVal tilex As Long, ByVal tiley As Long, ByVal PixelOffsetX As Long, ByVal PixelOffsetY As Long)
'**************************************************************
'
'
'
'**************************************************************
    Dim Y As Long             'Keeps track of where on map we are
    Dim X As Long             'Keeps track of where on map we are

    Dim screenminY As Integer    'Start Y pos on current screen
    Dim screenmaxY As Integer    'End Y pos on current screen
    Dim screenminX As Integer    'Start X pos on current screen
    Dim screenmaxX As Integer    'End X pos on current screen

    Dim minY As Integer             'Start Y pos on current map
    Dim maxY As Integer             'End Y pos on current map
    Dim minX As Integer             'Start X pos on current map
    Dim maxX As Integer             'End X pos on current map

    Dim ScreenX As Integer    'Keeps track of where to place tile on screen
    Dim ScreenY As Integer    'Keeps track of where to place tile on screen

    Dim minXOffset As Integer
    Dim minYOffset As Integer

    Dim PixelOffsetXTemp As Integer    'For centering grhs
    Dim PixelOffsetYTemp As Integer    'For centering grhs

    Dim ElapsedTime As Single

    ElapsedTime = Engine_ElapsedTime()

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth

    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize * 2    ' WyroX: Parche para que no desaparezcan techos y arboles
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize

    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If

    If maxY > YMaxMapSize Then maxY = YMaxMapSize

    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If

    If maxX > XMaxMapSize Then maxX = XMaxMapSize

    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If

    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1

    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If

    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1


    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX

            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY

            'Layer 1 **********************************
            If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
                Call DrawGrh(MapData(X, Y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, True, True)
            End If
            '******************************************

            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DrawGrh(MapData(X, Y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, True, True)
            End If
            '******************************************

            ScreenX = ScreenX + 1
        Next

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next


    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - TileBufferSize

    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
            If InMapBounds(X, Y) Then

                PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY

                With MapData(X, Y)
                    'Object Layer **********************************
                    If .ObjGrh.GrhIndex <> 0 Then
                        Call DrawGrh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, True, True)
                    End If
                    '***********************************************

                    'Char layer********************************
                    If .CharIndex <> 0 Then
                        Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                    End If
                    '*************************************************

                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex <> 0 Then
                        Call DrawGrh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, True, True)
                    End If
                End With
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y


    '<----- Layer 4 ----->
    ScreenY = minYOffset - TileBufferSize

    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY

            'Layer 4
            If MapData(X, Y).Graphic(4).GrhIndex Then
                If Not bTecho Then _
                   Call DrawGrh(MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, True, True)
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y

End Sub

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 16/09/2010 (Zama)
'Draw char's to screen without offcentering them
'16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long

    With CharList(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame

                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .arma.WeaponWalk(.Heading).Started = 1
                .escudo.ShieldWalk(.Heading).Started = 1

                'Char moved
                moved = True

                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If

            End If

            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame

                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .arma.WeaponWalk(.Heading).Started = 1
                .escudo.ShieldWalk(.Heading).Started = 1

                'Char moved
                moved = True

                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If

            End If
        End If

        'If done moving stop animation
        If Not moved Then
            'Stop animations

            '//Evito runtime
            If Not .Heading <> 0 Then .Heading = EAST

            .Body.Walk(.Heading).Started = 0

            '//Movimiento del arma y el escudo
            If Not .Movement Then
                .arma.WeaponWalk(.Heading).Started = 0

                .escudo.ShieldWalk(.Heading).Started = 0

            End If

            .Moving = False

        End If

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        Movement_Speed = 0.0025

        If Not .invisible Then


            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
               Call DrawGrh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, True, True)

            'Draw Head
            If .Head.Head(.Heading).GrhIndex Then
                Call DrawGrh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, True, False)

                'Draw Helmet
                If .casco.Head(.Heading).GrhIndex Then _
                   Call DrawGrh(.casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + 0, True, False)

                'Draw Weapon
                If .arma.WeaponWalk(.Heading).GrhIndex Then _
                   Call DrawGrh(.arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, True, True)

                'Draw Shield
                If .escudo.ShieldWalk(.Heading).GrhIndex Then _
                   Call DrawGrh(.escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, True, True)


                'Draw name over head
                If LenB(.Nombre) > 0 Then
                    If Nombres Then
                        Pos = getTagPosition(.Nombre)
                        'Pos = InStr(.Nombre, "<")
                        'If Pos = 0 Then Pos = Len(.Nombre) + 2
                        color = ColorRGB(0, 0, 0)


                        'Nick
                        line = left$(.Nombre, Pos - 2)
                        Call DrawText(PixelOffsetX + 16, PixelOffsetY + 30, line, color, True)

                        'Clan
                        line = mid$(.Nombre, Pos)
                        Call DrawText(PixelOffsetX + 16, PixelOffsetY + 45, line, color, True)
                    End If
                End If
            End If
        End If

        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, CharIndex)     '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo


        'Draw FX
        If .FxIndex <> 0 Then
            Call DrawGrh(.Fx, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, True, True, ColorRGBA(255, 255, 255, 150))

            'Check if animation is over
            If .Fx.Started = 0 Then _
               .FxIndex = 0
        End If
    End With
End Sub


Public Function RenderSounds()
'FIXME
'    If bLluvia(UserMap) = 1 Then
'        If bRain Then
'            If bTecho Then
'                If frmPrincipal.IsPlaying <> plLluviain Then
'                    Call frmPrincipal.StopSound
'                    Call frmPrincipal.Play("lluviain.wav", True)
'                   frmPrincipal.IsPlaying = plLluviain
'               End If


'Else
''   ' If frmPrincipal.IsPlaying <> plLluviaout Then
'  Call frmPrincipal.StopSound
' Call frmPrincipal.Play("lluviaout.wav", True)
' frmPrincipal.IsPlaying = plLluviaout
'End If


'End If
'End If
'End If

End Function


Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

    If GrhIndex > 0 Then

        HayUserAbajo = _
        CharList(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                       And CharList(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                       And CharList(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                       And CharList(UserCharIndex).Pos.Y <= Y

    End If

End Function



Function PixelPos(X As Integer) As Integer




    PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function


Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(App.path & "\Content\Textures\")
End Sub

Public Sub InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer)
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    On Error GoTo ErrHandler

    IniPath = App.path & "\Content\Init\"

    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder

    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft

    WindowTileHeight = Round(frmPrincipal.MainView.height / 32, 0)
    WindowTileWidth = Round(frmPrincipal.MainView.width / 32, 0)

    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2

    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

    MainViewWidth = (TilePixelWidth * WindowTileWidth)
    MainViewHeight = (TilePixelHeight * WindowTileHeight)

    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8

    FPS = 101
    FramesPerSecCounter = 101

    Engine_BaseSpeed = 0.018


    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call LoadGraphics
    Call CargarAnimsExtra
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos

    Exit Sub
ErrHandler:
    Call LogError("Error en InitTileEngine: " & Err.Number & "- " & Err.Description)
    Call CloseClient
End Sub


Sub CrearGrh(GrhIndex As Integer, Index As Integer)
    ReDim Preserve Grh(1 To Index) As Grh
    Grh(Index).FrameCounter = 1
    Grh(Index).GrhIndex = GrhIndex
    Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
    Grh(Index).Started = 1
End Sub

Sub CargarAnimsExtra()
    Call CrearGrh(6580, 1)
    Call CrearGrh(534, 2)
End Sub

Function ControlVelocidad(ByVal LastTime As Long) As Boolean
    ControlVelocidad = (GetTickCount - LastTime > 20)
End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
    Dim buf As Integer
    buf = InStr(Nick, "<")
    If buf > 0 Then
        getTagPosition = buf
        Exit Function
    End If
    buf = InStr(Nick, "[")
    If buf > 0 Then
        getTagPosition = buf
        Exit Function
    End If
    getTagPosition = Len(Nick) + 2
End Function


Public Sub ShowNextFrame(ByVal DisplayFormTop As Integer, _
                         ByVal DisplayFormLeft As Integer, _
                         ByVal MouseViewX As Integer, _
                         ByVal MouseViewY As Integer)
'***************************************************
'
'
'
'***************************************************
    If EngineRun Then
        Call modDirectx.BeginScene
        Call Batch.BeginDraw
        Call modDirectx.ClearBuffer(0)
        'Call modDirectx.BeginRender
        'Call modDirectx.video_set_shader_technique(ShaderHandler, "Technique1")


        If UserMoving Then
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If

            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If

        '
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)


        If ModoTrabajo Then Call DrawText(260, 260, "MODO TRABAJO", vbRed)
        If TaInvi > 20 Then Call DrawText(260, 275, "TIEMPO INVISIBLE " & Int(TaInvi / 30), vbWhite)
        Call Dialogos.Render
        Call RenderSounds
        If Cartel Then Call DibujarCartel

        'Calculamos los FPS y los mostramos
        Call Engine_Update_FPS

        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed

        'Call modDirectx.EndRender
        '

        Call Batch.EndDraw
        Call modDirectx.EndScene
        Call modDirectx.SwapBuffer
    End If
End Sub


Public Sub Engine_Update_FPS()
'***************************************
'Author: Standelf
'Last Modification: 09/09/2019
'Calculate $ Limitate (if active) FPS.
'***************************************

'If ClientSetup.LimiteFPS Then
'    While (GetTickCount - FPSLastCheck) \ 10 < FramesPerSecCounter
'        Call Sleep(5)
'    Wend
'End If

    If FPSLastCheck + 1000 < timeGetTime Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        FPSLastCheck = timeGetTime

    Else
        FramesPerSecCounter = FramesPerSecCounter + 1

    End If

End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim Start_Time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If

    'Get current time
    Call QueryPerformanceCounter(Start_Time)

    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000

    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Function Engine_ElapsedTime() As Long
'**************************************************************
'Gets the time that past since the last call
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ElapsedTime
'**************************************************************

    Dim Start_Time As Long

    'Get current time
    Start_Time = timeGetTime

    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - EndTime

    'Get next end time
    EndTime = Start_Time

End Function

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal Fx As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With CharList(CharIndex)    'fixme
        .FxIndex = Fx

        If .FxIndex > 0 Then
            Call InitGrh(.Fx, FxData(Fx).FxIndex)

            'FIXME
            .Fx.Loops = IIf(Loops > 0, Loops - 1, 0)
        End If
    End With
End Sub

Public Sub SpriteBatchDrawTexture(ByVal handle As Long, Pos As Vec2, RECT As RECT, ByVal color As Long)
    Call Batch.DrawTexture(handle, Pos.X, Pos.Y, RECT.left, RECT.top, RECT.right, RECT.bottom, color)
End Sub


Public Sub SpriteBatchDrawTextureEx(ByVal handle As Long, dest As RECT, src As RECT, ByVal color As Long)
    Call Batch.DrawTextureEx(handle, dest.left, dest.top, dest.right, dest.bottom, src.left, src.top, src.right, src.bottom, color)
End Sub

