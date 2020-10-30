Attribute VB_Name = "modApplication"
Option Explicit

Public prgRun As Boolean

Sub Main()

    Dim f As Boolean
    Dim loopc As Long
    Dim vsync As Long

    Dim OffsetCounterX As Double
    Dim OffsetCounterY As Double

    FrmIntro.Hide

    If MsgBox("Quieres cambiar la resolucion a 800x600?", vbYesNo, "Resolucion") = vbYes Then
        Call modResolution.SetResolution
    End If

    If MsgBox("Quieres activar la sincronización vertical?(Se recomienda no activarla si tiene una pc de bajos recursos)", vbYesNo, "Vsync") = vbYes Then
        vsync = 1
    End If


    Call WriteClientVer
    Call LoadConst
    Call CargarMensajes
    Call EstablecerRecompensas
    Call InitializeSoundModule

    Dialogos.font = frmPrincipal.font

    CartelOcultarse = Val(GetVar(App.path & "/Content/Init/Opciones.dat", "CARTELES", "Ocultarse"))
    CartelMenosCansado = Val(GetVar(App.path & "/Content/Init/Opciones.dat", "CARTELES", "MenosCansado"))
    CartelVestirse = Val(GetVar(App.path & "/Content/Init/Opciones.dat", "CARTELES", "Vestirse"))
    CartelNoHayNada = Val(GetVar(App.path & "/Content/Init/Opciones.dat", "CARTELES", "NoHayNada"))
    CartelRecuMana = Val(GetVar(App.path & "/Content/Init/Opciones.dat", "CARTELES", "RecuMana"))
    CartelSanado = Val(GetVar(App.path & "/Content/Init/Opciones.dat", "CARTELES", "Sanado"))
    NoRes = Val(GetVar(App.path & "/Content/Init/Opciones.dat", "CONFIG", "ModoVentana"))

    If App.PrevInstance Then
        Call MsgBox("¡Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If

    ChDrive App.path
    ChDir App.path


    Dim fMD5HushYo As String * 32
    HushYo = GenHash(App.path & "\" & App.EXEName & ".exe")


    If FileExist(App.path & "/Content/Init/Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If


    tipf = Config_Inicio.tip
    UserParalizado = False
    frmConectar.Visible = True
    LastTime = GetTickCount
    LoopMidi = True
    UserMap = 1
    PrimeraVez = True
    prgRun = True
    Pausa = False
    FramesPerSec = 60
    FramesPerSecCounter = 30
    LastTime = GetTickCount

    ENDL = Chr(13) & Chr(10)
    ENDC = Chr(1)

    Call modDirectx.Initialize(frmPrincipal.MainView.hwnd, frmPrincipal.MainView.width, frmPrincipal.MainView.height, vsync)
    Call InitTileEngine(frmPrincipal.hwnd, frmPrincipal.MainView.top, frmPrincipal.MainView.left, 32, 32, 13, 17, 9)

    If Musica = 0 Then
        Call CargarMIDI(DirMidi & MIdi_Inicio & ".mid")
        Play_Midi
    End If

    lFrameTimer = GetTickCount

    Do While prgRun
        If RequestPosTimer > 0 Then
            RequestPosTimer = RequestPosTimer - 1
            If RequestPosTimer = 0 Then
                Call SendData("RPU")
            End If
        End If

        Call RefreshAllChars
        Call ShowNextFrame(frmPrincipal.top, frmPrincipal.left, 0, 0)

        'If (GetTickCount - LastTime > 20) Then
        If Not Pausa And frmPrincipal.Visible And Not frmForo.Visible Then
            CheckKeys
            LastTime = GetTickCount
        End If
        'End If

        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            If FPSFLAG Then frmPrincipal.fpstext.Caption = FPS

            lFrameTimer = GetTickCount
        End If


        Call UpdateTimers
        DoEvents
    Loop

    Call CloseClient
End Sub

Private Sub UpdateTimers()
    Static ulttick As Long, esttick As Long
    Static timers(1 To 5) As Long
    Dim loopc As Long

    esttick = GetTickCount
    For loopc = 1 To UBound(timers)


        timers(loopc) = timers(loopc) + (esttick - ulttick)

        If timers(1) >= tUs Then
            timers(1) = 0
            NoPuedeUsar = False
        End If
    Next loopc
    ulttick = GetTickCount
End Sub




Private Sub UpdateScreen(ByVal deltaTime As Single)

End Sub
