Attribute VB_Name = "Mod_DX"
Option Explicit

Public oldResHeight As Long, oldResWidth As Long
Attribute oldResWidth.VB_VarUserMemId = 1073741836
Public bNoResChange As Boolean
Attribute bNoResChange.VB_VarUserMemId = 1073741838

Public Sub IniciarObjetosDirectX()

    On Error Resume Next

    If MsgBox("¿Desea Reproducir el Juego en Pantalla Completa?", vbQuestion + vbYesNo, "Resolución") = vbYes Then
        NoRes = 0
    Else
        NoRes = 1
    End If


    If NoRes Then
        CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
    Else
        CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
    End If

    If CambiarResolucion Then
        With MidevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = 32
        End With
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
    Else
        bNoResChange = True
    End If

    Call AddtoRichTextBox(frmCargando.Status, "¡DirectX OK!", 255, 150, 50, 1, , False)

    Exit Sub

End Sub

Public Sub LiberarObjetosDX()
    Err.Clear
    On Error GoTo fin:
    Dim loopc As Integer

    Set PrimarySurface = Nothing
    Set PrimaryClipper = Nothing
    Set BackBufferSurface = Nothing

    LiberarDirectSound

    Call SurfaceDB.BorrarTodo

    Set DirectDraw = Nothing

    For loopc = 1 To NumSoundBuffers
        Set DSBuffers(loopc) = Nothing
    Next loopc


    Set Loader = Nothing
    Set Perf = Nothing
    Set Seg = Nothing
    Set DirectSound = Nothing

    Set DirectX = Nothing
    Exit Sub
fin:     LogError "Error producido en Public Sub LiberarObjetosDX()"
End Sub

