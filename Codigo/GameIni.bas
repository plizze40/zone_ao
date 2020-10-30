Attribute VB_Name = "GameIni"
Option Explicit

Public Type tCabecera
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    Fx As Byte
    tip As Byte
    Password As String
    name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type


Public Type tRenderMods
    sName As String * 7
    bUseVideo As Long
    bNoCostas As Long
    bUseMMX As Long
    bNoAlpha As Long
    bNoTScan As Long
    bNoMusic As Long
    bNoSound As Long
    iImageSize As Long
End Type

Public Type AutoUpdate
    version As Long
    Fase As Byte
End Type

Public RenderMod As tRenderMods


Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni
Public AUpdate As AutoUpdate
Public Sub EscribirUpdate(ByRef Update As AutoUpdate)
    Dim n As Integer
    n = FreeFile
    Open App.path & "/Content/Init/AutoUpdate.con" For Binary As #n

    Put #n, , Update
    Close #n
End Sub
Public Function LeerAutoUpdate() As AutoUpdate
    Dim n As Integer
    Dim Up As AutoUpdate
    n = FreeFile
    Open App.path & "/Content/Init/AutoUpdate.con" For Binary As #n

    Get #n, , Up

    Close #n
    LeerAutoUpdate = Up
End Function
Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
    Dim n As Integer
    Dim GameIni As tGameIni
    n = FreeFile
    Open App.path & "/Content/Init/Inicio.con" For Binary As #n
    Get #n, , MiCabecera

    Get #n, , GameIni

    Close #n
    LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
    Dim n As Integer
    n = FreeFile
    Open App.path & "/Content/Init/Inicio.con" For Binary As #n
    Put #n, , MiCabecera
    Put #n, , GameIniConfiguration
    Close #n
End Sub

Public Sub LoadGameInit()

End Sub

Public Sub SaveGameInit()

End Sub

Public Sub SaveClientSetup()

End Sub
