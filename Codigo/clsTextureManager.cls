VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextureManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const HASH_TABLE_SIZE As Long = 1024

Private Type SURFACE_ENTRY_DYN
    filename As Long
    textureHandle As Long
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode
Private mGraphicsPath As String


Public Function GetTexture(ByVal filename As Long) As Long
    GetTexture = -1

    If filename = 0 Then
        Debug.Print "0 GRH ATTEMPT TO BE LOADED"
        Exit Function
    End If

    Dim i As Long

    ' Search the index on the list
    With TexList(filename Mod HASH_TABLE_SIZE)

        For i = 1 To .surfaceCount

            If .SurfaceEntry(i).filename = filename Then
                GetTexture = .SurfaceEntry(i).textureHandle
                Exit Function
            End If
        Next i

    End With

    'Not in memory, load it!
    GetTexture = LoadTexture(filename)
End Function

Private Function LoadTexture(ByVal Archivo As Long) As Long
    Dim Index As Long
    Dim handle As Long

    Index = Archivo Mod HASH_TABLE_SIZE

    With TexList(Index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN

        With .SurfaceEntry(.surfaceCount)

            'Nombre
            .filename = Archivo

            Dim Graphic As String
            Graphic = mGraphicsPath & .filename

            If FileExist(Graphic & ".png", vbNormal) Then
                handle = modDirectx.CreateTextureFromFile(Graphic & ".png")
            ElseIf FileExist(Graphic & ".bmp", vbNormal) Then
                handle = modDirectx.CreateTextureFromFile(Graphic & ".bmp")
            Else
                Call MsgBox("No se ha podido cargar el grafico con indice: " & Graphic)
                Call CloseClient
                Exit Function
            End If

            If handle = -1 Then
                Call MsgBox("No se ha podido cargar el grafico con indice: " & Graphic)
                Call CloseClient
                Exit Function
            End If

            .filename = Archivo
            .textureHandle = handle

            LoadTexture = handle
        End With
    End With
End Function


Public Sub Initialize(ByVal graphicPath As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Initializes the object (the DD object is needed to load surfaces)
'**************************************************************
    mGraphicsPath = graphicPath
End Sub

