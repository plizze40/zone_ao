VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStdPictureLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const HASH_TABLE_SIZE As Long = 512

Private Type SURFACE_ENTRY_DYN
    filename As String
    image As StdPicture
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode
Private mLoader As New StdPictureEx
Private mResourcePath As String

Public Sub Initialize(ByVal path As String)
    mResourcePath = path
End Sub


Public Function LoadStdPicture(ByVal filename As String, Optional ByVal path As String = vbNullString) As StdPicture
    Dim hash As Long: hash = Len(filename)
    Dim i As Long

    With TexList(hash Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).filename = filename Then
                Set LoadStdPicture = .SurfaceEntry(i).image
                Exit Function
            End If
        Next i
    End With

    Set LoadStdPicture = LoadPicture(filename, path)
End Function

Private Function LoadPicture(ByVal filename As String, ByVal path As String) As StdPicture
    On Error GoTo ErrHandler
    Dim Index As Long
    Dim Picture As String

    Index = Len(filename) Mod HASH_TABLE_SIZE

    With TexList(Index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN

        With .SurfaceEntry(.surfaceCount)
            Picture = IIf(path <> vbNullString, path & filename, mResourcePath & filename)
            Debug.Print Picture

            If Not FileExist(Picture, vbNormal) Then
                Call MsgBox("No se ha podido cargar el grafico: " & filename)
                Call CloseClient
                Exit Function
            End If

            'Nombre
            .filename = filename
            Set .image = StdPictureEx.LoadPictureEx(Picture)

            Set LoadPicture = .image
        End With
    End With

    Exit Function
ErrHandler:
    Call MsgBox("No se ha podido cargar el grafico: " & filename)
    Call CloseClient
End Function

