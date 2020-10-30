VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpriteBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIndexBuffer As Long
Private mVertexBuffer As Long
Private mLastTextureID As Long
Private mPosition As Long

Private mData() As Vertex

Private Sub Class_Initialize()
    mIndexBuffer = mVertexBuffer = mLastTextureID = -1
End Sub

Public Function Initialize(ByVal Capacity As Long) As Boolean
    ReDim mData(4 * Capacity) As Vertex

    mVertexBuffer = modDirectx.CreateVertexBuffer(UBound(mData) * LenB(mData(0)), ByVal 0, D3DUSAGE_DYNAMIC)

    Dim indices() As Integer

    ReDim indices(0 To Capacity * 6 - 1) As Integer

    Dim i As Long, j As Long

    For i = 0 To UBound(indices) Step 6
        indices(i) = j
        indices(i + 1) = j + 1
        indices(i + 2) = j + 2
        indices(i + 3) = j + 2
        indices(i + 4) = j + 3
        indices(i + 5) = j

        j = j + 4
    Next i

    mIndexBuffer = modDirectx.CreateIndexBuffer(UBound(indices) * LenB(indices(0)), VarPtr(indices(0)), D3DFMT_INDEX16, D3DUSAGE_WRITEONLY)

    Initialize = mVertexBuffer <> -1 And mIndexBuffer <> -1
End Function

Public Sub Flush()
    If mPosition > 0 Then
        Call modDirectx.UpdateVertexBuffer(mVertexBuffer, VarPtr(mData(0)), mPosition * LenB(mData(0)))
        Call modDirectx.DrawIndexedPrimitive(D3DPT_TRIANGLELIST, (mPosition \ 4) * 6, 0, (mPosition \ 4) * 2)
        mPosition = 0
    End If
End Sub

Public Sub SetTexture(ByVal handle As Long)
    If mLastTextureID <> handle Then
        If mLastTextureID <> -1 Then _
           Flush

        Call modDirectx.SetTexture(handle, 0)
        mLastTextureID = handle
    End If
End Sub

Public Sub BeginDraw()
    Call modDirectx.SetVertexBuffer(mVertexBuffer)
    Call modDirectx.SetIndexBuffer(mIndexBuffer)
    Call modDirectx.SetTexture(mLastTextureID, 0)

    mPosition = 0
End Sub

Public Sub EndDraw()
    Call Flush
End Sub

Public Sub DrawTexture(ByVal handle As Long, ByVal X As Single, ByVal Y As Single, ByVal left As Long, ByVal top As Long, ByVal right As Long, ByVal bottom As Long, ByVal color As Long)
    Call SetTexture(handle)
    Dim size As Vec2: size = GetTextureSize(handle)

    mData(mPosition).color = MakeVec4(1, 1, 1, 1)
    mData(mPosition).position = MakeVec2(X, Y)
    mData(mPosition).texCooord = MakeVec2(left / size.X, top / size.Y)
    mPosition = mPosition + 1

    mData(mPosition).color = MakeVec4(1, 1, 1, 1)
    mData(mPosition).position = MakeVec2(X, Y + bottom)
    mData(mPosition).texCooord = MakeVec2(left / size.X, (top + bottom) / size.Y)
    mPosition = mPosition + 1

    mData(mPosition).color = MakeVec4(1, 1, 1, 1)
    mData(mPosition).position = MakeVec2(X + right, Y + bottom)
    mData(mPosition).texCooord = MakeVec2((left + right) / size.X, (top + bottom) / size.Y)
    mPosition = mPosition + 1

    mData(mPosition).color = MakeVec4(1, 1, 1, 1)
    mData(mPosition).position = MakeVec2(X + right, Y)
    mData(mPosition).texCooord = MakeVec2((left + right) / size.X, top / size.Y)
    mPosition = mPosition + 1

    If mPosition >= (UBound(mData) / 4) Then _
       Flush
End Sub

Public Sub DrawTextureEx(ByVal handle As Long, ByVal dX As Single, ByVal dY As Single, ByVal dWidth As Long, ByVal dHeight As Long, ByVal sLeft As Long, ByVal sToop As Long, ByVal sRight As Long, ByVal sBottom As Long, ByVal color As Long)
    Call SetTexture(handle)
    Dim size As Vec2: size = GetTextureSize(handle)

    mData(mPosition).color = MakeVec4(1, 1, 1, 1)
    mData(mPosition).position = MakeVec2(dX, dY)
    mData(mPosition).texCooord = MakeVec2(sLeft / size.X, sToop / size.Y)
    mPosition = mPosition + 1

    mData(mPosition).color = MakeVec4(1, 1, 1, 1)
    mData(mPosition).position = MakeVec2(dX, dY + dHeight)
    mData(mPosition).texCooord = MakeVec2(sLeft / size.X, (sToop + sBottom) / size.Y)
    mPosition = mPosition + 1

    mData(mPosition).color = MakeVec4(1, 1, 1, 1)
    mData(mPosition).position = MakeVec2(dX + dWidth, dY + dHeight)
    mData(mPosition).texCooord = MakeVec2((sLeft + sRight) / size.X, (sToop + sBottom) / size.Y)
    mPosition = mPosition + 1

    mData(mPosition).color = MakeVec4(1, 1, 1, 1)
    mData(mPosition).position = MakeVec2(dX + dWidth, dY)
    mData(mPosition).texCooord = MakeVec2((sLeft + sRight) / size.X, sToop / size.Y)
    mPosition = mPosition + 1

    If mPosition >= (UBound(mData) / 4) Then _
       Flush
End Sub


Private Function GetTextureSize(ByVal handle As Long) As Vec2
    Dim W As Long, h As Long, hc As Long
    hc = video_get_texture_info(handle, W, h)

    GetTextureSize.X = W
    GetTextureSize.Y = h
End Function


