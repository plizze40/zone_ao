Attribute VB_Name = "modMath"
Option Explicit

Public Type Vec2
    X As Single
    Y As Single
End Type

Public Type Vec3
    X As Single
    Y As Single
    Z As Single
End Type

Public Type Vec4
    X As Single
    Y As Single
    Z As Single
    W As Single
End Type

Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type Matrix
    M11 As Single
    M12 As Single
    M13 As Single
    M14 As Single

    M21 As Single
    M22 As Single
    M23 As Single
    M24 As Single

    M31 As Single
    M32 As Single
    M33 As Single
    M34 As Single

    M41 As Single
    M42 As Single
    M43 As Single
    M44 As Single
End Type

Public Type Vertex
    position As Vec2
    texCooord As Vec2
    color As Vec4
End Type

Public Declare Sub matrix_ortho_off_centerRH Lib "utils.dll" (matrixIn As Matrix, ByVal l As Single, ByVal r As Single, ByVal b As Single, ByVal t As Single, ByVal zn As Single, ByVal zf As Single)
Public Declare Sub matrix_ortho_off_centerLH Lib "utils.dll" (matrixIn As Matrix, ByVal l As Single, ByVal r As Single, ByVal b As Single, ByVal t As Single, ByVal zn As Single, ByVal zf As Single)

Public Function MakeVec2(ByVal X As Single, ByVal Y As Single) As Vec2
    MakeVec2.X = X
    MakeVec2.Y = Y
End Function

Public Function MakeVec3(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Vec3
    MakeVec3.X = X
    MakeVec3.Y = Y
    MakeVec3.Z = Z
End Function

Public Function MakeVec4(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal W As Single) As Vec4
    MakeVec4.X = X
    MakeVec4.Y = Y
    MakeVec4.Z = Z
    MakeVec4.W = W
End Function

Public Function MakeRect(ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal h As Long) As RECT
    MakeRect.left = X
    MakeRect.top = Y
    MakeRect.right = W
    MakeRect.bottom = h
End Function

Public Function ColorRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    ColorRGB = -1
End Function

Public Function ColorRGBA(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
    ColorRGBA = -1
End Function
