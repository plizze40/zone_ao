Attribute VB_Name = "MD5"
Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)

Public Function MD5String(P As String) As String

    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(P)
    MDStringFix P, t, r
    MD5String = r
End Function
