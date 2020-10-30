Attribute VB_Name = "modFonts"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                               (Destination As Any, Source As Any, ByVal length As Long)

Private Type CharVA
    X As Integer
    Y As Integer
    W As Integer
    h As Integer

    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single
End Type

Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte    'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Long             'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
End Type

Private cfonts(1 To 2) As CustomFont    ' _Default2 As CustomFont

Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, _
                               ByVal Text As String, _
                               ByVal X As Long, _
                               ByVal Y As Long, _
                               ByVal color As Long, _
                               Optional ByVal Center As Boolean = False, _
                               Optional font As Integer = 1)

'*****************************************************************
'Render text with a custom font
'*****************************************************************
    Dim TempVA As CharVA
    Dim tempstr() As String
    Dim Count As Integer
    Dim ascii() As Byte
    Dim i As Long
    Dim j As Long
    Dim yOffset As Single


    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub


    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)


    If Center Then
        X = X - CInt(Engine_GetTextWidth(cfonts(font), Text) * 0.5)
    End If

    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0

            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)

            'Loop through the characters
            For j = 1 To Len(tempstr(i))

                Call CopyMemory(TempVA, UseFont.HeaderInfo.CharVA(ascii(j - 1)), 24)    'this number represents the size of "CharVA" struct

                TempVA.X = X + Count
                TempVA.Y = Y + yOffset


                Call SpriteBatchDrawTexture(UseFont.Texture, MakeVec2(TempVA.X, TempVA.Y), MakeRect(TempVA.Tx1, TempVA.Ty1, TempVA.Tx2, TempVA.Ty2), -1)
                'Call Batch.Draw(TempVA.X, TempVA.Y, TempVA.w, TempVA.h, Color, TempVA.Tx1, TempVA.Ty1, TempVA.Tx2, TempVA.Ty2)

                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))

            Next j

        End If
    Next i

End Sub

Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
'***************************************************
    Dim i As Integer
    Dim Len_text As Long

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function

    Len_text = Len(Text)

    'Loop through the text
    For i = 1 To Len_text

        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))

    Next i

End Function

Sub Engine_Init_FontTextures()
'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
'*****************************************************************
    Dim i As Long

    '*** Default font ***
    For i = 1 To UBound(cfonts)
        cfonts(i).Texture = modDirectx.CreateTextureFromFile(App.path & "\Content\Textures\Fonts\font" & i & ".bmp")
        If cfonts(i).Texture = -1 Then
            Call MsgBox("Error en la textura de fuente utilizada " & App.path & "\Content\Textures\Fonts\font" & "Font.png", vbCritical)
            Call CloseClient
        End If
    Next
End Sub

Sub Engine_Init_FontSettings()
'*****************************************************************
'Init the custom font settings
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontSettings
'*****************************************************************
    Dim FileNum As Byte
    Dim LoopChar As Long
    Dim Row As Single
    Dim u As Single
    Dim v As Single
    Dim i As Long
    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    For i = 1 To UBound(cfonts)

        Open App.path & "\Content\Textures\Fonts\font" & i & ".dat" For Binary As #FileNum
        Get #FileNum, , cfonts(i).HeaderInfo
        Close #FileNum

        'Calculate some common values
        cfonts(i).CharHeight = cfonts(i).HeaderInfo.CellHeight - 4
        cfonts(i).RowPitch = cfonts(i).HeaderInfo.BitmapWidth \ cfonts(i).HeaderInfo.CellWidth
        cfonts(i).ColFactor = cfonts(i).HeaderInfo.CellWidth / cfonts(i).HeaderInfo.BitmapWidth
        cfonts(i).RowFactor = cfonts(i).HeaderInfo.CellHeight / cfonts(i).HeaderInfo.BitmapHeight

        'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
        For LoopChar = 0 To 255


            'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
            Row = (LoopChar - cfonts(i).HeaderInfo.BaseCharOffset) \ cfonts(i).RowPitch
            u = ((LoopChar - cfonts(i).HeaderInfo.BaseCharOffset) - (Row * cfonts(i).RowPitch)) * cfonts(i).ColFactor
            v = Row * cfonts(i).RowFactor

            'Set the verticies
            With cfonts(i).HeaderInfo.CharVA(LoopChar)
                .X = 0
                .Y = 0
                .W = cfonts(i).HeaderInfo.CellWidth
                .h = cfonts(i).HeaderInfo.CellHeight
                .Tx1 = CLng(u * cfonts(i).HeaderInfo.BitmapWidth)
                .Ty1 = CLng(v * cfonts(i).HeaderInfo.BitmapHeight)
                .Tx2 = u + cfonts(i).HeaderInfo.CellWidth
                .Ty2 = v + cfonts(i).HeaderInfo.CellHeight
            End With
        Next LoopChar
    Next i
End Sub

Public Sub DrawText(ByVal X As Integer, _
                    ByVal Y As Integer, _
                    ByVal Text As String, _
                    ByVal color As Long, _
                    Optional Center As Boolean = False, _
                    Optional font As Integer = 1)

    Call Engine_Render_Text(cfonts(font), Text, X, Y, color, Center, font)
End Sub



