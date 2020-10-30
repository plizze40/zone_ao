VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   Caption         =   "Sistema de Canje"
   ClientHeight    =   4680
   ClientLeft      =   7830
   ClientTop       =   4215
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7035
   Begin VB.CommandButton Command1 
      Caption         =   "Canjear"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3600
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3480
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblPermisos 
      Height          =   975
      Left            =   3360
      TabIndex        =   8
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label lblStat 
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblPrecio 
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblNombre 
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clases Permitidas"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stats:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio:"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   555
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()

    If List1.Text = "Tunica de Rey (Altos)" Then Call SendData("/CANJEO T1")
    If List1.Text = "Sombrero Infernal" Then Call SendData("/CANJEO T2")
    If List1.Text = "Báculo de Mago Oscuro" Then Call SendData("/CANJEO T3")
    If List1.Text = "Túnica de la Alianza" Then Call SendData("/CANJEO T4")
    If List1.Text = "Túnica de las sombras" Then Call SendData("/CANJEO T5")
    If List1.Text = "Espada de Neithan +2" Then Call SendData("/CANJEO T6")
    If List1.Text = "Corona" Then Call SendData("/CANJEO T7")
    If List1.Text = "Espada Fantasmal" Then Call SendData("/CANJEO T8")
    If List1.Text = "Casco de Legionario" Then Call SendData("/CANJEO T9")
    If List1.Text = "Arco de las Sombras" Then Call SendData("/CANJEO T10")
    If List1.Text = "Arco de la Luz" Then Call SendData("/CANJEO T11")
    If List1.Text = "Arco largo engarzado" Then Call SendData("/CANJEO T12")
    If List1.Text = "Daga de Torneo" Then Call SendData("/CANJEO T13")
    If List1.Text = "Flecha +3" Then Call SendData("/CANJEO T14")
    If List1.Text = "Escudo Imperial +2" Then Call SendData("/CANJEO T15")
    If List1.Text = "Escudo de la Alianza" Then Call SendData("/CANJEO T16")
    If List1.Text = "Corona de Rey" Then Call SendData("/CANJEO T17")
    If List1.Text = "Daga de Hielo" Then Call SendData("/CANJEO T18")
    If List1.Text = "Escudo Dinal +1" Then Call SendData("/CANJEO T19")
    If List1.Text = "Túnica Angelical" Then Call SendData("/CANJEO T20")
    If List1.Text = "Gema del clan" Then Call SendData("/CANJEO T21")
    If List1.Text = "Sombrero de las sombras" Then Call SendData("/CANJEO T22")
    If List1.Text = "Sombrero de la alianza" Then Call SendData("/CANJEO T23")
    If List1.Text = "Armadura Thek" Then Call SendData("/CANJEO T24")
    If List1.Text = "Espada Ardiente" Then Call SendData("/CANJEO T25")
    If List1.Text = "Casco Thek" Then Call SendData("/CANJEO T26")
    If List1.Text = "Casco Oscuro" Then Call SendData("/CANJEO T27")
    If List1.Text = "Escudo Oscuro" Then Call SendData("/CANJEO T28")
    If List1.Text = "Escudo de Asesino" Then Call SendData("/CANJEO T29")
    If List1.Text = "Tunica de Rey (Bajos)" Then Call SendData("/CANJEO T30")


End Sub
Private Sub Form_Load()

    List1.AddItem "Tunica de Rey (Altos)"
    List1.AddItem "Tunica de Rey (Bajos)"
    List1.AddItem "Sombrero Infernal"
    List1.AddItem "Báculo de Mago Oscuro"
    List1.AddItem "Túnica de la Alianza"
    List1.AddItem "Túnica de las sombras"
    List1.AddItem "Espada de Neithan +2"
    List1.AddItem "Corona"
    List1.AddItem "Espada Fantasmal"
    List1.AddItem "Casco de Legionario"
    List1.AddItem "Arco de las Sombras"
    List1.AddItem "Arco de la Luz"
    List1.AddItem "Arco largo engarzado"
    List1.AddItem "Daga de Torneo"
    List1.AddItem "Flecha +3"
    List1.AddItem "Escudo Imperial +2"
    List1.AddItem "Escudo de la Alianza"
    List1.AddItem "Corona de Rey"
    List1.AddItem "Daga de Hielo"
    List1.AddItem "Escudo Dinal +1"
    List1.AddItem "Túnica Angelical"
    List1.AddItem "Gema del clan"
    List1.AddItem "Sombrero de las sombras"
    List1.AddItem "Sombrero de la alianza"
    List1.AddItem "Armadura Thek"
    List1.AddItem "Espada Ardiente"
    List1.AddItem "Casco Thek"
    List1.AddItem "Casco Oscuro"
    List1.AddItem "Escudo Oscuro"
    List1.AddItem "Escudo de Asesino"


End Sub



Private Sub list1_Click()

    If List1.Text = "Tunica de Rey (Altos)" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("685.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "24 Puntos de Canje"
        lblStat.Caption = "Min: 40 / Max: 40"
        lblPermisos.Caption = "Todas las Clases"
    End If
    If List1.Text = "Tunica de Rey (Bajos)" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16092.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "24 Puntos de Canje"
        lblStat.Caption = "Min: 40 / Max: 40"
        lblPermisos.Caption = "Todas las Clases"
    End If
    If List1.Text = "Sombrero Infernal" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16032.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "18 Puntos de Canje"
        lblStat.Caption = "Min: 15 / Max: 18"
        lblPermisos.Caption = "Mago"
    End If
    If List1.Text = "Báculo de Mago Oscuro" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16030.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "12 Puntos de Canje"
        lblStat.Caption = "Min: 0 / Max: 0"
        lblPermisos.Caption = "Mago"
    End If
    If List1.Text = "Túnica de la Alianza" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("535.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "90 Puntos de Canje"
        lblStat.Caption = "Min: 50 / Max: 50"
        lblPermisos.Caption = "Todas las Clases"
    End If
    If List1.Text = "Túnica de las sombras" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("534.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "90 Puntos de Canje"
        lblStat.Caption = "Min: 50 / Max: 50"
        lblPermisos.Caption = "Todas las Clases"
    End If
    If List1.Text = "Espada de Neithan +2" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16070.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "36 Puntos de Canje"
        lblStat.Caption = "Min: 21 / Max: 25"
        lblPermisos.Caption = "Guerrero"
    End If
    If List1.Text = "Corona" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("2023.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "48 Puntos de Canje"
        lblStat.Caption = "Min: 40 / Max: 45"
        lblPermisos.Caption = "Todas menos Guerrero"
    End If
    If List1.Text = "Espada Fantasmal" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("9630.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "48 Puntos de Canje"
        lblStat.Caption = "Min: 20 / Max: 23"
        lblPermisos.Caption = "Paladín y Guerrero"
    End If
    If List1.Text = "Casco de Legionario" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("2019.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "48 Puntos de Canje"
        lblStat.Caption = "Min: 25 / Max: 28"
        lblPermisos.Caption = "Paladín, Guerrero y Arquero"
    End If
    If List1.Text = "Arco de las Sombras" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16116.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "30 Puntos de Canje"
        lblStat.Caption = "Min: 10 / Max: 15"
        lblPermisos.Caption = "Cazador"
    End If
    If List1.Text = "Arco de la Luz" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16114.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "30 Puntos de Canje"
        lblStat.Caption = "Min: 10 / Max: 16"
        lblPermisos.Caption = "Arquero"
    End If
    If List1.Text = "Arco largo engarzado" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("1004.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "60 Puntos de Canje"
        lblStat.Caption = "Min: 14 / Max: 17"
        lblPermisos.Caption = "Arquero y Cazador"
    End If
    If List1.Text = "Daga de Torneo" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("3537.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "24 Puntos de Canje"
        lblStat.Caption = "Min: 9 / Max: 11"
        lblPermisos.Caption = "Bardo"
    End If
    If List1.Text = "Flecha +3" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("748.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "24 Puntos de Canje"
        lblStat.Caption = "Min: 0 / Max: 0"
        lblPermisos.Caption = "Arquero y Cazador"
    End If
    If List1.Text = "Escudo Imperial +2" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16058.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "24 Puntos de Canje"
        lblStat.Caption = "Min: 10 / Max: 15"
        lblPermisos.Caption = "Paladín y Guerrero"
    End If
    If List1.Text = "Escudo de la Alianza" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16068.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "24 Puntos de Canje"
        lblStat.Caption = "Min: 8 / Max: 14"
        lblPermisos.Caption = "Paladín y Guerrero"
    End If
    If List1.Text = "Corona de Rey" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16100.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "90 Puntos de Canje"
        lblStat.Caption = "Min: 50 / Max: 50"
        lblPermisos.Caption = "Todas menos Guerrero"
    End If
    If List1.Text = "Daga de Hielo" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16118.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "24 Puntos de Canje"
        lblStat.Caption = "Min: 10 / Max: 12"
        lblPermisos.Caption = "Asesino"
    End If
    If List1.Text = "Escudo Dinal +1" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16064.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "90 Puntos de Canje"
        lblStat.Caption = "Min: 10 / Max: 12"
        lblPermisos.Caption = "Bardo, Paladín y Guerrero"
    End If
    If List1.Text = "Túnica Angelical" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16112.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "24 Puntos de Canje"
        lblStat.Caption = "Min: 10 / Max: 12"
        lblPermisos.Caption = "Bardo, Paladín y Guerrero"
    End If

    If List1.Text = "Gema del clan" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("699.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "120 Puntos de Canje"
        lblStat.Caption = "Min: 0 / Max: 0"
        lblPermisos.Caption = "Todas las Clases"
    End If

    If List1.Text = "Sombrero de las sombras" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16124.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "18 Puntos de Canje"
        lblStat.Caption = "Min: 15 / Max: 18"
        lblPermisos.Caption = "Mago"
    End If

    If List1.Text = "Sombrero de la alianza" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16102.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "18 Puntos de Canje"
        lblStat.Caption = "Min: 15 / Max: 18"
        lblPermisos.Caption = "Mago"
    End If

    If List1.Text = "Armadura Thek" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("895.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "60 Puntos de Canje"
        lblStat.Caption = "Min: 50 / Max: 50"
        lblPermisos.Caption = "Paladin y Clerigo"
    End If

    If List1.Text = "Espada Ardiente" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("9629.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "48 Puntos de Canje"
        lblStat.Caption = "Min: 20 / Max: 23"
        lblPermisos.Caption = "Paladin, Clerigo y Guerrero"
    End If

    If List1.Text = "Casco Thek" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16094.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "42 Puntos de Canje"
        lblStat.Caption = "Min: 25 / Max: 28"
        lblPermisos.Caption = "Paladin Clerigo y Guerrero"
    End If

    If List1.Text = "Casco Oscuro" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16096.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "60 Puntos de Canje"
        lblStat.Caption = "Min: 30 / Max: 30"
        lblPermisos.Caption = "Guerrero y Paladin"
    End If

    If List1.Text = "Escudo Oscuro" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16066.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "54 Puntos de Canje"
        lblStat.Caption = "Min: 10 / Max: 12"
        lblPermisos.Caption = "Clérigo"
    End If

    If List1.Text = "Escudo de Asesino" Then
        Picture1.Picture = PictureLoader.LoadStdPicture("16120.png", App.path & "\Content\Textures\")
        lblNombre.Caption = List1.Text
        lblPrecio.Caption = "60 Puntos de Canje"
        lblStat.Caption = "Min: 10 / Max: 12"
        lblPermisos.Caption = "Asesino"
    End If





End Sub

