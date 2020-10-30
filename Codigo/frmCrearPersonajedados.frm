VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonajedados.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   12075.47
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCorreo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   3840
      TabIndex        =   30
      Top             =   2520
      Width           =   3720
   End
   Begin VB.TextBox txtPasswdCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   7800
      PasswordChar    =   "*"
      TabIndex        =   32
      Top             =   2520
      Width           =   3720
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   7800
      PasswordChar    =   "*"
      TabIndex        =   31
      Top             =   1800
      Width           =   3720
   End
   Begin VB.TextBox txtCorreo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   3840
      TabIndex        =   29
      Top             =   1800
      Width           =   3720
   End
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":57AD4
      Left            =   6000
      List            =   "frmCrearPersonajedados.frx":57ADE
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   3600
      Width           =   2040
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":57AF1
      Left            =   3840
      List            =   "frmCrearPersonajedados.frx":57B04
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   3600
      Width           =   2040
   End
   Begin VB.ComboBox lstHogar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":57B31
      Left            =   3840
      List            =   "frmCrearPersonajedados.frx":57B38
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   4320
      Width           =   2040
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   420
      Left            =   3960
      MaxLength       =   20
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   21
      Left            =   5520
      TabIndex        =   47
      Top             =   7800
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   31
      Left            =   5760
      Top             =   7920
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   30
      Left            =   5160
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label modCarisma 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   46
      Top             =   3240
      Width           =   690
   End
   Begin VB.Label modInteligencia 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   45
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label modConstitucion 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   44
      Top             =   2280
      Width           =   690
   End
   Begin VB.Label modAgilidad 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   43
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label modfuerza 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   42
      Top             =   1440
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1320
      MouseIcon       =   "frmCrearPersonajedados.frx":57B48
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblPass2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   11520
      TabIndex        =   41
      Top             =   2400
      Width           =   345
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   180
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblMailOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7560
      TabIndex        =   37
      Top             =   1800
      Width           =   240
   End
   Begin VB.Label lblMail2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7560
      TabIndex        =   35
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label lblPassOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   11520
      TabIndex        =   33
      Top             =   1680
      Width           =   345
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   42
      Left            =   5880
      MouseIcon       =   "frmCrearPersonajedados.frx":57E52
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   43
      Left            =   5160
      MouseIcon       =   "frmCrearPersonajedados.frx":57FA4
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   195
   End
   Begin VB.Label puntosquedan 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6600
      TabIndex        =   28
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6600
      TabIndex        =   27
      Top             =   6840
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":580F6
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":58248
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":5839A
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":584EC
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":5863E
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":58790
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":588E2
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":58A34
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":58B86
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":58CD8
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   5280
      MouseIcon       =   "frmCrearPersonajedados.frx":58E2A
      MousePointer    =   99  'Custom
      Top             =   5040
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   5280
      MouseIcon       =   "frmCrearPersonajedados.frx":58F7C
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   5280
      MouseIcon       =   "frmCrearPersonajedados.frx":590CE
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   2400
      MouseIcon       =   "frmCrearPersonajedados.frx":59220
      MousePointer    =   99  'Custom
      Top             =   5040
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":59372
      MousePointer    =   99  'Custom
      Top             =   5040
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":594C4
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":59616
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":59768
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":598BA
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":59A0C
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":59B5E
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   14
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":59CB0
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   16
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":59E02
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   18
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":59F54
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   3000
      MouseIcon       =   "frmCrearPersonajedados.frx":5A0A6
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   22
      Left            =   5880
      MouseIcon       =   "frmCrearPersonajedados.frx":5A1F8
      MousePointer    =   99  'Custom
      Top             =   5040
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   24
      Left            =   5760
      MouseIcon       =   "frmCrearPersonajedados.frx":5A34A
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   26
      Left            =   5760
      MouseIcon       =   "frmCrearPersonajedados.frx":5A49C
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   5880
      MouseIcon       =   "frmCrearPersonajedados.frx":5A5EE
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   29
      Left            =   5160
      MouseIcon       =   "frmCrearPersonajedados.frx":5A740
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   5880
      MouseIcon       =   "frmCrearPersonajedados.frx":5A892
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   33
      Left            =   5160
      MouseIcon       =   "frmCrearPersonajedados.frx":5A9E4
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   5880
      MouseIcon       =   "frmCrearPersonajedados.frx":5AB36
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   5280
      MouseIcon       =   "frmCrearPersonajedados.frx":5AC88
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   36
      Left            =   5760
      MouseIcon       =   "frmCrearPersonajedados.frx":5ADDA
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   5160
      MouseIcon       =   "frmCrearPersonajedados.frx":5AF2C
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   38
      Left            =   5880
      MouseIcon       =   "frmCrearPersonajedados.frx":5B07E
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   5280
      MouseIcon       =   "frmCrearPersonajedados.frx":5B1D0
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   40
      Left            =   5760
      MouseIcon       =   "frmCrearPersonajedados.frx":5B322
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   41
      Left            =   5160
      MouseIcon       =   "frmCrearPersonajedados.frx":5B474
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   255
   End
   Begin VB.Image boton 
      Height          =   495
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmCrearPersonajedados.frx":5B5C6
      MousePointer    =   99  'Custom
      Top             =   8400
      Width           =   1125
   End
   Begin VB.Image boton 
      Appearance      =   0  'Flat
      Height          =   570
      Index           =   0
      Left            =   9720
      MouseIcon       =   "frmCrearPersonajedados.frx":5B718
      MousePointer    =   99  'Custom
      Top             =   8400
      Width           =   2280
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   20
      Left            =   5400
      TabIndex        =   26
      Top             =   7560
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Index           =   19
      Left            =   5520
      TabIndex        =   25
      Top             =   7320
      Width           =   165
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   18
      Left            =   5400
      TabIndex        =   24
      Top             =   6960
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   17
      Left            =   5400
      TabIndex        =   23
      Top             =   6720
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   16
      Left            =   5400
      TabIndex        =   22
      Top             =   6480
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   15
      Left            =   5400
      TabIndex        =   21
      Top             =   6120
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   14
      Left            =   5400
      TabIndex        =   20
      Top             =   5850
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   13
      Left            =   5400
      TabIndex        =   19
      Top             =   5565
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   12
      Left            =   5400
      TabIndex        =   18
      Top             =   5280
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   11
      Left            =   5400
      TabIndex        =   17
      Top             =   4995
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   10
      Left            =   2520
      TabIndex        =   16
      Top             =   7830
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   9
      Left            =   2520
      TabIndex        =   15
      Top             =   7560
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   8
      Left            =   2520
      TabIndex        =   14
      Top             =   7275
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   7
      Left            =   2520
      TabIndex        =   13
      Top             =   6975
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   6
      Left            =   2520
      TabIndex        =   12
      Top             =   6720
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   5
      Left            =   2520
      TabIndex        =   11
      Top             =   6420
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   4
      Left            =   2520
      TabIndex        =   10
      Top             =   6120
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   3
      Left            =   2520
      TabIndex        =   9
      Top             =   5850
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   5565
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      Top             =   4995
      Width           =   165
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   5280
      Width           =   405
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   2160
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Option Explicit

Public SkillPoints As Byte
Function CheckData() As Boolean

    If UserRaza = 0 Then
        MsgBox "Seleccione la raza del personaje."
        Exit Function
    End If

    If UserHogar = 0 Then
        MsgBox "Seleccione el hogar del personaje."
        Exit Function
    End If

    If UserSexo = -1 Then
        MsgBox "Seleccione el sexo del personaje."
        Exit Function
    End If

    If SkillPoints > 0 Then
        MsgBox "Asigne los skillpoints del personaje."
        Exit Function
    End If

    Dim i As Integer
    For i = 1 To NUMATRIBUTOS
        If UserAtributos(i) = 0 Then
            MsgBox "Los atributos del personaje son invalidos."
            Exit Function
        End If
    Next i

    CheckData = True

End Function
Private Sub boton_Click(Index As Integer)
    Dim i As Integer
    Dim k As Object

    Call PlayWaveDS(SND_CLICK)

    Select Case Index
    Case 0
        LlegoConfirmacion = False
        Confirmacion = 0

        i = 1

        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next

        UserName = txtNombre.Text

        If right$(UserName, 1) = " " Then
            UserName = Trim(UserName)
            MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If

        UserRaza = lstRaza.ListIndex + 1
        UserSexo = lstGenero.ListIndex
        UserHogar = lstHogar.ListIndex + 5

        UserAtributos(1) = 1
        UserAtributos(2) = 1
        UserAtributos(3) = 1
        UserAtributos(4) = 1
        UserAtributos(5) = 1

        If CheckData() Then
            UserPassword = MD5String(txtPasswd.Text)
            UserEmail = txtCorreo.Text

            If Not CheckMailString(UserEmail) Then
                MsgBox "Direccion de mail inválida.", vbExclamation, "ZoneAO"
                txtCorreo.SetFocus
                Exit Sub
            End If

            If UserEmail <> txtCorreo2.Text Then
                MsgBox "Las direcciones de mail no coinciden.", vbExclamation, "ZoneAO"
                txtCorreo2.Text = ""
                txtCorreo2.SetFocus
                Exit Sub
            End If

            If Len(Trim(txtPasswd)) = 0 Then
                MsgBox "Tenés que ingresar una contraseña.", vbExclamation, "ZoneAO"
                txtPasswd.SetFocus
                Exit Sub
            End If

            If Len(Trim(txtPasswd)) < 6 Then
                MsgBox "El password debe tener al menos 6 caracteres.", vbExclamation, "ZoneAO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If

            If Trim(txtPasswd) <> Trim(txtPasswdCheck) Then
                MsgBox "Las contraseñas no coinciden.", vbInformation, "ZoneAO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If

            frmPrincipal.Socket1.HostName = GetIPAddress
            frmPrincipal.Socket1.RemotePort = GetPortAddress

            Me.MousePointer = 11
            EstadoLogin = CrearNuevoPj

            If Not frmPrincipal.Socket1.Connected Then
                Call MsgBox("Error: Se ha perdido la conexion con el server.")
                Unload Me
            Else
                Call Login(ValidarLoginMSG(CInt(bRK)))
            End If

            If Musica = 0 Then
                CurMidi = DirMidi & "2.mid"
                LoopMidi = 1
                Call CargarMIDI(CurMidi)
                Call Play_Midi
            End If

            frmConectar.Picture = PictureLoader.LoadStdPicture("conectar.png")
        End If

    Case 1
        If Musica = 0 Then
            CurMidi = DirMidi & "2.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If

        frmConectar.Picture = PictureLoader.LoadStdPicture("conectar.png")

        frmPrincipal.Socket1.Disconnect
        frmConectar.MousePointer = 1
        Unload Me
    End Select

End Sub
Private Sub command1_Click(Index As Integer)
    Call PlayWaveDS(SND_CLICK)

    Dim indice
    If Index Mod 2 = 0 Then
        If SkillPoints > 0 Then
            indice = Index \ 2
            Skill(indice).Caption = Val(Skill(indice).Caption) + 1
            SkillPoints = SkillPoints - 1
        End If
    Else
        If SkillPoints < 10 Then

            indice = Index \ 2
            If Val(Skill(indice).Caption) > 0 Then
                Skill(indice).Caption = Val(Skill(indice).Caption) - 1
                SkillPoints = SkillPoints + 1
            End If
        End If
    End If

    Puntos.Caption = SkillPoints
End Sub
Private Sub Form_Load()

    SkillPoints = 10
    Puntos.Caption = SkillPoints
    Me.Picture = PictureLoader.LoadStdPicture("CrearPersonajeConDados.png")
    Me.MousePointer = vbDefault

    Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 3"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 4"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
    End Select

End Sub

Private Sub Pîcture4_Click()

End Sub

Private Sub Image1_Click()
    PlayWaveDS (SND_CLICK)
    Call SendData("TIRDAD")
End Sub

Private Sub lstRaza_click()

    Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 3"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 4"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
    End Select

End Sub
Private Sub txtCorreo_Change()

    If Not CheckMailString(txtCorreo) Then
        lblMailOK = "O"
        lblMailOK.ForeColor = &HC0&
        lblMail2OK = "O"
        lblMail2OK.ForeColor = &HC0&
        Exit Sub
    End If

    lblMailOK = "P"
    lblMailOK.ForeColor = &H80FF&

    If (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Then
        lblMail2OK = "P"
        lblMail2OK.ForeColor = &H80FF&
    Else
        lblMail2OK = "O"
        lblMail2OK.ForeColor = &HC0&
    End If

End Sub
Private Sub txtCorreo_GotFocus()

    MsgBox "Te recordamos que la dirección de correo electrónico debe ser REAL, ya que en caso contrario, no podrás recuperar tu personaje en el futuro."

End Sub
Private Sub txtCorreo2_Change()

    If Not CheckMailString(txtCorreo) Then
        lblMailOK = "O"
        lblMailOK.ForeColor = &HC0&
        lblMail2OK = "O"
        lblMail2OK.ForeColor = &HC0&
        Exit Sub
    End If

    lblMailOK = "P"
    lblMailOK.ForeColor = &H80FF&

    If (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Then
        lblMail2OK = "P"
        lblMail2OK.ForeColor = &H80FF&
    Else
        lblMail2OK = "O"
        lblMail2OK.ForeColor = &HC0&
    End If

End Sub
Private Sub txtPasswd_Change()

    If Len(Trim(txtPasswd)) < 6 Then
        lblPass2OK = "O"
        lblPass2OK.ForeColor = &HC0&
        lblPassOK = "O"
        lblPassOK.ForeColor = &HC0&
        Exit Sub
    End If

    lblPass2OK = "P"
    lblPass2OK.ForeColor = &H80FF&

    If (txtPasswdCheck = txtPasswd) Then
        lblPassOK = "P"
        lblPassOK.ForeColor = &H80FF&
    Else
        lblPassOK = "O"
        lblPassOK.ForeColor = &HC0&
    End If

End Sub
Private Sub txtPasswdCheck_Change()

    If Len(Trim(txtPasswd)) < 6 Then
        lblPass2OK = "O"
        lblPass2OK.ForeColor = &HC0&
        lblPassOK = "O"
        lblPassOK.ForeColor = &HC0&
        Exit Sub
    End If

    lblPass2OK = "P"
    lblPass2OK.ForeColor = &H80FF&

    If (txtPasswdCheck = txtPasswd) Then
        lblPassOK = "P"
        lblPassOK.ForeColor = &H80FF&
    Else
        lblPassOK = "O"
        lblPassOK.ForeColor = &HC0&
    End If

End Sub
Private Sub txtNombre_Change()
    txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
    MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub
