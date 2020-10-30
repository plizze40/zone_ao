VERSION 5.00
Begin VB.Form frmPanelGM 
   Caption         =   "Panel GM"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame12 
      Caption         =   "Backup"
      Height          =   1215
      Left            =   9720
      TabIndex        =   78
      Top             =   5280
      Width           =   1455
      Begin VB.CommandButton Backup 
         Caption         =   "Backup"
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command56 
         Caption         =   "No Backup"
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command54 
      Caption         =   "Dungeon Verill"
      Height          =   615
      Left            =   9360
      TabIndex        =   73
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command53 
      Caption         =   "Dungeon Maravel"
      Height          =   615
      Left            =   9360
      TabIndex        =   72
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command52 
      Caption         =   "Templo Ancestral"
      Height          =   615
      Left            =   9360
      TabIndex        =   71
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame11 
      Caption         =   "Dungeons"
      Height          =   3735
      Left            =   9120
      TabIndex        =   70
      Top             =   960
      Width           =   1575
      Begin VB.CommandButton Command58 
         Caption         =   "Dungeon Mitral"
         Height          =   495
         Left            =   240
         TabIndex        =   77
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command57 
         Caption         =   "Dungeon Clan"
         Height          =   495
         Left            =   240
         TabIndex        =   76
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command51 
      Caption         =   "Hillidan"
      Height          =   375
      Left            =   7560
      TabIndex        =   69
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command50 
      Caption         =   "Arghal"
      Height          =   375
      Left            =   7560
      TabIndex        =   68
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command49 
      Caption         =   "Banderbill"
      Height          =   375
      Left            =   7560
      TabIndex        =   67
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command48 
      Caption         =   "Nix"
      Height          =   375
      Left            =   7560
      TabIndex        =   66
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command47 
      Caption         =   "Ullathorpe"
      Height          =   375
      Left            =   7560
      TabIndex        =   65
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame10 
      Caption         =   "Ciudades"
      Height          =   3615
      Left            =   7200
      TabIndex        =   64
      Top             =   840
      Width           =   1815
   End
   Begin VB.Frame Frame9 
      Caption         =   "Mapas!"
      Height          =   4095
      Left            =   7080
      TabIndex        =   63
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton Command46 
      Caption         =   "Hora"
      Height          =   255
      Left            =   3240
      TabIndex        =   62
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Sumonear"
      Height          =   255
      Left            =   3240
      TabIndex        =   61
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command44 
      Caption         =   "Hacerlo Neutro"
      Height          =   255
      Left            =   1800
      TabIndex        =   60
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command43 
      Caption         =   "Hacerlo PK"
      Height          =   255
      Left            =   1800
      TabIndex        =   59
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command42 
      Caption         =   "Hacerlo Ciuda"
      Height          =   255
      Left            =   1800
      TabIndex        =   58
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Unban"
      Height          =   255
      Left            =   1800
      TabIndex        =   57
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Banip"
      Height          =   255
      Left            =   1800
      TabIndex        =   56
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Oro"
      Height          =   255
      Left            =   360
      TabIndex        =   55
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Cabeza"
      Height          =   255
      Left            =   360
      TabIndex        =   54
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Ir al usuario"
      Height          =   255
      Left            =   360
      TabIndex        =   53
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Matar"
      Height          =   255
      Left            =   360
      TabIndex        =   52
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Revivir"
      Height          =   255
      Left            =   360
      TabIndex        =   51
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   320
      Left            =   2280
      TabIndex        =   50
      Text            =   "        Num o cantidad"
      Top             =   2160
      Width           =   1980
   End
   Begin VB.TextBox Text7 
      Height          =   320
      Left            =   360
      TabIndex        =   49
      Text            =   "           Nick del PJ"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Frame Frame8 
      Caption         =   "Editame, Editate"
      Height          =   2895
      Left            =   240
      TabIndex        =   48
      Top             =   1920
      Width           =   4335
      Begin VB.CommandButton Command60 
         Caption         =   "Echar Concilio"
         Height          =   255
         Left            =   3000
         TabIndex        =   86
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command59 
         Caption         =   "Echar Coalicion"
         Height          =   255
         Left            =   1560
         TabIndex        =   85
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command55 
         Caption         =   "Echar Consejo"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Coalicion 
         Caption         =   "Hacer Coalicion"
         Height          =   255
         Left            =   3000
         TabIndex        =   83
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Concilio 
         Caption         =   "Hacer Concilio"
         Height          =   255
         Left            =   3000
         TabIndex        =   82
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Consejo 
         Caption         =   "Hacer Consejo"
         Height          =   255
         Left            =   3000
         TabIndex        =   81
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Cuenta regresiva"
      Height          =   375
      Left            =   8160
      TabIndex        =   47
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   8160
      TabIndex        =   46
      Text            =   "Numero"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame7 
      Caption         =   "Cuenta"
      Height          =   1215
      Left            =   8040
      TabIndex        =   45
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Seguro"
      Height          =   375
      Left            =   6360
      TabIndex        =   44
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      Caption         =   "Mapa"
      Height          =   1215
      Left            =   6240
      TabIndex        =   43
      Top             =   5280
      Width           =   1695
      Begin VB.CommandButton Inseguro 
         Caption         =   "Inseguro"
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Crear NPC"
      Height          =   255
      Left            =   4920
      TabIndex        =   42
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4920
      TabIndex        =   41
      Text            =   "Num de Npc"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Npc"
      Height          =   1335
      Left            =   4800
      TabIndex        =   40
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Crear teleport"
      Height          =   375
      Left            =   3240
      TabIndex        =   39
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4200
      TabIndex        =   38
      Text            =   "Y"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3720
      TabIndex        =   37
      Text            =   "X"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   36
      Text            =   "Map"
      Top             =   5520
      Width           =   495
   End
   Begin VB.Frame Frame4 
      Caption         =   "Crear teleport"
      Height          =   1335
      Left            =   3000
      TabIndex        =   35
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Gms On"
      Height          =   195
      Left            =   1800
      TabIndex        =   34
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command29 
      Caption         =   "User on"
      Height          =   195
      Left            =   1800
      TabIndex        =   33
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Invisible"
      Height          =   195
      Left            =   1800
      TabIndex        =   32
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "General"
      Height          =   1335
      Left            =   1680
      TabIndex        =   31
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Crear item"
      Height          =   255
      Left            =   480
      TabIndex        =   30
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   29
      Text            =   "Num de item"
      Top             =   5640
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Items"
      Height          =   1335
      Left            =   360
      TabIndex        =   28
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "20 cupos"
      Height          =   255
      Left            =   8400
      TabIndex        =   27
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command25 
      Caption         =   "19 cupos"
      Height          =   255
      Left            =   7440
      TabIndex        =   26
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command24 
      Caption         =   "18 cupos"
      Height          =   255
      Left            =   6480
      TabIndex        =   25
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command23 
      Caption         =   "17 cupos"
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      Caption         =   "16 cupos"
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command21 
      Caption         =   "15 cupos"
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command20 
      Caption         =   "14 cupos"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "13 cupos"
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "12 cupos"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "11 cupos"
      Height          =   255
      Left            =   8640
      TabIndex        =   18
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "10 cupos"
      Height          =   255
      Left            =   7680
      TabIndex        =   17
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "9 cupos"
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "8 cupos"
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "7 cupos"
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "6 cupos"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "5 cupos"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "4 cupos"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "3 cupos"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2 cupos"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "1 cupo"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   7080
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "TORNEO"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   6840
      Width           =   9495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Activar Quest"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Gm Invisible"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Apagar servidor"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restringir servidor"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar mundo"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpieza del mundo"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "BY WON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   74
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "COMANDOS GM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPanelGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Backup_Click()
    Call SendData("/MODMAPINFO BACKUP 1")
End Sub

Private Sub Coalicion_Click()
    Call SendData("/ACEPTCONCI" & " " & Text7.Text)
End Sub

Private Sub command1_Click()
    Call SendData("/RESTRINGIR")
End Sub

Private Sub Command10_Click()
    Call SendData("/TORNEO 4")
End Sub

Private Sub Command11_Click()
    Call SendData("/TORNEO 5")
End Sub

Private Sub Command12_Click()
    Call SendData("/TORNEO 6")
End Sub

Private Sub Command13_Click()
    Call SendData("/TORNEO 7")
End Sub

Private Sub Command15_Click()
    Call SendData("/TORNEO 9")
End Sub

Private Sub Command16_Click()
    Call SendData("/TORNEO 10")
End Sub

Private Sub Command17_Click()
    Call SendData("/TORNEO 11")
End Sub

Private Sub Command18_Click()
    Call SendData("/TORNEO 12")
End Sub

Private Sub Command19_Click()
    Call SendData("/TORNEO 13")
End Sub

Private Sub Command2_Click()
    Call SendData("/LIMPIARMUNDO")
End Sub

Private Sub Command20_Click()
    Call SendData("/TORNEO 14")
End Sub

Private Sub Command21_Click()
    Call SendData("/TORNEO 15")
End Sub

Private Sub Command22_Click()
    Call SendData("/TORNEO 16")
End Sub

Private Sub Command23_Click()
    Call SendData("/TORNEO 17")
End Sub

Private Sub Command24_Click()
    Call SendData("/TORNEO 18")
End Sub

Private Sub Command25_Click()
    Call SendData("/TORNEO 19")
End Sub

Private Sub Command26_Click()
    Call SendData("/TORNEO 20")
End Sub

Private Sub Command27_Click()
    Call SendData("/ITEM" & " " & Text1.Text)
End Sub

Private Sub Command28_Click()
    Call SendData("/INVISIBLE")
End Sub

Private Sub Command29_Click()
    Call SendData("/ONLINE")
End Sub

Private Sub Command3_Click()
    Call SendData("/DOBACKUP")
End Sub

Private Sub Command30_Click()
    Call SendData("/ONLINEGM")
End Sub

Private Sub Command31_Click()
    Call SendData("/CT" & " " & Text2.Text & " " & Text3.Text & " " & Text4.Text)
End Sub

Private Sub Command32_Click()
    Call SendData("/ACC" & " " & Text5.Text)
End Sub

Private Sub Command33_Click()
    Call SendData("/MODMAPINFO PK 1")
End Sub

Private Sub Command35_Click()
    Call SendData("/REVIVIR" & " " & Text7.Text)
End Sub

Private Sub Command36_Click()
    Call SendData("/KILL" & " " & Text7.Text)
End Sub

Private Sub Command37_Click()
    Call SendData("/IRA" & " " & Text7.Text)
End Sub

Private Sub Command38_Click()
    Call SendData("/MOD" & " " & Text7.Text & " " & "head" & " " & Text8.Text)
End Sub

Private Sub Command39_Click()
    Call SendData("/MOD" & " " & Text7.Text & " " & "oro" & " " & Text8.Text)
End Sub

Private Sub Command4_Click()
    Call SendData("/APAGAR")
End Sub

Private Sub Command40_Click()
    Call SendData("/BANIP" & " " & Text7.Text)
End Sub

Private Sub Command41_Click()
    Call SendData("/UNBAN" & " " & Text7.Text)
End Sub

Private Sub Command42_Click()
    Call SendData("/MOD" & " " & Text7.Text & " " & "bando" & " " & "1")
End Sub

Private Sub Command43_Click()
    Call SendData("/MOD" & " " & Text7.Text & " " & "bando" & " " & "2")
End Sub

Private Sub Command44_Click()
    Call SendData("/MOD" & " " & Text7.Text & " " & "bando" & " " & "0")
End Sub

Private Sub Command45_Click()
    Call SendData("/SUM" & " " & Text7.Text)
End Sub

Private Sub Command47_Click()
    Call SendData("/GO 1")
End Sub

Private Sub Command48_Click()
    Call SendData("/GO 34")
End Sub

Private Sub Command49_Click()
    Call SendData("/GO 59")
End Sub

Private Sub Command5_Click()
    Call SendData("/INVISIBLE")
End Sub

Private Sub Command50_Click()
    Call SendData("/GO 98")
End Sub

Private Sub Command51_Click()
    Call SendData("/GO 149")
End Sub

Private Sub Command52_Click()
    Call SendData("/GO 181")
End Sub

Private Sub Command53_Click()
    Call SendData("/GO  76")
End Sub

Private Sub Command54_Click()
    Call SendData("/GO 139")
End Sub

Private Sub Command55_Click()
    Call SendData("/KICKCONSE" & " " & Text7.Text)
End Sub

Private Sub Command56_Click()
    Call SendData("/MODMAPINFO BACKUP 0")
End Sub

Private Sub Command57_Click()
    Call SendData("/GO 209")
End Sub

Private Sub Command58_Click()
    Call SendData("/GO 201")
End Sub

Private Sub Command59_Click()
    Call SendData("/KICKCONSECAOS" & " " & Text7.Text)
End Sub

Private Sub Command6_Click()
    Call SendData("/MODOQUEST")
End Sub

Private Sub Command60_Click()
    Call SendData("/KICKCONCI" & " " & Text7.Text)
End Sub

Private Sub Command7_Click()
    Call SendData("/TORNEO 1")
End Sub

Private Sub Cupos_Click()
    Call SendData("/TORNEO 2")
End Sub

Private Sub Command9_Click()
    Call SendData("/TORNEO 3")
End Sub

Private Sub Concilio_Click()
    Call SendData("/ACEPTCONSECAOS" & " " & Text7.Text)
End Sub

Private Sub Consejo_Click()
    Call SendData("/ACEPTCONSE" & " " & Text7.Text)
End Sub

Private Sub Inseguro_Click()
    Call SendData("/MODMAPINFO PK 0")
End Sub

