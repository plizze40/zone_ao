VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Begin VB.Form frmPrincipal 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9015
   ClientLeft      =   6525
   ClientTop       =   2235
   ClientWidth     =   11985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmMain.frx":1CCA
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   120
      Top             =   2760
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   "FlamiusAO"
      HostName        =   "FlamiusAO"
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   10200
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   10200
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox MainView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6240
      Left            =   120
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   89
      Top             =   2160
      Width           =   8160
   End
   Begin Captura.wndCaptura Captura1 
      Left            =   120
      Top             =   2280
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   10320
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   87
      Top             =   9000
      Width           =   975
   End
   Begin VB.Timer AntiCheat 
      Interval        =   1000
      Left            =   2040
      Top             =   2760
   End
   Begin RichTextLib.RichTextBox rectxt 
      Height          =   1545
      Left            =   120
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2725
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":25C89
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frInvent 
      BorderStyle     =   0  'None
      Height          =   4245
      Left            =   8640
      TabIndex        =   15
      Top             =   1920
      Width           =   3090
      Begin VB.Image Image5 
         Height          =   435
         Index           =   3
         Left            =   1440
         MouseIcon       =   "frmMain.frx":25D08
         MousePointer    =   99  'Custom
         Top             =   3840
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   435
         Index           =   2
         Left            =   1440
         MouseIcon       =   "frmMain.frx":26012
         MousePointer    =   99  'Custom
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   1
         Left            =   1680
         MouseIcon       =   "frmMain.frx":2631C
         MousePointer    =   99  'Custom
         Top             =   3600
         Width           =   435
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   0
         Left            =   1080
         MouseIcon       =   "frmMain.frx":26626
         MousePointer    =   99  'Custom
         Top             =   3600
         Width           =   435
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   480
         Left            =   3240
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   51
         Top             =   840
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   1320
         TabIndex        =   38
         Top             =   600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Height          =   480
         Index           =   3
         Left            =   1320
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   25
         Left            =   2640
         TabIndex        =   73
         Top             =   2760
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   24
         Left            =   2160
         TabIndex        =   72
         Top             =   2760
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   23
         Left            =   1680
         TabIndex        =   71
         Top             =   2760
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   22
         Left            =   1200
         TabIndex        =   70
         Top             =   2760
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   21
         Left            =   720
         TabIndex        =   69
         Top             =   2760
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   16
         Left            =   720
         TabIndex        =   68
         Top             =   2280
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   17
         Left            =   1200
         TabIndex        =   67
         Top             =   2280
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   18
         Left            =   1680
         TabIndex        =   66
         Top             =   2280
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   19
         Left            =   2160
         TabIndex        =   65
         Top             =   2280
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   20
         Left            =   2640
         TabIndex        =   64
         Top             =   2280
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   15
         Left            =   2640
         TabIndex        =   63
         Top             =   1800
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   14
         Left            =   2160
         TabIndex        =   62
         Top             =   1800
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   13
         Left            =   1680
         TabIndex        =   61
         Top             =   1800
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   12
         Left            =   1200
         TabIndex        =   60
         Top             =   1800
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   11
         Left            =   720
         TabIndex        =   59
         Top             =   1800
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   10
         Left            =   2640
         TabIndex        =   58
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   9
         Left            =   2160
         TabIndex        =   57
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   8
         Left            =   1680
         TabIndex        =   56
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   1200
         TabIndex        =   55
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   54
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   53
         Top             =   840
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   50
         Top             =   840
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   49
         Top             =   840
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   40
         Top             =   600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblHechizos 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1440
         MouseIcon       =   "frmMain.frx":26930
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   0
         Width           =   1560
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   1320
         TabIndex        =   33
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   840
         TabIndex        =   39
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   1800
         TabIndex        =   37
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   36
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   360
         TabIndex        =   35
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   840
         TabIndex        =   34
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   11
         Left            =   360
         TabIndex        =   30
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   12
         Left            =   840
         TabIndex        =   29
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   13
         Left            =   1320
         TabIndex        =   28
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   14
         Left            =   1800
         TabIndex        =   27
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   15
         Left            =   2280
         TabIndex        =   26
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   16
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   17
         Left            =   840
         TabIndex        =   24
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   18
         Left            =   1320
         TabIndex        =   23
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   19
         Left            =   1800
         TabIndex        =   22
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   20
         Left            =   2280
         TabIndex        =   21
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   21
         Left            =   360
         TabIndex        =   20
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   22
         Left            =   840
         TabIndex        =   19
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   23
         Left            =   1320
         TabIndex        =   18
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   24
         Left            =   1800
         TabIndex        =   17
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   25
         Left            =   2280
         TabIndex        =   16
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   1800
         TabIndex        =   32
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   10
         Left            =   2280
         TabIndex        =   31
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   2
         Left            =   840
         Stretch         =   -1  'True
         Top             =   600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   4
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   5
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   6
         Left            =   360
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   7
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   8
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   9
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   10
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   11
         Left            =   360
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   12
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   13
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   14
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   15
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   16
         Left            =   360
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   17
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   18
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   19
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   20
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   21
         Left            =   360
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   22
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   23
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   24
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   25
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   480
      End
      Begin VB.Image imgFondoInvent 
         Height          =   4635
         Left            =   0
         Top             =   0
         Width           =   3240
      End
   End
   Begin VB.Timer tmrBmp 
      Left            =   1560
      Top             =   2280
   End
   Begin VB.Timer trabajo 
      Enabled         =   0   'False
      Left            =   600
      Top             =   2760
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   1080
      Top             =   2760
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   600
      Top             =   2280
   End
   Begin VB.Timer FPS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   2280
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2040
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.Timer Attack 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   2760
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1800
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Frame frHechizos 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   4275
      Left            =   8640
      TabIndex        =   42
      Top             =   1920
      Width           =   3105
      Begin VB.ListBox lstHechizos 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2565
         Left            =   240
         TabIndex        =   43
         Top             =   960
         Width           =   2715
      End
      Begin VB.Label lblInvent 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         MouseIcon       =   "frmMain.frx":26C3A
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   0
         Width           =   1650
      End
      Begin VB.Label lblLanzar 
         BackStyle       =   0  'Transparent
         Height          =   720
         Left            =   120
         MouseIcon       =   "frmMain.frx":26F44
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   3600
         Width           =   2025
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Height          =   480
         Left            =   1800
         MouseIcon       =   "frmMain.frx":2724E
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Label lblAbajo 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         MouseIcon       =   "frmMain.frx":27558
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   600
         Width           =   300
      End
      Begin VB.Label lblArriba 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         MouseIcon       =   "frmMain.frx":27862
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   600
         Width           =   300
      End
      Begin VB.Image imgFondoHechizos 
         Height          =   4395
         Left            =   0
         Picture         =   "frmMain.frx":27B6C
         Top             =   0
         Width           =   3240
      End
   End
   Begin VB.Label lblcanjes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   9720
      TabIndex        =   88
      Top             =   8040
      Width           =   735
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   8280
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   8400
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   10920
      TabIndex        =   86
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(100%)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   8550
      TabIndex        =   85
      Top             =   885
      Width           =   3300
   End
   Begin VB.Label barrita 
      BackColor       =   &H00000040&
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8580
      TabIndex        =   84
      Top             =   900
      Width           =   3255
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9030
      TabIndex        =   83
      Top             =   615
      Width           =   375
   End
   Begin VB.Image Image10 
      Height          =   285
      Left            =   8520
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   3
      Left            =   8400
      MouseIcon       =   "frmMain.frx":2C1AF
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   3285
   End
   Begin VB.Image Party 
      Height          =   285
      Left            =   14760
      MouseIcon       =   "frmMain.frx":2C4B9
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   1170
   End
   Begin VB.Label NumOnline 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4320
      TabIndex        =   82
      Top             =   8595
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10155
      TabIndex        =   81
      Top             =   1305
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10680
      TabIndex        =   80
      Top             =   1320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6840
      TabIndex        =   79
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9690
      TabIndex        =   78
      Top             =   1320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label modo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1 Normal"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   77
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Agilidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6480
      TabIndex        =   76
      Top             =   8565
      Width           =   225
   End
   Begin VB.Label Fuerza 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   7425
      TabIndex        =   75
      Top             =   8565
      Width           =   300
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   960
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label casco 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3405
      TabIndex        =   1
      Top             =   8595
      Width           =   540
   End
   Begin VB.Label armadura 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   8595
      Width           =   540
   End
   Begin VB.Label escudo 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2385
      TabIndex        =   13
      Top             =   8595
      Width           =   540
   End
   Begin VB.Label arma 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1380
      TabIndex        =   12
      Top             =   8595
      Width           =   540
   End
   Begin VB.Label mapa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ullathorpe"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   8520
      Width           =   3015
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   5040
      Top             =   8760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label cantidadhp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11025
      TabIndex        =   9
      Top             =   6645
      Width           =   75
   End
   Begin VB.Label cantidadagua 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   11280
      TabIndex        =   8
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label cantidadsta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   135
      Left            =   11400
      TabIndex        =   10
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label cantidadhambre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10680
      TabIndex        =   7
      Top             =   8520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label cantidadmana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8730
      TabIndex        =   6
      Top             =   6645
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   11280
      MouseIcon       =   "frmMain.frx":2C7C3
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00003E25&
      X1              =   16
      X2              =   551.467
      Y1              =   126.333
      Y2              =   126.333
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   11640
      MouseIcon       =   "frmMain.frx":2CACD
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   495
   End
   Begin VB.Label fpstext 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   8760
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DarkTester"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9600
      TabIndex        =   4
      Top             =   480
      Width           =   1185
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   120
      Left            =   8280
      Top             =   8400
      Width           =   1245
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   1485
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   5160
      TabIndex        =   3
      Top             =   8760
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Shape Hpshp 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   10320
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   1485
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   120
      Left            =   9000
      Top             =   8400
      Width           =   525
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   120
      Left            =   11520
      Top             =   8040
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   0
      Left            =   10560
      MouseIcon       =   "frmMain.frx":2CDD7
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   1
      Left            =   10560
      MouseIcon       =   "frmMain.frx":2D0E1
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   2
      Left            =   10560
      MouseIcon       =   "frmMain.frx":2D3EB
      MousePointer    =   99  'Custom
      Top             =   8040
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   11640
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const AC_SRC_OVER = &H0
Dim blendlong As Long
Dim Contador As Integer

Public ActualSecond As Long
Public LastSecond As Long
Public tx As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long

Dim gFileName As String
Public IsPlaying As Byte
Public boton As Integer
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal length As Long)


Private Sub Form_Activate()

    If frmParty.Visible Then frmParty.SetFocus
    If frmParty2.Visible Then frmParty2.SetFocus

End Sub




Private Sub Image10_Click()
    frmCanjes.Show
End Sub

Private Sub Image5_Click(Index As Integer)

    If (ItemElegido <= 0 Or ItemElegido > MAX_INVENTORY_SLOTS) Then Exit Sub
    If ItemElegido = 1 And Index = 0 Then Exit Sub
    If ItemElegido = MAX_INVENTORY_SLOTS And Index = 1 Then Exit Sub
    If ItemElegido < 6 And Index = 2 Then Exit Sub
    If ItemElegido > MAX_INVENTORY_SLOTS - 5 And Index = 3 Then Exit Sub

    Call SendData("ZI" & ItemElegido & "," & Index)

    Select Case Index
    Case 0
        Shape1.top = imgObjeto(ItemElegido - 1).top
        Shape1.left = imgObjeto(ItemElegido - 1).left
        ItemElegido = ItemElegido - 1
    Case 1
        Shape1.top = imgObjeto(ItemElegido + 1).top
        Shape1.left = imgObjeto(ItemElegido + 1).left
        ItemElegido = ItemElegido + 1
    Case 2
        Shape1.top = imgObjeto(ItemElegido - 5).top
        Shape1.left = imgObjeto(ItemElegido - 5).left
        ItemElegido = ItemElegido - 5
    Case 3
        Shape1.top = imgObjeto(ItemElegido + 5).top
        Shape1.left = imgObjeto(ItemElegido + 5).left
        ItemElegido = ItemElegido + 5
    End Select

End Sub



Private Sub Label2_Click(Index As Integer)

    If ItemElegido <> Index And UserInventory(Index).name <> "Nada" Then
        Shape1.Visible = True
        Shape1.top = imgObjeto(Index).top
        Shape1.left = imgObjeto(Index).left
        ItemElegido = Index
    End If

End Sub

Private Sub Label3_Click()

    Call SendData("#N")

End Sub

Private Sub Label5_Click()

    Call SendData("#!")

End Sub

Private Sub Label7_Click()

    Call SendData("#O")

End Sub

Private Sub lblarriba_Click()

    If lstHechizos.ListIndex < 1 Then Exit Sub

    If lstHechizos.ListIndex >= 1 Then Call SendData("DESPHE" & 1 & "," & lstHechizos.ListIndex + 1)
    lstHechizos.ListIndex = lstHechizos.ListIndex - 1

End Sub
Private Sub lblabajo_Click()

    If lstHechizos.ListIndex > 11 Then Exit Sub    ' 2 NÚMEROS MENOS DE LINEAS PARA QUE NO BAJE MAS DE LO DICHO

    If lstHechizos.ListIndex <= 11 Then Call SendData("DESPHE" & 2 & "," & lstHechizos.ListIndex + 1)    ' 2 NÚMEROS MENOS DE LINEAS PARA QUE NO BAJE MAS DE LO DICHO
    lstHechizos.ListIndex = lstHechizos.ListIndex + 1

End Sub
Private Sub FX_Timer()
    Dim n As Byte

    If Fx = 0 And RandomNumber(1, 150) < 12 Then
        n = RandomNumber(1, 45)
        Select Case n
        Case Is <= 15
            Call PlayWaveDS("22.wav")
        Case Is <= 30
            Call PlayWaveDS("21.wav")
        Case Is <= 35
            Call PlayWaveDS("28.wav")
        Case Is <= 40
            Call PlayWaveDS("29.wav")
        Case Is <= 45
            Call PlayWaveDS("34.wav")
        End Select
    End If

End Sub
Private Sub imgObjeto_Click(Index As Integer)

    If ItemElegido <> Index And UserInventory(Index).name <> "Nada" Then
        Shape1.Visible = True
        Shape1.top = imgObjeto(Index).top
        Shape1.left = imgObjeto(Index).left
        ItemElegido = Index
    End If

End Sub
Private Sub imgObjeto_DblClick(Index As Integer)

    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

    If ItemElegido = Index Then Call SendData("USE" & ItemElegido)

End Sub
Private Sub lblHechizos_Click()

    Call PlayWaveDS(SND_CLICK)
    frHechizos.Visible = True
    frInvent.Visible = False

End Sub
Private Sub lblInvent_Click()

    Call PlayWaveDS(SND_CLICK)
    frInvent.Visible = True
    frHechizos.Visible = False

End Sub
Private Sub lblObjCant_Click(Index As Integer)

    If ItemElegido <> Index And UserInventory(Index).name <> "Nada" Then
        Shape1.Visible = True
        Shape1.top = imgObjeto(Index).top
        Shape1.left = imgObjeto(Index).left
        ItemElegido = Index
    End If

End Sub
Private Sub lblObjCant_DblClick(Index As Integer)

    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

    If ItemElegido = Index Then Call SendData("USE" & ItemElegido)

End Sub

Public Sub Play(ByVal Nombre As String, Optional ByVal LoopSound As Boolean = False)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun Then
        prgRun = False
        Cancel = 1
    End If

End Sub
Private Sub FPS_Timer()

    If logged And Not frmPrincipal.Visible Then
        Unload frmConectar
        frmPrincipal.Show
    End If

End Sub
Private Sub Image2_Click()

    Me.WindowState = vbMinimized

End Sub
Private Sub Image4_Click()

    ItemElegido = FLAGORO
    If UserGLD > 1 Then frmCantidad.Show

End Sub

Private Sub LvlLbl_Click()

    If UserPasarNivel > 0 Then
        frmPrincipal.LvlLbl.Caption = "Exp: " & PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
    Else
        frmPrincipal.LvlLbl.Caption = "¡Nivel máximo!"
    End If

End Sub

Private Sub MainView_Click()
    If Cartel Then Cartel = False

    If Comerciando = 0 Then
        Call ConvertCPtoTP(MouseX, MouseY, tx, tY)
        If Abs(UserPos.Y - tY) > 6 Then Exit Sub
        If Abs(UserPos.X - tx) > 8 Then Exit Sub
        If EligiendoWhispereo Then
            Call SendData("WH" & tx & "," & tY)
            EligiendoWhispereo = False
            Exit Sub
        End If

        If UsingSkill = 0 Then
            SendData "LC" & tx & "," & tY
        Else
            frmPrincipal.MousePointer = vbDefault
            If UsingSkill = Magia Then
                If (TiempoTranscurrido(LastHechizo) < IntervaloSpell Or TiempoTranscurrido(Hechi) < IntervaloSpell / 4) Then
                    Exit Sub
                Else: Hechi = Timer
                End If
            ElseIf UsingSkill = Proyectiles Then
                If (TiempoTranscurrido(LastFlecha) < IntervaloFlecha Or TiempoTranscurrido(Flecho) < IntervaloFlecha / 4) Then
                    Exit Sub
                Else: Flecho = Timer
                End If
            End If
            Call SendData("WLC" & tx & "," & tY & "," & UsingSkill)
            UsingSkill = 0
        End If
    End If

    If boton = vbRightButton Then Call SendData("/TELEPLOC")
    boton = 0

End Sub

Private Sub MainView_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tx & "," & tY
    End If
End Sub

Private Sub MainView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    boton = Button
End Sub

Private Sub MainView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y

    LvlLbl.Visible = True
    exp.Visible = False

End Sub

Private Sub Party_Click()

    frmParty.ListaIntegrantes.Clear
    LlegoParty = False
    Call SendData("PARINF")
    Do While Not LlegoParty
        DoEvents
    Loop
    frmParty.Visible = True
    frmParty.SetFocus
    LlegoParty = False

End Sub

Private Sub RecTxt_GotFocus()

    SendTxt.Visible = False
    frmPrincipal.SetFocus

End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Call ProcesaEntradaCmd(stxtbuffer)
        stxtbuffer = ""
        frmPrincipal.SendTxt.Text = ""
        frmPrincipal.SendTxt.Visible = False
        KeyCode = 0
    End If

End Sub

Private Sub Second_Timer()
    ActualSecond = mid$(time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = LastSecond Then End
    LastSecond = ActualSecond
End Sub





Private Sub TirarItem()
    If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            SendData "TI" & ItemElegido & "," & 1
        Else
            If UserInventory(ItemElegido).Amount > 1 Then
                frmCantidad.Show
            End If
        End If
    End If


End Sub

Private Sub AgarrarItem()
    SendData "AG"

End Sub

Private Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then
        SendData "USA" & ItemElegido
    End If

End Sub
Public Sub EquiparItem()

    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then _
       SendData "EQUI" & ItemElegido

End Sub





Private Sub lblLanzar_Click()

    If lstHechizos.List(lstHechizos.ListIndex) <> "Nada" And TiempoTranscurrido(LastHechizo) >= IntervaloSpell And TiempoTranscurrido(Hechi) >= IntervaloSpell / 4 Then
        Call SendData("LH" & lstHechizos.ListIndex + 1)
        Call SendData("UK" & Magia)
    End If

End Sub
Private Sub lblInfo_Click()
    Call SendData("INFS" & lstHechizos.ListIndex + 1)
End Sub



Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving = False And Button = vbLeftButton And Desplazar = True Then
        DX = X
        dy = Y
        bmoving = True
    End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Not SendTxt.Visible) Then

        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

            Select Case KeyCode
            Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                If Not IsPlayingCheck Then
                    Musica = 0
                    Play_Midi
                    frmOpciones.PictureMusica.Picture = PictureLoader.LoadStdPicture("tick1.png")
                Else
                    Musica = 1
                    frmOpciones.PictureMusica.Picture = PictureLoader.LoadStdPicture("tick2.png")
                    Stop_Midi
                End If    'X

            Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                Call AgarrarItem    'X

            Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                Call EquiparItem    'X

            Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                Nombres = Not Nombres    'X

            Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                Call SendData("UK" & Domar)    'X

            Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                Call SendData("UK" & Robar)    'X

            Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                Call SendData("UK" & Ocultarse)    'X

            Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                Call TirarItem    'X

            Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                If Not NoPuedeUsar Then
                    NoPuedeUsar = True
                    Call UsarItem
                End If    'X

            Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                Call SendData("RPU")
                Beep

                '..........................ShaFTeR..........................
            Case CustomKeys.BindedKey(eKeyType.mKeyNormal)
                frmPrincipal.modo = "1 Normal"
                If EligiendoWhispereo Then
                    EligiendoWhispereo = False
                    MousePointer = 1
                End If

            Case CustomKeys.BindedKey(eKeyType.mKeySusurrar)
                Call AddtoRichTextBox(frmPrincipal.rectxt, "Has click sobre el usuario al que quieres susurrar.", 255, 255, 255, 1, 0)
                frmPrincipal.modo = "2 Susurrar"
                MousePointer = 2
                EligiendoWhispereo = True

            Case CustomKeys.BindedKey(eKeyType.mKeyClan)
                frmPrincipal.modo = "3 Clan"
                If EligiendoWhispereo Then
                    EligiendoWhispereo = False
                    MousePointer = 1
                End If

            Case CustomKeys.BindedKey(eKeyType.mKeyGrito)
                frmPrincipal.modo = "4 Grito"
                If EligiendoWhispereo Then
                    EligiendoWhispereo = False
                    MousePointer = 1
                End If

            Case CustomKeys.BindedKey(eKeyType.mKeyRol)
                frmPrincipal.modo = "5 Rol"
                If EligiendoWhispereo Then
                    EligiendoWhispereo = False
                    MousePointer = 1
                End If

            Case CustomKeys.BindedKey(eKeyType.mKeyParti)
                frmPrincipal.modo = "6 Party"
                If EligiendoWhispereo Then
                    EligiendoWhispereo = False
                    MousePointer = 1
                End If


                '..........................ShaFTeR..........................

                '          Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
            Case CustomKeys.BindedKey(eKeyType.mKeyParty)
                frmParty.ListaIntegrantes.Clear
                LlegoParty = False
                Call SendData("PARINF")
                Do While Not LlegoParty
                    DoEvents
                Loop
                frmParty.Visible = True
                frmParty.SetFocus
                LlegoParty = False

            End Select
        Else

        End If
    End If

    Select Case KeyCode

    Case vbKeyF1:
        Call SendData("/SUBIR")

    Case vbKeyF2:
        Call SendData("/ORO")

    Case vbKeyF8:
        frmRecanje.Show

    Case vbKeyF3:
        Call SendData("/INVISIBLE")
        '   Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
    Case CustomKeys.BindedKey(eKeyType.mKeyInvi)
        Call SendData("/INVISIBLE")

        '   Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
    Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
        Dim i As Integer
        Captura1.Area = Ventana
        Captura1.Captura
        For i = 1 To 1000
            If Not FileExist(App.path & "\screenshots\Imagen" & i & ".bmp", vbNormal) Then Exit For
        Next
        Call SavePicture(Captura1.Imagen, App.path & "/screenshots/Imagen" & i & ".bmp")
        Call AddtoRichTextBox(frmPrincipal.rectxt, "Una imagen fue guardada en la carpeta de screenshots bajo el nombre de Imagen" & i & ".bmp", 255, 150, 50, False, False, False)


    Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
        Call frmOpciones.Show(vbModeless, frmPrincipal)

    Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
        Call SendData("/MEDITAR")    'X

        '   Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)


    Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
        Call SendData("/SALIR")    'X

    Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
        If (TiempoTranscurrido(LastGolpe) >= IntervaloGolpe) And (TiempoTranscurrido(Golpeo) >= IntervaloGolpe / 4) And (Not UserDescansar) And _
           (Not UserMeditar) Then
            Call SendData("AT")
            Golpeo = Timer
        End If    'X

    Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
        If Not frmCantidad.Visible Then
            SendTxt.Visible = True
            SendTxt.SetFocus
        End If    'X

        'Standelf
    Case CustomKeys.BindedKey(eKeyType.mKeyUnlock)
        Call SendData("(A")    'X
    End Select
End Sub

Sub Form_Load()
'BETA
    IPdelServidor = "lovely.zonagame.com.ar"
    PuertoDelServidor = 7676

    FPSFLAG = True

    Me.Picture = PictureLoader.LoadStdPicture("Principal.png")

    frmPrincipal.imgFondoInvent.Picture = PictureLoader.LoadStdPicture("Centronuevoinventario.png")
    frmPrincipal.imgFondoHechizos.Picture = PictureLoader.LoadStdPicture("Centronuevohechizos.png")

End Sub
Private Sub lstHechizos_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub
Private Sub lstHechizos_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub
Private Sub lstHechizos_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub
Private Sub Image1_Click(Index As Integer)
    Call PlayWaveDS(SND_CLICK)

    Select Case Index
    Case 0
        Call frmOpciones.Show(vbModeless, frmPrincipal)
    Case 1
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
        SendData "ATRI"
        SendData "ESKI"
        SendData "FAMA"
        Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama Or Not LlegoMinist
            DoEvents
        Loop
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
    Case 2
        If frmGuildLeader.Visible Then frmGuildLeader.Visible = False
        If frmGuildsNuevo.Visible Then frmGuildsNuevo.Visible = False
        If frmGuildAdm.Visible Then frmGuildAdm.Visible = False
        Call SendData("GLINFO")
    Case 3
        frmMapa.Visible = True
    End Select

End Sub

Private Sub Image3_Click()
    frmSalir.Show


End Sub

Private Sub Label1_Click()
    LlegaronSkills = False
    SendData "ESKI"

    Do While Not LlegaronSkills
        DoEvents
    Loop

    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = SkillPoints
    frmSkills3.Show
End Sub
Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim mx As Integer
    Dim my As Integer
    Dim aux As Integer
    mx = X \ 32 + 1
    my = Y \ 32 + 1
    aux = (mx + (my - 1) * 5) + OffsetDelInv

End Sub
Private Sub RecTxt_Change()
    On Error Resume Next

    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (Not frmComerciar.Visible) And _
           (Not frmSkills3.Visible) And _
           (Not frmMSG.Visible) And _
           (Not frmForo.Visible) And _
           (Not frmEstadisticas.Visible) And _
           (Not frmCantidad.Visible) Then
        ' Picture1.SetFocus
    End If

End Sub
Private Sub SendTxt_Change()

    stxtbuffer = SendTxt.Text

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

End Sub






Private Sub Socket1_Connect()

    Second.Enabled = True

    If EstadoLogin = CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = dados Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = RecuperarPass Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Activar Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = BorrarPj Then
        Call SendData("gIvEmEvAlcOde")
    End If
End Sub


Private Sub Socket1_Disconnect()
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False

    Socket1.Cleanup

    frmConectar.MousePointer = vbNormal
    frmCrearPersonaje.Visible = False
    frmConectar.Visible = True

    frmPrincipal.Visible = False

    Pausa = False
    UserMeditar = False

    UserSexo = 0
    UserRaza = 0
    UserEmail = ""
    bO = 100

    Dim i As Integer
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.RemoveAllDialogs
End Sub
Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)

    Select Case ErrorCode
    Case 24036
        Call MsgBox("Calmate flaco, todavía se está completando la conexión.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub

    Case 24038, 24061
        Call MsgBox("El servidor está cerrado.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

    Case 24053
        Call MsgBox("Se perdió la conexión.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

    Case 24060
        Call MsgBox("Se terminó el tiempo de espera.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

    Case Else
        Call MsgBox(ErrorString, vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

    End Select

    frmConectar.MousePointer = 1
    Response = 0
    LastSecond = 0
    Second.Enabled = False

    frmPrincipal.Socket1.Disconnect

    If Not frmCrearPersonaje.Visible Then
        frmConectar.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If

End Sub
Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String

    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer

    Call Socket1.Read(RD, DataLength)

    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    sChar = 1

    For loopc = 1 To Len(RD)
        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            rBuffer(CR) = mid$(RD, sChar, loopc - sChar)
            sChar = loopc + 1
        End If

    Next loopc

    If Len(RD) - (sChar - 1) <> 0 Then TempString = mid$(RD, sChar, Len(RD))

    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc

End Sub
