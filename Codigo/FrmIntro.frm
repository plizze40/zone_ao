VERSION 5.00
Begin VB.Form FrmIntro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   355
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   840
      MouseIcon       =   "FrmIntro.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   840
      MouseIcon       =   "FrmIntro.frx":030A
      MousePointer    =   99  'Custom
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   960
      MouseIcon       =   "FrmIntro.frx":0614
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Image Image5 
      Height          =   735
      Left            =   840
      MouseIcon       =   "FrmIntro.frx":091E
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   960
      MouseIcon       =   "FrmIntro.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   3135
   End
End
Attribute VB_Name = "FrmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call PictureLoader.Initialize(App.path & "\Content\Interface\")

    Me.Picture = PictureLoader.LoadStdPicture("MenuRapido.png")    'LoadPicture(App.path & "\Graficos\MenuRapido.jpg")
End Sub
Private Sub Image2_Click()
    Call Main
End Sub

Private Sub Image6_Click()
    Unload Me
End Sub
