VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmSkills31.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   855
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   43
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":030A
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   42
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":0614
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   180
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   22
      Left            =   5280
      TabIndex        =   22
      Top             =   3120
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   21
      Top             =   240
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   20
      Top             =   480
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   19
      Top             =   840
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   18
      Top             =   1080
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   17
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   16
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   15
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   2400
      TabIndex        =   14
      Top             =   2280
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   13
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   12
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   11
      Top             =   3120
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   5280
      TabIndex        =   10
      Top             =   240
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   0
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":091E
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   2
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   3
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":0F32
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   4
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":123C
      MousePointer    =   99  'Custom
      Top             =   840
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   5
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":1546
      MousePointer    =   99  'Custom
      Top             =   840
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   6
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":1850
      MousePointer    =   99  'Custom
      Top             =   1080
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   7
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":1B5A
      MousePointer    =   99  'Custom
      Top             =   1080
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   8
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":1E64
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   9
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":216E
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   10
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":2478
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   11
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":2782
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   12
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":2A8C
      Top             =   1920
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   13
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":2D96
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   14
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":30A0
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   225
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   15
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":33AA
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   16
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":36B4
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   17
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":39BE
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   18
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":3CC8
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   19
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":3FD2
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   20
      Left            =   2760
      MouseIcon       =   "frmSkills31.frx":42DC
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   21
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":45E6
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   22
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":48F0
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   23
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":4BFA
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   24
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":4F04
      MousePointer    =   99  'Custom
      Top             =   600
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   25
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":520E
      MousePointer    =   99  'Custom
      Top             =   600
      Width           =   180
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   5280
      TabIndex        =   9
      Top             =   600
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   26
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":5518
      MousePointer    =   99  'Custom
      Top             =   840
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   27
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":5822
      MousePointer    =   99  'Custom
      Top             =   840
      Width           =   195
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   5280
      TabIndex        =   8
      Top             =   840
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   28
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":5B2C
      MousePointer    =   99  'Custom
      Top             =   1080
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   29
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":5E36
      MousePointer    =   99  'Custom
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   5280
      TabIndex        =   7
      Top             =   1080
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   30
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":6140
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   31
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":644A
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   180
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   16
      Left            =   5280
      TabIndex        =   6
      Top             =   1320
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   32
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":6754
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   33
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":6A5E
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   180
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   17
      Left            =   5280
      TabIndex        =   5
      Top             =   1680
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   34
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":6D68
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   35
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":7072
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   180
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   18
      Left            =   5280
      TabIndex        =   4
      Top             =   1920
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   1
      Left            =   2160
      MouseIcon       =   "frmSkills31.frx":737C
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   19
      Left            =   5280
      TabIndex        =   3
      Top             =   2280
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   36
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":7686
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":7990
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   20
      Left            =   5280
      TabIndex        =   2
      Top             =   2520
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   38
      Left            =   5640
      MouseIcon       =   "frmSkills31.frx":7C9A
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   39
      Left            =   5040
      MouseIcon       =   "frmSkills31.frx":7FA4
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   180
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   21
      Left            =   5280
      TabIndex        =   1
      Top             =   2760
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   40
      Left            =   5640
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   41
      Left            =   5040
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label puntos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Left            =   3720
      TabIndex        =   0
      Top             =   3480
      Width           =   225
   End
End
Attribute VB_Name = "frmSkills3"
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

Private Sub command1_Click(Index As Integer)

    Call PlayWaveDS(SND_CLICK)

    Dim indice
    If Index Mod 2 = 0 Then
        If Alocados > 0 Then
            indice = Index \ 2 + 1
            If indice > NUMSKILLS Then indice = NUMSKILLS
            If UserSkills(indice) < MAXSKILLPOINTS And Val(Text1(indice).Caption) < 100 Then
                Text1(indice).Caption = Val(Text1(indice).Caption) + 1
                flags(indice) = flags(indice) + 1
                Alocados = Alocados - 1
            End If

        End If
    Else
        If Alocados < SkillPoints Then

            indice = Index \ 2 + 1
            If Val(Text1(indice).Caption) > 0 And flags(indice) > 0 Then
                Text1(indice).Caption = Val(Text1(indice).Caption) - 1
                flags(indice) = flags(indice) - 1
                Alocados = Alocados + 1
            End If
        End If
    End If

    Puntos.Caption = Alocados
End Sub

Private Sub Form_Deactivate()

    Me.Visible = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving = False And Button = vbLeftButton Then
        DX = X
        dy = Y
        bmoving = True
    End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving And ((X <> DX) Or (Y <> dy)) Then Move left + (X - DX), top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then bmoving = False

End Sub

Private Sub Form_Load()

    Me.Picture = PictureLoader.LoadStdPicture("AgregarPuntosSkills.png")




    Dim i As Integer

    ReDim flags(1 To NUMSKILLS)






End Sub

Private Sub Image1_Click()

    Dim i As Integer
    Dim cad As String
    For i = 1 To NUMSKILLS
        cad = cad & flags(i) & ","
    Next
    SendData "SKSE" & cad
    If Alocados = 0 Then frmPrincipal.Label1.Visible = False
    SkillPoints = Alocados
    Unload Me
End Sub
