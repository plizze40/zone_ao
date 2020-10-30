VERSION 5.00
Begin VB.Form frmConectar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   900
   ClientTop       =   2010
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1995
      Width           =   3105
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1110
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   555
      Index           =   0
      Left            =   1605
      MouseIcon       =   "frmConnect.frx":296B5
      MousePointer    =   99  'Custom
      Top             =   4140
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   780
      Index           =   1
      Left            =   2250
      MouseIcon       =   "frmConnect.frx":299BF
      MousePointer    =   99  'Custom
      Top             =   2595
      Width           =   1710
   End
   Begin VB.Image Image1 
      Height          =   555
      Index           =   2
      Left            =   1620
      MouseIcon       =   "frmConnect.frx":29CC9
      MousePointer    =   99  'Custom
      Top             =   5835
      Width           =   2985
   End
End
Attribute VB_Name = "frmConectar"
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
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call PlayWaveDS(SND_CLICK)

        If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect

        If frmConectar.MousePointer = 11 Then
            frmConectar.MousePointer = 1
            Exit Sub
        End If


        UserName = txtUser.Text
        Dim aux As String
        aux = txtPass.Text
        UserPassword = MD5String(aux)
        If CheckUserData(False) = True Then
            frmPrincipal.Socket1.HostName = GetIPAddress
            frmPrincipal.Socket1.RemotePort = GetPortAddress

            EstadoLogin = Normal
            Me.MousePointer = 11
            frmPrincipal.Socket1.Connect
        End If
    End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Call SaveGameini
        prgRun = False
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyI And Shift = vbCtrlMask Then
        KeyCode = 0
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
    EngineRun = False

    Dim j
    For Each j In Image1()
        j.Tag = "0"
    Next

    IntervaloPaso = 0.19
    IntervaloUsar = 0.14
    Picture = PictureLoader.LoadStdPicture("conectar.png")
End Sub

Private Sub Image1_Click(Index As Integer)

    CurServer = 0

    Call PlayWaveDS(SND_CLICK)

    Select Case Index
    Case 0

        If Musica = 0 Then
            CurMidi = DirMidi & "56.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If


        EstadoLogin = dados
        frmPrincipal.Socket1.HostName = GetIPAddress
        frmPrincipal.Socket1.RemotePort = GetPortAddress
        Me.MousePointer = 11
        frmPrincipal.Socket1.Connect

    Case 1

        If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect

        If frmConectar.MousePointer = 11 Then
            frmConectar.MousePointer = 1
            Exit Sub
        End If



        UserName = txtUser.Text
        Dim aux As String
        aux = txtPass.Text
        UserPassword = MD5String(aux)
        If CheckUserData(False) = True Then
            frmPrincipal.Socket1.HostName = GetIPAddress
            frmPrincipal.Socket1.RemotePort = GetPortAddress

            EstadoLogin = Normal
            Me.MousePointer = 11
            frmPrincipal.Socket1.Connect
        End If

    Case 2
        If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect

        If frmConectar.MousePointer = 11 Then
            frmConectar.MousePointer = 1
            Exit Sub
        End If

        frmPrincipal.Socket1.HostName = GetIPAddress
        frmPrincipal.Socket1.RemotePort = GetPortAddress
        EstadoLogin = BorrarPj
        Me.MousePointer = 11
        frmPrincipal.Socket1.Connect

    End Select

End Sub

Private Sub imgGetPass_Click()

    If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect

    If frmConectar.MousePointer = 11 Then
        frmConectar.MousePointer = 1
        Exit Sub
    End If

    frmPrincipal.Socket1.HostName = GetIPAddress
    frmPrincipal.Socket1.RemotePort = GetPortAddress
    EstadoLogin = RecuperarPass
    Me.MousePointer = 11
    frmPrincipal.Socket1.Connect

End Sub
