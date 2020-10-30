VERSION 5.00
Begin VB.Form frmGuildsNuevo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   600
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3240
      Width           =   6495
   End
   Begin VB.ListBox MembersList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   6495
   End
   Begin VB.ListBox GuildList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   600
      TabIndex        =   0
      Top             =   5760
      Width           =   6495
   End
   Begin VB.Image command5 
      Height          =   375
      Left            =   3000
      MouseIcon       =   "GuildNuevo.frx":0000
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Image command4 
      Height          =   255
      Left            =   3000
      MouseIcon       =   "GuildNuevo.frx":030A
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Image command8 
      Height          =   255
      Left            =   120
      MouseIcon       =   "GuildNuevo.frx":0614
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   735
   End
End
Attribute VB_Name = "frmGuildsNuevo"
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

Public Function ListaDeClanes(ByVal data As String) As Integer
    Dim a As Integer
    Dim i As Integer

    a = Val(ReadField(1, data, Asc("¬")))
    ReDim oClan(1 To a) As Clan

    For i = 1 To a
        oClan(i).name = left$(ReadField(i + 1, data, Asc("¬")), Len(ReadField(i + 1, data, Asc("¬"))) - 2)
        oClan(i).Relation = right$(ReadField(1 + i, data, Asc("¬")), 1)
    Next

    For i = 1 To a
        If oClan(i).Relation = 4 Then
            Call GuildList.AddItem(oClan(i).name)
        End If
    Next

    For i = 1 To a
        If oClan(i).Relation = 1 Then
            Call GuildList.AddItem(oClan(i).name & " (A)")
        End If
    Next

    For i = 1 To a
        If oClan(i).Relation = 2 Then
            Call GuildList.AddItem(oClan(i).name & " (E)")
        End If
    Next

    For i = 1 To a
        If oClan(i).Relation = 0 Then
            Call GuildList.AddItem(oClan(i).name)
        End If
    Next

    ListaDeClanes = a + 2

End Function
Public Sub ParseMemberInfo(ByVal data As String)

    GuildList.Clear
    MembersList.Clear
    Text1 = ""

    If Me.Visible Then Exit Sub

    Dim a As Integer
    Dim b As Integer
    Dim i As Integer

    b = ListaDeClanes(data)

    a = Val(ReadField(b, data, Asc("¬")))

    For i = 1 To a
        Call MembersList.AddItem(ReadField(b + i, data, Asc("¬")))
    Next

    b = b + a + 1

    Text1 = Replace(ReadField(b, data, Asc("¬")), "º", vbCrLf)

    Call Me.Show(vbModeless, frmPrincipal)
    Call Me.SetFocus

End Sub
Private Sub Command4_Click()

    frmCharInfo.frmmiembros = 2
    Call SendData("1HRINFO<" & MembersList.List(MembersList.ListIndex))

End Sub
Private Sub Command5_Click()
    Dim GuildName As String


    GuildName = GuildList.List(GuildList.ListIndex)
    If right$(GuildName, 1) = ")" Then GuildName = left$(GuildName, Len(GuildName) - 4)

    Call SendData("CLANDETAILS" & GuildName)

End Sub
Private Sub Command8_Click()

    Me.Visible = False
    frmPrincipal.SetFocus

End Sub
Private Sub Form_Load()

    Me.Picture = PictureLoader.LoadStdPicture("GuildMember.png")

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

