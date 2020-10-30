Attribute VB_Name = "mdlLeeMapas"
'F�nixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Option Explicit







Public Type TileMap
    bloqueado As Byte

    grafs1 As Integer
    grafs2 As Integer
    grafs3 As Integer
    grafs4 As Integer
    trigger As Integer

    t1 As Integer
End Type

Public Type TileInf
    dest_mapa As Integer
    dest_x As Integer
    dest_y As Integer

    Npc As Integer

    obj_ind As Integer
    obj_cant As Integer

    t1 As Integer
    t2 As Integer
End Type

Public Declare Function MAPCargaMapa Lib "LeeMapas.dll" (ByVal archmap As String, ByVal archinf As String) As Long
Public Declare Function MAPCierraMapa Lib "LeeMapas.dll" (ByVal Dm As Long) As Long

Public Declare Function MAPLeeMapa Lib "LeeMapas.dll" (ByVal Dm As Long, Tile_Map As TileMap, Tile_Inf As TileInf) As Long

