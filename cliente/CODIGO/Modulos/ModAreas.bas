Attribute VB_Name = "ModAreas"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Const TilesBuffer As Byte = 4

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public Const MargenX     As Integer = 16

Public Const MargenY     As Integer = 12

Public Sub CambioDeArea(ByVal x As Integer, ByVal y As Integer, ByVal Head As Byte)

    Dim loopX As Long, loopY As Long

    Dim MinX  As Integer

    Dim MinY  As Integer

    Dim MaxX  As Integer

    Dim MaxY  As Integer

    MinX = x
    MinY = y
    MaxX = x
    MaxY = y

    If Head = E_Heading.south Then
        MinX = MinX - MargenX
        MaxX = MaxX + MargenX
        MinY = MinY - MargenY - 1
        MaxY = MinY
    ElseIf Head = E_Heading.north Then
        MinX = MinX - MargenX
        MaxX = MaxX + MargenX
        MinY = MinY + MargenY + 1
        MaxY = MinY
        
    ElseIf Head = E_Heading.east Then
        MinX = MinX - MargenX - 1
        MaxX = MinX
        MinY = MinY - MargenY
        MaxY = MaxY + MargenY
        
    ElseIf Head = E_Heading.west Then
        MinX = MinX + MargenX + 1
        MaxX = MinX
        MinY = MinY - MargenY
        MaxY = MaxY + MargenY
    
    End If
    
    If MinY < 1 Then MinY = 1
    If MinX < 1 Then MinX = 1
    If MaxY > MapInfo.Height Then MaxY = MapInfo.Height
    If MaxX > MapInfo.Width Then MaxX = MapInfo.Width

    For loopX = MinX To MaxX
        For loopY = MinY To MaxY

            If MapData(loopX, loopY).CharIndex > 0 Then
                If MapData(loopX, loopY).CharIndex <> UserCharIndex Then
                    Call EraseChar(MapData(loopX, loopY).CharIndex)

                End If

            End If

            'Erase OBJs
            If MapData(loopX, loopY).ObjGrh.GrhIndex = GrhFogata Then
                MapData(loopX, loopY).Graphic(3).GrhIndex = 0
                Call Light_Destroy_ToMap(loopX, loopY)

            End If

            MapData(loopX, loopY).ObjGrh.GrhIndex = 0
        Next loopY
    Next loopX
    
    'Call RefreshAllChars
End Sub

Public Sub LimpiarArea()

    Dim x As Integer

    Dim y As Integer

    For x = UserPos.x - MargenX * 2 To UserPos.x + MargenX * 2
        For y = UserPos.y - MargenY * 2 To UserPos.y + MargenY * 2

            If InMapBounds(x, y) Then
                If MapData(x, y).CharIndex > 0 Then
                    If MapData(x, y).CharIndex <> UserCharIndex Then
                        Call EraseChar(MapData(x, y).CharIndex)

                    End If

                End If

                If MapData(x, y).ObjGrh.GrhIndex = GrhFogata Then
                    MapData(x, y).Graphic(3).GrhIndex = 0
                    Call Light_Destroy_ToMap(x, y)

                End If

                MapData(x, y).ObjGrh.GrhIndex = 0

            End If

        Next y
    Next x

End Sub

Public Sub LimpiarAreaTelep()

    Dim x As Integer

    Dim y As Integer

    For x = UserPos.x - MargenX * 2 To UserPos.x + MargenX * 2
        For y = UserPos.y - MargenY * 2 To UserPos.y + MargenY * 2

            If InMapBounds(x, y) Then
                If MapData(x, y).CharIndex > 0 Then
                    Call EraseChar(MapData(x, y).CharIndex)

                End If

                MapData(x, y).ObjGrh.GrhIndex = 0

            End If

        Next y
    Next x

End Sub
