Attribute VB_Name = "modBarcos"
Option Explicit

Public Const NUM_PUERTOS       As Byte = 6

Public Const PUERTO_NIX        As Byte = 1

Public Const PUERTO_ULLATHORPE As Byte = 2

Public Const PUERTO_BANDER     As Byte = 3

Public Const PUERTO_ARGHAL     As Byte = 4

Public Const PUERTO_LINDOS     As Byte = 5

Public Const PUERTO_ARKHEIN    As Byte = 6

Public Type tPuerto

    Paso(0 To 1) As Byte
    nombre As String

End Type

Public Puertos(1 To NUM_PUERTOS) As tPuerto

Public Barco(0 To 1)             As clsBarco

Public RutaBarco(0 To 1)         As String

Public Sub InitBarcos()

    Puertos(PUERTO_NIX).nombre = "Nix"
    Puertos(PUERTO_NIX).Paso(0) = 0
    Puertos(PUERTO_NIX).Paso(1) = 24

    Puertos(PUERTO_ULLATHORPE).nombre = "Ullathorpe"
    Puertos(PUERTO_ULLATHORPE).Paso(0) = 3
    Puertos(PUERTO_ULLATHORPE).Paso(1) = 20

    Puertos(PUERTO_BANDER).nombre = "Banderbill"
    Puertos(PUERTO_BANDER).Paso(0) = 7
    Puertos(PUERTO_BANDER).Paso(1) = 16

    Puertos(PUERTO_ARGHAL).nombre = "Arghal"
    Puertos(PUERTO_ARGHAL).Paso(0) = 14
    Puertos(PUERTO_ARGHAL).Paso(1) = 9

    Puertos(PUERTO_LINDOS).nombre = "Lindos"
    Puertos(PUERTO_LINDOS).Paso(0) = 17
    Puertos(PUERTO_LINDOS).Paso(1) = 5

    Puertos(PUERTO_ARKHEIN).nombre = "Arkhein"
    Puertos(PUERTO_ARKHEIN).Paso(0) = 23
    Puertos(PUERTO_ARKHEIN).Paso(1) = 0

    RutaBarco(0) = "161,1248;36,1248;36,897;171,897;36,897;36,22;302,22;302,55;566,55;566,65;635,65;635,54;800,54;800,307;801,307;870,307;870,999;887,999;870,999;870,1224;644,1224;644,1361;645,1361;645,1371;645,1385;644,1385;644,1472;196,1472;196,1266;169,1266;169,1248"
    RutaBarco(1) = "642,1384;649,1384;649,1228;875,1228;875,995;887,995;875,995;875,303;806,303;806,313;806,50;631,50;631,61;570,61;570,51;308,51;308,62;308,18;31,18;31,893;171,893;31,893;31,1251;167,1251;167,1243;167,1270;191,1270;191,1476;648,1476;648,1384"

End Sub

Public Sub RenderBarcos(ByVal x As Integer, _
                        ByVal y As Integer, _
                        ByVal TileX As Integer, _
                        ByVal TileY As Integer, _
                        ByVal PixelOffSetX As Single, _
                        ByVal PixelOffSetY As Single)

    Dim I As Byte

    If Zonas(ZonaActual).Mapa > 1 Then Exit Sub

    For I = 0 To 1

        If Not Barco(I) Is Nothing Then
            If Barco(I).x = x And Barco(I).y = y Then
                Call Barco(I).Render(TileX - 4, TileY - 5, PixelOffSetX, PixelOffSetY)

            End If

        End If

    Next I

    Exit Sub

End Sub

Public Sub CalcularBarcos()

    Dim I As Byte

    For I = 0 To 1

        If Not Barco(I) Is Nothing Then
            Call Barco(I).Calcular

        End If

    Next I

End Sub
