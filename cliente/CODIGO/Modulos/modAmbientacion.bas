Attribute VB_Name = "modAmbientacion"

Public Enum TipoPaso

    CONST_BOSQUE = 1
    CONST_NIEVE = 2
    CONST_CABALLO = 3
    CONST_DUNGEON = 4
    CONST_PISO = 5
    CONST_DESIERTO = 6
    CONST_PESADO = 7

End Enum

Public Type tPaso

    CantPasos As Byte
    Wav() As Integer

End Type

Public Const NUM_PASOS  As Byte = 7

Public Pasos()          As tPaso

Public luz_dia(0 To 24) As D3DCOLORVALUE

Public Iluminacion      As Long

Public IluRGB           As D3DCOLORVALUE

Public Hora             As Byte

Private Function GetTerrenoDePaso(ByVal TerrainFileNum As Integer, _
                                  Terrain2FileNum) As TipoPaso

    If (TerrainFileNum >= 6000 And TerrainFileNum <= 6004) Or (TerrainFileNum >= 550 And TerrainFileNum <= 552) Or (TerrainFileNum >= 6018 And TerrainFileNum <= 6020) Or (TerrainFileNum >= 1478 And TerrainFileNum <= 1487) Or (TerrainFileNum >= 1548 And TerrainFileNum <= 1551) Or (TerrainFileNum >= 10013 And TerrainFileNum <= 10015) Or (TerrainFileNum >= 1073 And TerrainFileNum <= 1074) Or TerrainFileNum = 14638 Or TerrainFileNum = 14656 Or Terrain2FileNum = 8007 Then
        GetTerrenoDePaso = CONST_BOSQUE
        Exit Function
    ElseIf (TerrainFileNum >= 7501 And TerrainFileNum <= 7507) Or (TerrainFileNum = 7500 Or TerrainFileNum = 7508 Or TerrainFileNum = 1533 Or TerrainFileNum = 2508) Then
        GetTerrenoDePaso = CONST_DUNGEON
        Exit Function
    ElseIf (TerrainFileNum >= 13106 And TerrainFileNum <= 13115) Or Terrain2FileNum = 13117 Then
        GetTerrenoDePaso = CONST_NIEVE
        Exit Function
    ElseIf (TerrainFileNum >= 6018 And TerrainFileNum <= 6021) Or (TerrainFileNum >= 14551 And TerrainFileNum <= 14553) Or TerrainFileNum = 14564 Then
        GetTerrenoDePaso = CONST_DESIERTO
        Exit Function
    Else
        GetTerrenoDePaso = CONST_PISO

    End If

End Function

Public Sub CargarPasos()

    ReDim Pasos(1 To NUM_PASOS) As tPaso

    Pasos(TipoPaso.CONST_BOSQUE).CantPasos = 2
    ReDim Pasos(TipoPaso.CONST_BOSQUE).Wav(1 To Pasos(TipoPaso.CONST_BOSQUE).CantPasos) As Integer
    Pasos(TipoPaso.CONST_BOSQUE).Wav(1) = 193
    Pasos(TipoPaso.CONST_BOSQUE).Wav(2) = 194

    Pasos(TipoPaso.CONST_NIEVE).CantPasos = 2
    ReDim Pasos(TipoPaso.CONST_NIEVE).Wav(1 To Pasos(TipoPaso.CONST_NIEVE).CantPasos) As Integer
    Pasos(TipoPaso.CONST_NIEVE).Wav(1) = 195
    Pasos(TipoPaso.CONST_NIEVE).Wav(2) = 196

    Pasos(TipoPaso.CONST_DUNGEON).CantPasos = 2
    ReDim Pasos(TipoPaso.CONST_DUNGEON).Wav(1 To Pasos(TipoPaso.CONST_DUNGEON).CantPasos) As Integer
    Pasos(TipoPaso.CONST_DUNGEON).Wav(1) = 23
    Pasos(TipoPaso.CONST_DUNGEON).Wav(2) = 24

    Pasos(TipoPaso.CONST_DESIERTO).CantPasos = 2
    ReDim Pasos(TipoPaso.CONST_DESIERTO).Wav(1 To Pasos(TipoPaso.CONST_DESIERTO).CantPasos) As Integer
    Pasos(TipoPaso.CONST_DESIERTO).Wav(1) = 197
    Pasos(TipoPaso.CONST_DESIERTO).Wav(2) = 198

    Pasos(TipoPaso.CONST_PISO).CantPasos = 2
    ReDim Pasos(TipoPaso.CONST_PISO).Wav(1 To Pasos(TipoPaso.CONST_PISO).CantPasos) As Integer
    Pasos(TipoPaso.CONST_PISO).Wav(1) = 23
    Pasos(TipoPaso.CONST_PISO).Wav(2) = 24

    Pasos(TipoPaso.CONST_PESADO).CantPasos = 3
    ReDim Pasos(TipoPaso.CONST_PESADO).Wav(1 To Pasos(TipoPaso.CONST_PESADO).CantPasos) As Integer
    Pasos(TipoPaso.CONST_PESADO).Wav(1) = 220
    Pasos(TipoPaso.CONST_PESADO).Wav(2) = 221
    Pasos(TipoPaso.CONST_PESADO).Wav(3) = 222

End Sub

Public Sub DoPasosFx(ByVal CharIndex As Integer)

    Dim FileNum       As Integer

    Dim FileNum2      As Integer

    Dim TerrenoDePaso As TipoPaso

    If UserNavegando Or HayAgua(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y) Then
        If RandomNumber(1, 5) = 1 Then
            Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)

        End If

        'ElseIf charlist(CharIndex).equitando = True Then
        '    Call Audio.PlayWave(68, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    Else

        With charlist(CharIndex)

            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5 Or CharIndex = UserCharIndex) Then
            
                FileNum = MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex

                If FileNum > 0 Then FileNum = GrhData(FileNum).FileNum
                FileNum2 = MapData(.Pos.X, .Pos.Y).Graphic(2).GrhIndex

                If FileNum2 > 0 Then FileNum2 = GrhData(FileNum2).FileNum
                    
                TerrenoDePaso = GetTerrenoDePaso(FileNum, FileNum2)
            
                .pie = Not .pie
                
                If .pie Then
                    Call Audio.PlayWave(Pasos(TerrenoDePaso).Wav(1), .Pos.X, .Pos.Y)
                Else
                    Call Audio.PlayWave(Pasos(TerrenoDePaso).Wav(2), .Pos.X, .Pos.Y)

                End If

            End If

        End With

    End If

End Sub

Public Sub setup_ambient()

    'Noche 87, 61, 43
    luz_dia(0).R = 225
    luz_dia(0).G = 225
    luz_dia(0).b = 225
    luz_dia(1).R = 255
    luz_dia(1).G = 255
    luz_dia(1).b = 255
    luz_dia(2).R = 250
    luz_dia(2).G = 250
    luz_dia(2).b = 250
    luz_dia(3).R = 245
    luz_dia(3).G = 245
    luz_dia(3).b = 245
    '4 am 124,117,91
    luz_dia(4).R = 255
    luz_dia(4).G = 225
    luz_dia(4).b = 225
    '5,6 am 143,137,135
    luz_dia(5).R = 150
    luz_dia(5).G = 150
    luz_dia(5).b = 150
    luz_dia(6).R = 75
    luz_dia(6).G = 75
    luz_dia(6).b = 75
    '7 am 212,205,207
    luz_dia(7).R = 150
    luz_dia(7).G = 150
    luz_dia(7).b = 150
    luz_dia(8).R = 225
    luz_dia(8).G = 225
    luz_dia(8).b = 225
    luz_dia(9).R = 255
    luz_dia(9).G = 255
    luz_dia(9).b = 255
    luz_dia(10).R = 225
    luz_dia(10).G = 225
    luz_dia(10).b = 225
    luz_dia(11).R = 150
    luz_dia(11).G = 150
    luz_dia(11).b = 150
    luz_dia(12).R = 75
    luz_dia(12).G = 75
    luz_dia(12).b = 75
    'Dia 255, 255, 255
    luz_dia(12).R = 150
    luz_dia(12).G = 150
    luz_dia(12).b = 150
    luz_dia(13).R = 225
    luz_dia(13).G = 225
    luz_dia(13).b = 225
    'Medio Dia 255, 200, 255
    luz_dia(14).R = 255
    luz_dia(14).G = 255
    luz_dia(14).b = 255
    luz_dia(15).R = 225
    luz_dia(15).G = 225
    luz_dia(15).b = 225
    luz_dia(16).R = 150
    luz_dia(16).G = 150
    luz_dia(16).b = 150
    '17/18 0, 100, 255
    luz_dia(17).R = 75
    luz_dia(17).G = 75
    luz_dia(17).b = 75
    '18/19 0, 100, 255
    luz_dia(18).R = 150
    luz_dia(18).G = 150
    luz_dia(18).b = 150
    '19/20 156, 142, 83
    luz_dia(19).R = 225
    luz_dia(19).G = 225
    luz_dia(19).b = 225
    luz_dia(20).R = 255
    luz_dia(20).G = 255
    luz_dia(20).b = 255
    luz_dia(21).R = 225
    luz_dia(21).G = 225
    luz_dia(21).b = 225
    luz_dia(22).R = 150
    luz_dia(22).G = 150
    luz_dia(22).b = 150
    luz_dia(23).R = 75
    luz_dia(23).G = 75
    luz_dia(23).b = 75
    luz_dia(24).R = 150
    luz_dia(24).G = 150
    luz_dia(24).b = 150

End Sub

Public Sub SetDayLight(Optional ByVal WithSound As Boolean = False)
    'Dim pHora As Byte
    'If Zonas(ZonaActual).Terreno = eTerreno.Dungeon Then
    '    pHora = 21
    '    Hora = pHora
    'Else
    '    pHora = Hora
    
    frmMain.lblDIATEST.Caption = "Hora: " & Hora & " - " & "pHora: " & pHora & "-" & "Transp: " & luz_dia(pHora).R
   
    If WithSound = True Then

        Select Case luz_dia(Hora).R

            Case 75 'Noche
                Audio.PlayWave (81)

            Case 150 'Tarde
                Audio.PlayWave (53)

            Case 225 ' Mañana/Mediodia
                Audio.PlayWave (64)

            Case 255 ' Dia
                Audio.PlayWave (63)

        End Select

    End If

    'Exit Sub
If GoingHome = 1 Then
    IluRGB.R = 75
    IluRGB.G = 75
    IluRGB.b = 75
    
    Iluminacion = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)
    ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, bAlpha)
Else
   IluRGB.R = 255
    IluRGB.G = 255
    IluRGB.b = 255
    
    Iluminacion = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)
    ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, bAlpha)
End If
   

 

End Sub

Public Sub DoRelampago()

    On Error GoTo err

    If Zonas(ZonaActual).Terreno = eTerreno.Dungeon Then Exit Sub

    randomRelampagoX = RandomNumber(charlist(UserCharIndex).Pos.X - 10, charlist(UserCharIndex).Pos.X + 10)
    randomRelampagoY = RandomNumber(charlist(UserCharIndex).Pos.Y - 10, charlist(UserCharIndex).Pos.Y + 10)
    
    Call Audio.PlayWave(105, randomRelampagoX, randomRelampagoY)

    If bTecho = True Then Exit Sub

    IluRGB.R = 255
    IluRGB.G = 247
    IluRGB.b = 210

    OrigHora = Hora
    Hora = 25

    Iluminacion = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)
    ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, bAlpha)

    AlphaRelampago = 255
    HayRelampago = True

    Call SetAreaFx(randomRelampagoX, randomRelampagoY, 61, 0)
err:

End Sub

