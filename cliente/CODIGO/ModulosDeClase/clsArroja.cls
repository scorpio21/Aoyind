VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArroja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const PI = 3.14159265

Const PIs2 = 1.570796325

Dim TIPO        As Integer

Dim Particulas  As clsParticulas

Dim Particulas2 As clsParticulas

Dim Empieza     As Position

Dim Termina     As Position

Dim Efecto      As Byte

Dim Ang         As Single

Dim Radio       As Single

Dim Ang2        As Single

Dim Atacante    As Integer

Dim Victima     As Integer

Dim fX          As Integer

Dim Wav         As Byte

Dim Loops       As Integer

Dim Llego       As Boolean

Function CalcularAngulo(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Single

    Dim Relacion As Single

    Dim Esquina  As Integer

    If X1 < X2 And Y1 >= Y2 Then
        Esquina = 0
    ElseIf X1 >= X2 And Y1 > Y2 Then
        Esquina = 1
    ElseIf X1 > X2 And Y1 <= Y2 Then
        Esquina = 2
    ElseIf X1 <= X2 And Y1 < Y2 Then
        Esquina = 3

    End If

    If (Esquina = 0 Or Esquina = 2) And X2 <> X1 Then
        Relacion = (Y1 - Y2) / (X2 - X1)
    ElseIf Y2 <> Y1 Then
        Relacion = (X2 - X1) / (Y2 - Y1)

    End If

    CalcularAngulo = PI * 2 - Atn(Relacion) - Esquina * PI / 2

End Function

Function CalcularDistancia(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Single
    CalcularDistancia = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)

End Function

Public Sub Init(ByVal pAtacante As Integer, _
                ByVal pVictima As Integer, _
                ByVal pFX As Integer, _
                ByVal pLoops As Integer, _
                ByVal pWav As Byte, _
                ByVal pEfecto As Byte, _
                ByVal x As Integer, _
                ByVal y As Integer, _
                ByVal Xd As Integer, _
                Yd As Integer)

    Atacante = pAtacante
    Victima = pVictima
    fX = pFX
    Wav = pWav
    Loops = pLoops
    Empieza.x = x
    Empieza.y = y

    If Victima = 0 Then
        Termina.x = Xd
        Termina.y = Yd
    Else
        Termina.x = charlist(Victima).Pos.x
        Termina.y = charlist(Victima).Pos.y

    End If

    Efecto = pEfecto

    Ang = CalcularAngulo(Empieza.x, Empieza.y, Termina.x, Termina.y)
    Ang2 = PIs2

    Set Particulas = New clsParticulas

    If Efecto = eEffects.Inmovilizar Then
        Particulas.Texture = 14341
        Particulas.ParticleCounts = 70
        Particulas.b = 1
        Particulas.Efecto = eEffects.Inmovilizar
        Set Particulas2 = New clsParticulas
        Particulas2.Texture = 14341
        Particulas2.ParticleCounts = 70
        Particulas2.G = 1
        Particulas2.Efecto = eEffects.Inmovilizar
    
        'Flechas
    ElseIf Efecto = eEffects.Flecha1 Or Efecto = eEffects.Flecha2 Or Efecto = eEffects.Flecha3 Or Efecto = eEffects.Flecha4 Then
    
        Select Case Efecto

            Case 2
                Particulas.Texture = 747

            Case 3
                Particulas.Texture = 748

            Case 4
                Particulas.Texture = 749

            Case 5
                Particulas.Texture = 750

        End Select
    
        Efecto = eEffects.Flecha1
        Particulas.ParticleCounts = 1
        Particulas.R = 1
        Particulas.G = 1
        Particulas.b = 1
        Particulas.Desa = 1
        Particulas.Efecto = eEffects.Flecha1
        Particulas.Angulo = Ang + 0.785398163397448
    
    ElseIf Efecto = eEffects.Bala Then
        Particulas.Texture = 14780
        Efecto = eEffects.Bala
        Particulas.ParticleCounts = 1
        Particulas.R = 1
        Particulas.G = 1
        Particulas.b = 1
        Particulas.Desa = 1
        Particulas.Efecto = eEffects.Bala
        Particulas.Angulo = Ang + 0.785398163397448
    
    End If

End Sub

Public Sub Render(TileX As Integer, _
                  TileY As Integer, _
                  ScreenOffX As Single, _
                  ScreenOffY As Single)

    If Efecto = eEffects.Inmovilizar Then
        Call MovimientoInmo(TileX, TileY, ScreenOffX, ScreenOffY)
    ElseIf Efecto = eEffects.Flecha1 Then
        Call MovimientoFlecha(TileX, TileY, ScreenOffX, ScreenOffY)
    ElseIf Efecto = eEffects.Bala Then
        Call MovimientoBala(TileX, TileY, ScreenOffX, ScreenOffY)

    End If

End Sub

Sub MovimientoInmo(TileX As Integer, _
                   TileY As Integer, _
                   ScreenOffX As Single, _
                   ScreenOffY As Single)

    Dim Movimiento As Byte

    If Radio > CalcularDistancia(Empieza.x, Empieza.y, Termina.x, Termina.y) * 32 Then
        'Arrojas.Remove (1)
    
        Particulas.Matar = True
        Particulas2.Matar = True
        Movimiento = 2
    
        If Not Llego Then
            If fX > 0 Then
                Call SetCharacterFx(Victima, fX, Loops)

            End If

            If Wav > 0 Then
                Call Audio.PlayWave(CStr(Wav) & "", charlist(Victima).Pos.x, charlist(Victima).Pos.y)

            End If

            Llego = True

        End If

    Else
        Radio = Radio + 2 * timerElapsedTime / 5
        Movimiento = 1

    End If

    Ang2 = Ang2 + 0.06 * timerElapsedTime / 5

    'La posicion de la particulas siempre es relativa a su origen
    Call Particulas.Move((Empieza.x - TileX + 16) * 32 + ScreenOffX, (Empieza.y - TileY + 13) * 32 + ScreenOffY - 24)
    Call Particulas2.Move((Empieza.x - TileX + 16) * 32 + ScreenOffX, (Empieza.y - TileY + 13) * 32 + ScreenOffY - 24)
    
    If Movimiento = 1 Then

        Call Particulas.ReLocate(16 + Cos(Ang) * Radio + Cos(Ang2) * 25 * Cos(Ang + PIs2), Sin(Ang) * Radio + Cos(Ang2) * 25 * Sin(Ang + PIs2))
        Call Particulas2.ReLocate(16 + Cos(Ang) * Radio - Cos(Ang2) * 25 * Cos(Ang + PIs2), Sin(Ang) * Radio - Cos(Ang2) * 25 * Sin(Ang + PIs2))
    Else
        'Call Particulas.Move((Termina.x - TileX + 8) * 32 + ScreenOffX, (Termina.y - TileY + 7) * 32 + ScreenOffY - 24)
        'Call Particulas2.Move((Termina.x - TileX + 8) * 32 + ScreenOffX, (Termina.y - TileY + 7) * 32 + ScreenOffY - 24)
        'offX = charlist(Victima).Pos.x - Termina.x + charlist(Victima).MoveOffsetX
        'offX = charlist(Victima).Pos.y - Termina.y + charlist(Victima).MoveOffsetY

        Call Particulas.ReLocate((charlist(Victima).Pos.x - Termina.x) * 32 + charlist(Victima).MoveOffsetX + 16 + Cos(Ang) * Radio + Sin(Ang2 * 2 / 2) * 25, (charlist(Victima).Pos.y - Termina.y) * 32 + charlist(Victima).MoveOffsetY + Sin(Ang) * Radio + Sin(Ang2 * 1 / 2) * 25)
        Call Particulas2.ReLocate((charlist(Victima).Pos.x - Termina.x) * 32 + charlist(Victima).MoveOffsetX + 16 + Cos(Ang) * Radio + Sin((Ang2 + 2) * 1 / 2) * 25, (charlist(Victima).Pos.y - Termina.y) * 32 + charlist(Victima).MoveOffsetY + Sin(Ang) * Radio + Sin((Ang2 + 2) * 2 / 2) * 25)

    End If

    'Sin(Ang2 * ax / 10) * 30 + 400, bx * Sin(Ang2 * bx / 10) * 30 + 400
    'Call Particulas.ReLocate(Cos(Ang) * Radio, Sin(Ang) * Radio)

    Call Particulas.Update
    Call Particulas.Render
    Call Particulas2.Update
    Call Particulas2.Render

    If Particulas.Desa = 0 Then
        Arrojas.Remove (1)

    End If

End Sub

Sub MovimientoFlecha(TileX As Integer, _
                     TileY As Integer, _
                     ScreenOffX As Single, _
                     ScreenOffY As Single)

    'If Radio >= CalcularDistancia(Empieza.x, Empieza.Y, Termina.x, Termina.Y) * 32 - 12 Then
    If Radio >= CalcularDistancia(Empieza.x, Empieza.y, Termina.x, Termina.y) * 32 - 12 Then
        Arrojas.Remove (1)
    
        If fX > 0 Then
            Call SetCharacterFx(Victima, fX, Loops)

        End If

        If Wav > 0 Then
            Call Audio.PlayWave(CStr(Wav) & "", charlist(Victima).Pos.x, charlist(Victima).Pos.y)

        End If

    Else
        Radio = Radio + timerElapsedTime * 1.5

    End If

    'La posicion de la particulas siempre es relativa a su origen
    Call Particulas.Move((Empieza.x - TileX + 12) * 32 + ScreenOffX, (Empieza.y - TileY + 10) * 32 + ScreenOffY - 24)

    Call Particulas.ReLocate(Cos(Ang) * Radio, Sin(Ang) * Radio - 20)

    'Sin(Ang2 * ax / 10) * 30 + 400, bx * Sin(Ang2 * bx / 10) * 30 + 400
    'Call Particulas.ReLocate(Cos(Ang) * Radio, Sin(Ang) * Radio)

    Call Particulas.Update
    Call Particulas.Render

End Sub

Sub MovimientoBala(TileX As Integer, _
                   TileY As Integer, _
                   ScreenOffX As Single, _
                   ScreenOffY As Single)

    If Radio >= CalcularDistancia(Empieza.x, Empieza.y, Termina.x, Termina.y) * 32 - 12 Then
    
        Arrojas.Remove (1)

        Dim destX As Integer

        Dim destY As Integer

        destX = Termina.x
        destY = Termina.y
    
        'Hacemos el daño
        If Atacante = UserCharIndex Then
            Call WriteCreateEfectoClientAction(Efecto, destX, destY)

        End If
    
        If fX > 0 Then
            If Victima > 0 Then
                Call SetCharacterFx(Victima, fX, Loops)
            Else
                Call SetAreaFx(Termina.x, Termina.y + 2, fX, Loops)

            End If

        End If
    
        If Wav > 0 Then
            Call Audio.PlayWave(Wav, Termina.x, Termina.y)

        End If
    
    Else
        Radio = Radio + timerElapsedTime

    End If

    'La posicion de la particulas siempre es relativa a su origen
    Call Particulas.Move((Empieza.x - TileX + 12) * 32 + ScreenOffX, (Empieza.y - TileY + 10) * 32 + ScreenOffY - 24)

    Call Particulas.ReLocate(Cos(Ang) * Radio, Sin(Ang) * Radio - 20)

    'Sin(Ang2 * ax / 10) * 30 + 400, bx * Sin(Ang2 * bx / 10) * 30 + 400
    'Call Particulas.ReLocate(Cos(Ang) * Radio, Sin(Ang) * Radio)

    Call Particulas.Update
    Call Particulas.Render

End Sub

