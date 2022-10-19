Attribute VB_Name = "Mod_TileEngine"
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

Public OffSetConsola As Byte

Public Const ComienzoY As Integer = 604

Public UltimaLineavisible As Boolean

Public Const MaxLineas As Byte = 6

Type TConsola

    T As String
    '   Color As Long
    R As Byte
    G As Byte
    b As Byte

End Type

Public Con(1 To MaxLineas) As TConsola

Public Declare Function TransparentBlt _
                         Lib "msimg32.dll" (ByVal hDCDest As Long, _
                                            ByVal nXOriginDest As Long, _
                                            ByVal nYOriginDest As Long, _
                                            ByVal nWidthDest As Long, _
                                            ByVal nHeightDest As Long, _
                                            ByVal hDCSrc As Long, _
                                            ByVal nXOriginSrc As Long, _
                                            ByVal nYOriginSrc As Long, _
                                            ByVal nWidthSrc As Long, _
                                            ByVal nHeightSrc As Long, _
                                            ByVal crTransparent As Long) As Long

'Map sizes in tiles
Public Const XMaxMapSize As Integer = 1100

Public Const YMaxMapSize As Integer = 1500

Public Const RelacionMiniMapa As Single = 1.92120075046904

Public Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Encabezado bmp
Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

'Posicion en un mapa
Public Type Position

    X As Long
    Y As Long

End Type

'Posicion en el Mundo
Public Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Type GrhData

    sX As Integer
    sY As Integer

    FileNum As Integer

    PixelWidth As Integer
    PixelHeight As Integer

    TileWidth As Single
    TileHeight As Single

    NumFrames As Integer
    Frames() As Integer

    Speed As Single

End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    'particulas ore
    angle As Single
    'particulas ore

    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer

End Type

'Lista de cuerpos
Type BodyData

    Walk(E_Heading.north To E_Heading.west) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas
Type HeadData

    Head(E_Heading.north To E_Heading.west) As Grh

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

    WeaponWalk(E_Heading.north To E_Heading.west) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

    ShieldWalk(E_Heading.north To E_Heading.west) As Grh
    '[ANIM ATAK]
    ShieldAttack As Byte

End Type

Type Alas
    GrhIndex(E_Heading.north To E_Heading.west) As Grh
End Type

Public alaArray() As Alas

Public NPCMuertos As New Collection

'Apariencia del personaje
Public Type Char
    alaIndex As Byte
    Alas As Alas
    'particulas ore
    particle_count As Integer
    particle_group() As Long
    'particulas ore

    'quest
    simbolo As Byte
    'quest
    'Render
    Elv As Byte
    Gld As Long
    Clase As Byte

    equitando As Boolean
    congelado As Boolean
    Chiquito As Boolean
    nadando As Boolean
    inmovilizado As Boolean

    ACTIVE As Byte
    Heading As E_Heading
    Pos As Position
    LastPos As Position

    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean

    fX As Grh
    FxIndex As Integer

    Criminal As Byte

    nombre As String

    scrollDirectionX As Integer
    scrollDirectionY As Integer

    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single

    pie As Boolean
    logged As Boolean
    muerto As Boolean
    invisible As Boolean
    oculto As Boolean
    Alpha As Byte
    ContadorInvi As Integer
    iTick As Long
    priv As Byte
    'aura
    aura(0 To 5) As tAuras
    'aura
    Quieto As Byte

End Type

'Info de un objeto
Public Type Obj

    OBJIndex As Integer
    Amount As Integer

End Type

Public Type MapInformation

    Name As String
    MapVersion As Integer
    Width As Integer
    Height As Integer
    offset As Integer
    Date As String

End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    'particulas ore
    particle_group As Integer
    'particulas ore
    PasosIndex As Byte

    Graphic(1 To 5) As Grh
    CharIndex As Integer
    ObjGrh As Grh

    particle_group_index As Integer

    Blocked As Byte
    Trigger As Byte
    'Particulas
    particle_index As Integer
    'Particulas
    Map As Byte
    Elemento As Object

    light_value(3) As Long
    Hora As Byte

    fX As Integer
    fXGrh As Grh
    'sangre
    Blood As Byte
    Blood2 As Byte
    'sangre
End Type

Public IniPath As String

Public MapPath As String

'Status del user
Public CurMap As Integer            'Mapa actual

Public UserIndex As Integer

Public UserMoving As Byte

Public UserBody As Integer

Public UserHead As Integer

Public UserPos As Position             'Posicion

Public AddtoUserPos As Position             'Si se mueve

Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long

Public FramesPerSecCounter As Long

Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer

Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer

Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer

Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer

Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer

Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer

Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single

Public timerTicksPerFrame As Single

Public engineBaseSpeed As Single

Public NumBodies As Integer

Public Numheads As Integer

Public NumFxs As Integer

Public NumWeaponAnims As Integer

Public NumShieldAnims As Integer

Public NumChars As Integer

Public LastChar As Integer

Private MainDestRect As RECT

Private MainViewRect As RECT

Private BackBufferRect As RECT

Private MainViewWidth As Integer

Private MainViewHeight As Integer

Private MouseTileX As Integer

Private MouseTileY As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData            'Guarda todos los grh

Public BodyData() As BodyData

Public HeadData() As HeadData

Public FxData() As tIndiceFx

Public WeaponAnimData() As WeaponAnimData

Public ShieldAnimData() As ShieldAnimData

Public CascoAnimData() As HeadData

Public Arrojas As New Collection

Public Tooltips As New Collection
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData(1 To XMaxMapSize, 1 To YMaxMapSize) As MapBlock    ' Mapa
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain As Boolean            'está raineando?

Public bTecho As Boolean            'hay techo?

Public bAlpha As Byte

Public nAlpha As Byte

Public tTick As Long

Public ColorTecho As Long

Public brstTick As Long

Private RLluvia(7) As RECT          'RECT de la lluvia

Private iFrameIndex As Byte          'Frame actual de la LL

Private llTick As Long          'Contador

Private LTLluvia(7) As Integer

Public charlist(1 To 10000) As Char

Public AperturaPergamino As Single

#If SeguridadAlkon Then

    Public MI(1 To 1233) As clsManagerInvisibles

    Public CualMI As Integer

#End If

' Used by GetTextExtentPoint32
Private Type Size

    cx As Long
    cy As Long

End Type

'[CODE 001]:MatuX
Public Enum PlayLoop

    plNone = 0
    plLluviain = 1
    plLluviaout = 2

End Enum

'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public FrameTime As Long

Public MapaY As Single

Public VerMapa As Boolean

Public Entrada As Byte

Public FrameUseMotionBlur As Boolean

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency _
                          Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                          Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function BitBlt _
                          Lib "gdi32" (ByVal hDestDC As Long, _
                                       ByVal X As Long, _
                                       ByVal Y As Long, _
                                       ByVal nWidth As Long, _
                                       ByVal nHeight As Long, _
                                       ByVal hSrcDC As Long, _
                                       ByVal xSrc As Long, _
                                       ByVal ySrc As Long, _
                                       ByVal dwRop As Long) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 _
                          Lib "gdi32" _
                              Alias "GetTextExtentPoint32A" (ByVal hDC As Long, _
                                                             ByVal lpsz As String, _
                                                             ByVal cbString As Long, _
                                                             lpSize As Size) As Long

Public PosMapX As Single

Public PosMapY As Single

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
'Public FrameTime          As Long

Sub CargarCabezas()

    Dim N            As Integer

    Dim I            As Long

    Dim Numheads     As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open PathInit & "\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For I = 1 To Numheads
        Get #N, , Miscabezas(I)
        
        If Miscabezas(I).Head(1) Then
            Call InitGrh(HeadData(I).Head(1), Miscabezas(I).Head(1), 0)
            Call InitGrh(HeadData(I).Head(2), Miscabezas(I).Head(2), 0)
            Call InitGrh(HeadData(I).Head(3), Miscabezas(I).Head(3), 0)
            Call InitGrh(HeadData(I).Head(4), Miscabezas(I).Head(4), 0)

        End If

    Next I
    
    Close #N

End Sub

Sub CargarCascos()

    Dim N            As Integer

    Dim I            As Long

    Dim NumCascos    As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open PathInit & "\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For I = 1 To NumCascos
        Get #N, , Miscabezas(I)
        
        If Miscabezas(I).Head(1) Then
            Call InitGrh(CascoAnimData(I).Head(1), Miscabezas(I).Head(1), 0)
            Call InitGrh(CascoAnimData(I).Head(2), Miscabezas(I).Head(2), 0)
            Call InitGrh(CascoAnimData(I).Head(3), Miscabezas(I).Head(3), 0)
            Call InitGrh(CascoAnimData(I).Head(4), Miscabezas(I).Head(4), 0)

        End If

    Next I
    
    Close #N

End Sub

Sub CargarCuerpos()

    Dim N            As Integer

    Dim I            As Long

    Dim NumCuerpos   As Integer

    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open PathInit & "\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For I = 1 To NumCuerpos
        Get #N, , MisCuerpos(I)
        
        If MisCuerpos(I).Body(1) Then
            InitGrh BodyData(I).Walk(1), MisCuerpos(I).Body(1), 0
            InitGrh BodyData(I).Walk(2), MisCuerpos(I).Body(2), 0
            InitGrh BodyData(I).Walk(3), MisCuerpos(I).Body(3), 0
            InitGrh BodyData(I).Walk(4), MisCuerpos(I).Body(4), 0
            
            BodyData(I).HeadOffset.X = MisCuerpos(I).HeadOffsetX
            BodyData(I).HeadOffset.Y = MisCuerpos(I).HeadOffsetY

        End If

    Next I
    
    Close #N

End Sub

Sub CargarFxs()

    Dim N      As Integer

    Dim I      As Long

    Dim NumFxs As Integer
    
    N = FreeFile()
    Open PathInit & "\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For I = 1 To NumFxs
        Get #N, , FxData(I)
    Next I
    
    Close #N

End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, _
                  ByVal viewPortY As Integer, _
                  ByRef tX As Integer, _
                  ByRef tY As Integer)
    '******************************************
    'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
    '******************************************

    tX = UserPos.X + (viewPortX + 16) \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + (viewPortY + 16) \ TilePixelHeight - WindowTileHeight \ 2

    'frmMain.lblPosTest2.Caption = "X: " & tX & "; Y:" & tY
   
End Sub

Sub MakeChar(ByVal CharIndex As Integer, _
             ByVal Body As Integer, _
             ByVal Head As Integer, _
             ByVal Heading As Byte, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal Arma As Integer, _
             ByVal Escudo As Integer, _
             ByVal Casco As Integer)

    On Error Resume Next

    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)

        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .ACTIVE = 0 Then NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        '[ANIM ATAK]
        .Arma.WeaponAttack = 0
        .Escudo.ShieldAttack = 0
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        .Alpha = 255
        .iTick = 0
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        .muerto = Head = CASPER_HEAD Or Head = CASPER_HEAD_CRIMI Or Body = FRAGATA_FANTASMAL

        If .muerto Then .Alpha = 80 Else .Alpha = 255
        'Make active
        .ACTIVE = 1
        
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
    
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    With charlist(CharIndex)
        .Gld = 0
        .Elv = 1
        .Clase = 0
    
        .equitando = False
        .ACTIVE = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
        
        #If SeguridadAlkon Then
            Call MI(CualMI).ResetInvisible(CharIndex)
        #End If
        
        .Moving = 0
        .muerto = False
        .Alpha = 255
        .iTick = 0
        .ContadorInvi = 0
        .nombre = ""
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .LastPos.X = 0
        .LastPos.Y = 0
        .UsandoArma = False

    End With

End Sub

Sub EraseChar(ByVal CharIndex As Integer)

    '*****************************************************************
    'Erases a character from CharList and map
    '*****************************************************************
    On Error Resume Next

    With charlist(CharIndex)

        .ACTIVE = 0
    
        'Update lastchar
        If CharIndex = LastChar Then

            Do Until charlist(LastChar).ACTIVE = 1
                LastChar = LastChar - 1

                If LastChar = 0 Then Exit Do
            Loop

        End If

        If .Pos.X > 0 And .Pos.Y > 0 Then
            MapData(.Pos.X, .Pos.Y).CharIndex = 0
    
            If .FxIndex <> 0 And .fX.Loops > -1 Then
                MapData(.Pos.X, .Pos.Y).fX = .FxIndex
                MapData(.Pos.X, .Pos.Y).fXGrh = .fX

            End If
    
            'Remove char's dialog
            Call Dialogos.RemoveDialog(CharIndex)
    
            Call ResetCharInfo(CharIndex)
    
            'Update NumChars
            NumChars = NumChars - 1

        End If
    
    End With

End Sub

Public Sub InitGrh(ByRef Grh As Grh, _
                   ByVal GrhIndex As Integer, _
                   Optional ByVal Started As Byte = 2)
    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    Grh.GrhIndex = GrhIndex

    If Grh.GrhIndex = 0 Then Exit Sub
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0

        End If

    Else

        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started

    End If
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0

    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed

End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************

    Dim AddX   As Integer

    Dim AddY   As Integer

    Dim X      As Integer

    Dim Y      As Integer

    Dim nX     As Integer

    Dim nY     As Integer

    Dim tmpInt As Integer
   ' On Error GoTo err
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        If X = 0 Or Y = 0 Then Exit Sub
        
        .LastPos.X = X
        .LastPos.Y = Y
        
        'Figure out which way to move
        Select Case nHeading

            Case E_Heading.north
                AddY = -1
        
            Case E_Heading.east
                AddX = 1
        
            Case E_Heading.south
                AddY = 1
            
            Case E_Heading.west
                AddX = -1

        End Select
        
        nX = X + AddX
        nY = Y + AddY
        
        If MapData(nX, nY).CharIndex > 0 Then
            tmpInt = MapData(nX, nY).CharIndex

            If charlist(tmpInt).muerto = False Then
                tmpInt = 0
            Else
                charlist(tmpInt).Pos.X = X
                charlist(tmpInt).Pos.Y = Y
                charlist(tmpInt).Heading = InvertHeading(nHeading)
                charlist(tmpInt).MoveOffsetX = 1 * (TilePixelWidth * AddX)
                charlist(tmpInt).MoveOffsetY = 1 * (TilePixelHeight * AddY)
                
                charlist(tmpInt).Moving = 1
                
                charlist(tmpInt).scrollDirectionX = -AddX
                charlist(tmpInt).scrollDirectionY = -AddY
                
                'Si el fantasma soy yo mueve la pantalla
                If tmpInt = UserCharIndex Then Call MoveScreen(charlist(tmpInt).Heading)

            End If

        Else
            tmpInt = 0

        End If

        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = tmpInt
        
        If UserEstado <> 1 Then
            Call vPasos.CreatePasos(X, Y, DamePasos(nHeading))

        End If
        
        .MoveOffsetX = -1 * (TilePixelWidth * AddX)
        .MoveOffsetY = -1 * (TilePixelHeight * AddY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = AddX
        .scrollDirectionY = AddY


        For X = .Pos.X - 5 To .Pos.X + 5
            For Y = .Pos.Y - 5 To .Pos.Y + 5
                If (.Pos.X <> X Or .Pos.Y <> Y) And MapData(X, Y).CharIndex = CharIndex Then
                    MapData(X, Y).CharIndex = 0
                End If
            Next Y
        Next X
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    'If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    '    If CharIndex <> UserCharIndex Then
    '        Call EraseChar(CharIndex)
    '    End If
    'End If
'err:
End Sub

Public Function InvertHeading(ByVal nHeading As E_Heading) As E_Heading

    Select Case nHeading

        Case E_Heading.east
            InvertHeading = west

        Case E_Heading.west
            InvertHeading = east

        Case E_Heading.south
            InvertHeading = north

        Case E_Heading.north
            InvertHeading = south

    End Select

End Function

Public Sub DoFogataFx()

    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)

        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0

        End If

    Else
        bFogata = HayFogata(location)

        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave(SND_FUEGO, location.X, location.Y, LoopStyle.Enabled)

    End If

End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean

    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - 11 And .X < UserPos.X + 111 And .Y > UserPos.Y - 9 And .Y < UserPos.Y + 9

    End With

End Function

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

    On Error Resume Next

    Dim X           As Integer

    Dim Y           As Integer

    Dim AddX        As Integer

    Dim AddY        As Integer

    Dim nHeading    As E_Heading

    Dim tmpInt      As Integer

    Dim hayColision As Boolean
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        If X > 0 And Y > 0 Then
                
            AddX = nX - X
            AddY = nY - Y
        
            If Sgn(AddX) = 1 Then
                nHeading = E_Heading.east
            ElseIf Sgn(AddX) = -1 Then
                nHeading = E_Heading.west
            ElseIf Sgn(AddY) = -1 Then
                nHeading = E_Heading.north
            ElseIf Sgn(AddY) = 1 Then
                nHeading = E_Heading.south

            End If
        
            If MapData(nX, nY).CharIndex > 0 Then
                tmpInt = MapData(nX, nY).CharIndex
            
                'Si está muerto lo pisamos
                If charlist(tmpInt).muerto = False Then

                    'Si pisó el PJ volvemos a su posición anterior
                    If MapData(nX, nY).CharIndex = UserCharIndex Then
                        Debug.Print "************************************************************ COLISIÓN ********************************************************************************"
                        WAIT_ACTION = eWAIT_FOR_ACTION.RPU
                        Call WriteRequestPositionUpdate

                    End If

                    tmpInt = 0
                Else
                    charlist(tmpInt).Pos.X = X
                    charlist(tmpInt).Pos.Y = Y
                    charlist(tmpInt).Heading = InvertHeading(nHeading)
                    charlist(tmpInt).MoveOffsetX = 1 * (TilePixelWidth * AddX)
                    charlist(tmpInt).MoveOffsetY = 1 * (TilePixelHeight * AddY)
                
                    charlist(tmpInt).Moving = 1
                
                    charlist(tmpInt).scrollDirectionX = -AddX
                    charlist(tmpInt).scrollDirectionY = -AddY
                
                    'Si el fantasma soy yo mueve la pantalla
                    If tmpInt = UserCharIndex Then Call MoveScreen(charlist(tmpInt).Heading)

                End If

            Else
                tmpInt = 0

            End If
        
            MapData(X, Y).CharIndex = tmpInt
        
            MapData(nX, nY).CharIndex = CharIndex
       
            '
            '         If hayColision = True Then
            '            'Call EraseChar(CharIndex)
            '            'charlist(UserCharIndex).Pos.x = .LastPos.x
            '            'charlist(UserCharIndex).Pos.y = .LastPos.y
            '            'Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY, True, nX, nY)
            '            'Call CharRender(charlist(UserCharIndex), UserCharIndex, charlist(UserCharIndex).Pos.X, charlist(UserCharIndex).Pos.Y)
            '        End If
        
            .Pos.X = nX
            .Pos.Y = nY
        
            .MoveOffsetX = -1 * (TilePixelWidth * AddX)
            .MoveOffsetY = -1 * (TilePixelHeight * AddY)
        
            .Moving = 1
            .Heading = nHeading
        
            .scrollDirectionX = Sgn(AddX)
            .scrollDirectionY = Sgn(AddY)
        
            'parche para que no medite cuando camina
            If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
                .FxIndex = 0

            End If

        End If
        
        If Not EstaPCarea(CharIndex) Then
            Call Dialogos.RemoveDialog(CharIndex)
        Else

            If .muerto = False Then
                Call vPasos.CreatePasos(X, Y, DamePasos(nHeading))

            End If

        End If
        

        For X = .Pos.X - 5 To .Pos.X + 5
            For Y = .Pos.Y - 5 To .Pos.Y + 5
                If (.Pos.X <> X Or .Pos.Y <> Y) And MapData(X, Y).CharIndex = CharIndex Then
                    MapData(X, Y).CharIndex = 0
                End If
            Next Y
        Next X
    End With
    

    
    '    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    '        Call EraseChar(CharIndex)
    '    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)

    '******************************************
    'Starts the screen moving in a direction
    '******************************************
    Dim X  As Integer

    Dim Y  As Integer

    Dim tX As Integer

    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading

        Case E_Heading.north
            Y = -1
        
        Case E_Heading.east
            X = 1
        
        Case E_Heading.south
            Y = 1
        
        Case E_Heading.west
            X = -1

    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < 1 Or tX > MapInfo.Width Or tY < 1 Or tY > MapInfo.Height Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 7 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

    End If

End Sub

Private Function HayFogata(ByRef location As Position) As Boolean

    Dim J As Long

    Dim k As Long
    
    For J = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6

            If InMapBounds(J, k) Then
                If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                    
                    location.X = J
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function

                End If

            End If

        Next k
    Next J

End Function

Function NextOpenChar() As Integer

    '*****************************************************************
    'Finds next open char slot in CharList
    '*****************************************************************
    Dim LoopC As Long

    Dim Dale  As Boolean
    
    LoopC = 1

    Do While charlist(LoopC).ACTIVE And Dale
        LoopC = LoopC + 1
        Dale = (LoopC <= UBound(charlist))
    Loop
    
    NextOpenChar = LoopC

End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean

    On Error GoTo ErrorHandler

    Dim Grh         As Long

    Dim Frame       As Long

    Dim grhCount    As Long

    Dim Handle      As Integer

    Dim fileVersion As Long
    
    'Open files
    Handle = FreeFile()
    
    Open IniPath & GraphicsFile For Binary Access Read As Handle
    Seek #1, 1
    
    'Get file version
    Get Handle, , fileVersion
    
    'Get number of grhs
    Get Handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(Handle)

        Get Handle, , Grh

        If Grh > 0 Then

            With GrhData(Grh)
                'Get number of frames
                Get Handle, , .NumFrames

                If .NumFrames <= 0 Then GoTo ErrorHandler
            
                ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
                If .NumFrames > 1 Then

                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get Handle, , .Frames(Frame)

                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                            GoTo ErrorHandler

                        End If

                    Next Frame
                
                    Get Handle, , .Speed
                
                    If .Speed <= 0 Then GoTo ErrorHandler
                
                    'Compute width and height
                    .PixelHeight = GrhData(.Frames(1)).PixelHeight
                    'If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                    .PixelWidth = GrhData(.Frames(1)).PixelWidth
                    'If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    'If .TileWidth <= 0 Then GoTo ErrorHandler
                
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    'If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get Handle, , .FileNum

                    If .FileNum <= 0 Then GoTo ErrorHandler
                
                    Get Handle, , GrhData(Grh).sX

                    If .sX < 0 Then GoTo ErrorHandler
                
                    Get Handle, , .sY

                    If .sY < 0 Then GoTo ErrorHandler
                
                    Get Handle, , .PixelWidth

                    If .PixelWidth <= 0 Then GoTo ErrorHandler
                
                    Get Handle, , .PixelHeight

                    If .PixelHeight <= 0 Then GoTo ErrorHandler
                
                    'Compute width and height
                    .TileWidth = .PixelWidth / TilePixelHeight
                    .TileHeight = .PixelHeight / TilePixelWidth
                
                    .Frames(1) = Grh

                End If

            End With

        End If

    Wend
    
    Close Handle
    
    LoadGrhData = True
    Exit Function

ErrorHandler:
    LoadGrhData = False

End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is legal
    '*****************************************************************
    'Limites del mapa
    If X < 1 Or X > MapInfo.Width Or Y < 1 Or Y > MapInfo.Height Then
        Exit Function

    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function

    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function

    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function

    End If
    
    LegalPos = True

End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 10/05/2009
    'Checks to see if a tile position is legal, including if there is a casper in the tile
    '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
    '*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < 1 Or X > MapInfo.Width Or Y < 1 Or Y > MapInfo.Height Then
        Exit Function

    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function

    End If
    
    CharIndex = MapData(X, Y).CharIndex

    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function

        End If
        
        With charlist(CharIndex)

            ' Si no es casper, no puede pasar
            If .muerto = False Or .nombre = "" Then
                Exit Function
            Else

                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else

                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function

                End If

            End If

        End With

    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function

    End If
    
    MoveToLegalPos = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************
    If X < 1 Or X > MapInfo.Width Or Y < 1 Or Y > MapInfo.Height Then
        Exit Function

    End If
    
    InMapBounds = True

End Function

Sub DrawGrhIndexLuz(ByVal GrhIndex As Integer, _
                    ByVal X As Integer, _
                    ByVal Y As Integer, _
                    ByVal Center As Byte, _
                    ByRef Color() As Long)

    With GrhData(GrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
        
        Call Engine_Render_Rectangle(X, Y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , , .FileNum, Color(0), Color(1), Color(2), Color(3))

    End With

End Sub

Sub DrawGrhIndex(ByVal GrhIndex As Integer, _
                 ByVal X As Integer, _
                 ByVal Y As Integer, _
                 ByVal Center As Byte, _
                 ByVal Color As Long)

    With GrhData(GrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
        
        Call Engine_Render_Rectangle(X, Y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , , .FileNum, Color, Color, Color, Color)
      
    End With

End Sub

Sub DrawGrhLuz(ByRef Grh As Grh, _
               ByVal X As Integer, _
               ByVal Y As Integer, _
               ByVal Center As Byte, _
               ByVal Animate As Single, _
               ByRef Color() As Long)

    Dim CurrentGrhIndex As Integer
    
    On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * Animate * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0

                    End If

                End If

            End If

        End If

    End If

    If Grh.GrhIndex > 0 Then
        'Figure out what frame to draw (always 1 if not animated)
        CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
        With GrhData(CurrentGrhIndex)

            'Center Grh over X,Y pos
            If Center Then
                If .TileWidth <> 1 Then
                    X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

                End If
            
                If .TileHeight <> 1 Then
                    Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

                End If

            End If
                
            'If COLOR = -1 Then COLOR = Iluminacion

            Call Engine_Render_Rectangle(X, Y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, Color(0), Color(1), Color(2), Color(3))

        End With

    End If

    Exit Sub

Error:

    If err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & vbCrLf & err.Description, vbExclamation, "[ " & err.Number & " ] Error"
        End

    End If

End Sub

Sub DrawGrhShadow(ByRef Grh As Grh, _
                  ByVal X As Integer, _
                  ByVal Y As Integer, _
                  ByVal Center As Byte, _
                  ByVal Animate As Single, _
                  Optional Shadow As Byte = 0, _
                  Optional Color As Long = -1, _
                  Optional ShadowAlpha As Single = 255, _
                  Optional Chiquitolin As Boolean = False)

    Dim CurrentGrhIndex As Integer
    
    On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * Animate * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0

                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
        
        Dim PixelWidth  As Integer

        Dim PixelHeight As Integer
        
        ' <<<< CHIQUITOLIN >>>>
        If Chiquitolin = True Then
            PixelWidth = PixelWidth * 0.7
            PixelHeight = PixelHeight * 0.7
        Else
            PixelWidth = .PixelWidth
            PixelHeight = .PixelHeight

        End If
                
        'Draw
        'Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
      
        If mOpciones.Shadows = True And Chiquitolin = False And Conectar = False Then
            If Shadow = 1 Then
                ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
                Call Engine_Render_Rectangle(X, Y, PixelWidth, PixelHeight, .sX, .sY, PixelWidth, PixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
            ElseIf Shadow = 2 Then
                ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
                Call Engine_Render_Rectangle(X + 10, Y - 16, .PixelWidth, PixelHeight, .sX, .sY, PixelWidth, PixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)

            End If

        End If
      
        If Color = -1 Then Color = Iluminacion

    End With

    Exit Sub

Error:

    If err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & vbCrLf & err.Description, vbExclamation, "[ " & err.Number & " ] Error"
        End

    End If

End Sub

Sub DrawGrhShadowOff(ByRef Grh As Grh, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal Center As Byte, _
                     ByVal Animate As Single, _
                     Optional Color As Long = -1, _
                     Optional Chiquitolin As Boolean = False)

    Dim CurrentGrhIndex As Integer
    
    On Error GoTo Error
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
                
        If Color = -1 Then Color = Iluminacion

        Dim PixelWidth  As Integer

        Dim PixelHeight As Integer
        
        ' <<<< CHIQUITOLIN >>>>
        If Chiquitolin = True Then
            PixelWidth = .PixelWidth * 0.7
            PixelHeight = .PixelHeight * 0.7
        Else
            PixelWidth = .PixelWidth
            PixelHeight = .PixelHeight

        End If

        Call Engine_Render_Rectangle(X, Y, PixelWidth, PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, Color, Color, Color, Color)

    End With

    Exit Sub

Error:

    If err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & vbCrLf & err.Description, vbExclamation, "[ " & err.Number & " ] Error"
        End

    End If

End Sub

Sub DrawGrh(ByRef Grh As Grh, _
            ByVal X As Integer, _
            ByVal Y As Integer, _
            ByVal Center As Byte, _
            ByVal Animate As Single, _
            Optional Shadow As Byte = 0, _
            Optional Color As Long = -1, _
            Optional ShadowAlpha As Single = 255)

    Dim CurrentGrhIndex As Integer
    
    On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * Animate * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1

                        If Grh.Loops = 0 Then
                            Grh.Started = 0

                        End If

                    Else
                        Grh.Started = 0

                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
                
        'Draw
        'Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        If Shadow = 1 Then
            ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
            Call Engine_Render_Rectangle(X, Y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
        ElseIf Shadow = 2 Then
            ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
            Call Engine_Render_Rectangle(X + 10, Y - 16, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)

        End If

        If Color = -1 Then Color = Iluminacion

        Call Engine_Render_Rectangle(X, Y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, Color, Color, Color, Color)

    End With

    Exit Sub

Error:

    If err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & vbCrLf & err.Description, vbExclamation, "[ " & err.Number & " ] Error"
        'End

    End If

End Sub

Public Sub RenderNiebla()

    If Zonas(ZonaActual).Niebla = 0 Then
        If nAlpha > 0 Then
            If GTCPres < (GetTickCount() And &H7FFFFFFF) - GTCInicial Then
                nAlpha = nAlpha - 1
                GTCPres = (GetTickCount() And &H7FFFFFFF)

            End If

        Else
            Exit Sub

        End If

    End If

    If nAlpha < Zonas(ZonaActual).Niebla Then
        If GTCPres < (GetTickCount() And &H7FFFFFFF) - GTCInicial Then
            nAlpha = nAlpha + IIf(nAlpha + 1 < Zonas(ZonaActual).Niebla, 1, Zonas(ZonaActual).Niebla - nAlpha)
            GTCPres = (GetTickCount() And &H7FFFFFFF)

        End If

    End If

    Dim Mueve As Single 'Niebla

    Dim T     As Single

    Dim Color As Long
   
    GTCPres = Abs((GetTickCount() And &H7FFFFFFF) - GTCInicial)
    T = (GTCPres - 4000) / 1000
    Mueve = (T * 20) Mod 512

    Color = D3DColorRGBA(Zonas(ZonaActual).NieblaR, Zonas(ZonaActual).NieblaG, Zonas(ZonaActual).NieblaB, nAlpha)

    Call Engine_Render_D3DXSprite(0, 0, 512 - Mueve, 512, Mueve, 0, Color, 14706, 0)
    Call Engine_Render_D3DXSprite(0, 512, 512 - Mueve, 256, Mueve, 0, Color, 14706, 0)

    Call Engine_Render_D3DXSprite(512 - Mueve, 0, 512, 512, 0, 0, Color, 14706, 0)
    Call Engine_Render_D3DXSprite(512 - Mueve, 512, 512, 256, 0, 0, Color, 14706, 0)

    Call Engine_Render_D3DXSprite(1024 - Mueve, 0, Mueve, 512, 0, 0, Color, 14706, 0)
    Call Engine_Render_D3DXSprite(1024 - Mueve, 512, Mueve, 256, 0, 0, Color, 14706, 0)

End Sub

Function GetBitmapDimensions(ByVal BmpFile As String, _
                             ByRef bmWidth As Long, _
                             ByRef bmHeight As Long)

    '*****************************************************************
    'Gets the dimensions of a bmp
    '*****************************************************************
    Dim BMHeader    As BITMAPFILEHEADER

    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight

End Function

Sub DrawGrhtoHdc(ByVal hDC As Long, _
                 ByVal GrhIndex As Integer, _
                 ByRef SourceRect As RECT, _
                 ByRef destRect As RECT)
    '*****************************************************************
    'Draws a Grh's portion to the given area of any Device Context
    '*****************************************************************
    'Call SurfaceDB.Surface(GrhData(GrhIndex).FileNum).BltToDC(hDC, SourceRect, destRect)
    Call TransparentBlt(hDC, 0, 0, 32, 32, Inventario.Grafico(GrhData(GrhIndex).FileNum), 0, 0, 32, 32, vbMagenta)

End Sub

Public Sub CargarTile(X As Long, Y As Long, ByRef DataMap() As Byte)

    Dim ByFlags As Byte

    Dim Rango   As Byte

    Dim I       As Integer

    Dim tmpInt  As Integer

    Dim Pos     As Long
On Error GoTo err
    Pos = MapInfo.offset + (X - 1) * 10 + (Y - 1) * MapInfo.Width * 10

    ByFlags = DataMap(Pos)
    ByFlags = ByFlags Xor ((X Mod 200) + 55)
    Pos = Pos + 1

    If ByFlags = 50 Then
        MapData(X, Y).Blocked = 1
    Else
        MapData(X, Y).Blocked = 0

    End If

    MapData(X, Y).Trigger = ByFlags

    For I = 1 To 4
        tmpInt = (DataMap(Pos + 1) And &H7F) * &H100 Or DataMap(Pos) Or -(DataMap(Pos + 1) > &H7F) * &H8000
        Pos = Pos + 2

        Select Case I

            Case 1
                MapData(X, Y).Graphic(1).GrhIndex = (tmpInt Xor (Y + 301) Xor (X + 721)) - X

            Case 2
                MapData(X, Y).Graphic(2).GrhIndex = (tmpInt Xor (Y + 501) Xor (X + 529)) - X

            Case 3
                MapData(X, Y).Graphic(3).GrhIndex = (tmpInt Xor (X + 239) Xor (Y + 319)) - X

            Case 4
                MapData(X, Y).Graphic(4).GrhIndex = (tmpInt Xor (X + 671) Xor (Y + 129)) - X

        End Select
    
        If MapData(X, Y).Graphic(I).GrhIndex > 0 Then
            InitGrh MapData(X, Y).Graphic(I), MapData(X, Y).Graphic(I).GrhIndex

        End If

    Next I

    'Get ArchivoMapa, , Rango
    Rango = DataMap(Pos)
    Pos = Pos + 1

    MapData(X, Y).Map = UserMap

    MapData(X, Y).light_value(0) = D3DColorRGBA(255, 255, 255, 255)
    MapData(X, Y).light_value(1) = D3DColorRGBA(255, 255, 255, 255)
    MapData(X, Y).light_value(2) = D3DColorRGBA(255, 255, 255, 255)
    MapData(X, Y).light_value(3) = D3DColorRGBA(255, 255, 255, 255)
    MapData(X, Y).Hora = 99

    Call Light_Destroy_ToMap(X, Y)

    If MapData(X, Y).Graphic(3).GrhIndex < 0 Then
        Call Light_Create(X, Y, 255, 255, 255, Rango, -MapData(X, Y).Graphic(3).GrhIndex - 1)

    End If
err:
End Sub

Sub RenderScreen(ByVal TileX As Integer, _
                 ByVal TileY As Integer, _
                 ByVal PixelOffSetX As Single, _
                 ByVal PixelOffSetY As Single)

'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y As Long    'Keeps track of where on map we are

    Dim X As Long    'Keeps track of where on map we are

    Dim screenminY As Integer    'Start Y pos on current screen

    Dim screenmaxY As Integer    'End Y pos on current screen

    Dim screenminX As Integer    'Start X pos on current screen

    Dim screenmaxX As Integer    'End X pos on current screen

    Dim MinY As Integer    'Start Y pos on current map

    Dim MaxY As Integer    'End Y pos on current map

    Dim MinX As Integer    'Start X pos on current map

    Dim MaxX As Integer    'End X pos on current map

    Dim ScreenX As Integer    'Keeps track of where to place tile on screen

    Dim ScreenY As Integer    'Keeps track of where to place tile on screen

    Dim minXOffset As Integer

    Dim minYOffset As Integer

    Dim PixelOffSetXTemp As Integer    'For centering grhs

    Dim PixelOffSetYTemp As Integer    'For centering grhs

    Dim tmpInt As Integer

    Dim tmpLong As Long

    Dim SupIndex As Integer

    Dim ByFlags As Byte

    Dim I As Integer

    Dim Color As Long

    Dim Eliminados As Integer

    Dim Cant As Integer

    If UserMap = 0 Then Exit Sub

    Dim BufferX1 As Integer

    Dim BufferX2 As Integer

    Dim BufferX3 As Integer

    Dim BufferX4 As Integer

    Dim BufferY1 As Integer

    Dim BufferY2 As Integer

    Dim BufferY3 As Integer

    Dim BufferY4 As Integer

    BufferX1 = HalfWindowTileWidth
    BufferY1 = HalfWindowTileHeight

    BufferX2 = HalfWindowTileWidth + 1
    BufferY2 = HalfWindowTileHeight + 1

    BufferX3 = HalfWindowTileWidth + 8
    BufferY3 = HalfWindowTileHeight + 8

    BufferX4 = HalfWindowTileWidth + 16
    BufferY4 = HalfWindowTileHeight + 16

    'Particulas
    '*********
    ParticleOffsetX = (Engine_PixelPosX(TileX - 17) - PixelOffSetX)
    ParticleOffsetY = (Engine_PixelPosY(TileY - 12) - PixelOffSetY)
    'Particulas
    '***************

    'Dim CambioHora As Boolean
    'Cargar mapa
    For Y = TileY - BufferY4 To TileY + BufferY4
        For X = TileX - BufferX4 To TileX + BufferX4

            If X > 0 And Y > 0 And X <= MapInfo.Width And Y <= MapInfo.Height Then
                If MapData(X, Y).Map <> UserMap Then
                    If UserMap = 1 Then
                        Call CargarTile(X, Y, DataMap1)
                    Else
                        Call CargarTile(X, Y, DataMap2)

                    End If

                End If

                'If MapData(X, Y).Hora <> Hora And UserMap = 1 Then

                '    For I = 0 To 3
                '        MapData(X, Y).Light_Value(I) = Iluminacion
                '    Next I

                '    MapData(X, Y).Hora = Hora

                'CambioHora = True
                'End If

            End If

        Next X
    Next Y

    Light_Render_Area

    'Draw floor layer
    For Y = TileY - BufferY2 To TileY + BufferY2
        For X = TileX - BufferX2 To TileX + BufferX2

            If X > 0 And Y > 0 And X <= MapInfo.Width And Y <= MapInfo.Height Then
                ScreenX = X - TileX + BufferX1
                ScreenY = Y - TileY + BufferY1
                'Layer 1 **********************************
                Call DrawGrhLuz(MapData(X, Y).Graphic(1), ScreenX * TilePixelWidth + PixelOffSetX, ScreenY * TilePixelHeight + PixelOffSetY, 0, 1, MapData(X, Y).light_value)

                '******************************************
            End If

        Next X

    Next Y




    For Y = TileY - BufferY3 - 5 To TileY + BufferY3 + 5
        For X = TileX - BufferX3 To TileX + BufferX3

            If X > 0 And Y > 0 And X <= MapInfo.Width And Y <= MapInfo.Height Then
                ScreenX = X - TileX + BufferX1
                ScreenY = Y - TileY + BufferY1

                'Layer 2 **********************************
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                    Call DrawGrhLuz(MapData(X, Y).Graphic(2), ScreenX * TilePixelWidth + PixelOffSetX, ScreenY * TilePixelHeight + PixelOffSetY, 1, 1, MapData(X, Y).light_value)

                End If

            End If

        Next X

    Next Y

    Dim mNPCMuerto As clsNPCMuerto

    Eliminados = 0
    Cant = NPCMuertos.count

    For I = 1 To Cant
        Set mNPCMuerto = NPCMuertos(I - Eliminados)
        Call mNPCMuerto.Update    '(TileX, TileY, PixelOffSetX, PixelOffSetY)

        If mNPCMuerto.KillMe Then
            NPCMuertos.Remove (I - Eliminados)
            Eliminados = Eliminados + 1

        End If

    Next I
    'vida abajo del pj
    If Vidarender = True Then
        Dim CantVidax As Integer
If UserMaxHP > 0 Then
        CantVidax = (((UserMinHP / 33) / (UserMaxHP / 33)) * 33)

        Call Engine_Render_Rectangle(496, 375, 16, CantVidax, 0, 0, 16, CantVidax, , , 0, 14810)
End If
    End If
    ' vida abajo del pj
    'mana abajo del pj
    If Manarender = True Then
        Dim CantManx As Integer

        If UserMaxMAN > 0 Then
            CantManx = (((UserMinMAN / 33) / (UserMaxMAN / 33)) * 33)

            Call Engine_Render_Rectangle(530, 408, -16, -CantManx, 0, 0, -16, -CantManx, , , 0, 14810)

        End If
    End If



    'sangre
    
    
                     Engine_Render_Blood
           

   

    'sangre
    'Draw Transparent Layers
    ScreenY = minYOffset

    For Y = TileY - BufferY4 To TileY + BufferY4
        For X = TileX - BufferX4 To TileX + BufferX4

            If X > 0 And Y > 0 And X <= MapInfo.Width And Y <= MapInfo.Height Then
                ScreenX = X - TileX + BufferX1
                ScreenY = Y - TileY + BufferY1

                PixelOffSetXTemp = ScreenX * TilePixelWidth + PixelOffSetX
                PixelOffSetYTemp = ScreenY * TilePixelHeight + PixelOffSetY

                With MapData(X, Y)

                    'Pasos
                    If .PasosIndex <> 0 Then Call vPasos.RenderPasos(PixelOffSetXTemp, PixelOffSetYTemp, .PasosIndex)

                    'Object Layer **********************************
                    If .ObjGrh.GrhIndex <> 0 Then
                        Call DrawGrhLuz(.ObjGrh, PixelOffSetXTemp, PixelOffSetYTemp, 1, 1, MapData(X, Y).light_value)

                    End If

                    '***********************************************

                    If Not .Elemento Is Nothing Then    'Render de Npc Muertos
                        Call .Elemento.Render(PixelOffSetXTemp, PixelOffSetYTemp)

                    End If

                    'Char layer ************************************
                    If .CharIndex <> 0 Then
                        Call CharRender(charlist(.CharIndex), .CharIndex, PixelOffSetXTemp, PixelOffSetYTemp)

                        If .CharIndex <> UserCharIndex And UserPos.X = charlist(.CharIndex).Pos.X And UserPos.Y = charlist(.CharIndex).Pos.Y Then
                            Debug.Print "ME PISO CHEEE ******************************************************************************"
                            'verr post de los bost
                            charlist(.CharIndex).Pos.X = charlist(.CharIndex).Pos.X + 1
                            charlist(.CharIndex).Pos.Y = charlist(.CharIndex).Pos.Y + 1

                            'ver post de los bots
                        End If

                    End If

                    If UserMap = 1 Then
                        Call RenderBarcos(X, Y, TileX, TileY, PixelOffSetX, PixelOffSetY)

                    End If

                    'Particulas
                    '****************
                    If .particle_index > 0 Then
                        Effect_Begin .particle_index, PixelOffSetXTemp, PixelOffSetYTemp, 9, 200

                    End If

                    'Particulas
                    '*************************************************

                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex > 0 Then
                        'Draw
                        SupIndex = GrhData(.Graphic(3).GrhIndex).FileNum

                        If ((SupIndex >= 7000 And SupIndex <= 7008) Or (SupIndex >= 1261 And SupIndex <= 1287) Or SupIndex = 648 Or SupIndex = 645) Then
                            If mOpciones.TransparencyTree = True And UserPos.X >= X - 3 And UserPos.X <= X + 3 And UserPos.Y >= Y - 5 And UserPos.Y <= Y Then
                                Call DrawGrh(.Graphic(3), PixelOffSetXTemp, PixelOffSetYTemp, 1, 1, 0, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 180))
                            Else
                                Call DrawGrhLuz(.Graphic(3), PixelOffSetXTemp, PixelOffSetYTemp, 1, 1, MapData(X, Y).light_value)

                            End If

                        Else
                            Call DrawGrh(.Graphic(3), PixelOffSetXTemp, PixelOffSetYTemp, 1, 1)

                        End If

                    End If

                    '*************************************************

                    'Layer 3 Plus FX *****************************************
                    If .fXGrh.Started = 1 Then
                        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

                        Call DrawGrh(.fXGrh, PixelOffSetXTemp - FxData(.fX).OffSetX, PixelOffSetYTemp - FxData(.fX).OffSetY, 1, 1)

                        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                        If .fXGrh.Started = 0 Then .fX = 0

                    End If

                    '************************************************

                End With

            End If

        Next X
    Next Y


    'particulas ORE

    For Y = TileY - BufferY4 To TileY + BufferY4
        For X = TileX - BufferX4 To TileX + BufferX4
            If X > 0 And Y > 0 And X <= MapInfo.Width And Y <= MapInfo.Height Then
                ScreenY = Y - TileY + BufferY1
                ScreenX = X - TileX + BufferX1

                With MapData(X, Y)
                    If .particle_group > 0 Then
                        ParticlesORE.Particle_Group_Render .particle_group, ScreenX * 32 + PixelOffSetX, ScreenY * 32 + PixelOffSetY
                    End If
                End With
            End If
        Next X

    Next Y
    'Particulas ORE

    Dim mArroja As clsArroja

    Dim Elemento

    For Each Elemento In Arrojas

        Set mArroja = Elemento
        Call mArroja.Render(TileX, TileY, PixelOffSetX, PixelOffSetY)
    Next Elemento

    'Particulas
    '**************

    Effect_UpdateAll


    'Clear the shift-related variables
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    'Particuas
    '****************


    Dim mTooltip As clsToolTip

    Eliminados = 0
    Cant = Tooltips.count

    For I = 1 To Cant
        Set mTooltip = Tooltips(I - Eliminados)
        #If RenderFull = 0 Then
            Call mTooltip.Render(TileX - 5, TileY - 3, PixelOffSetX, PixelOffSetY)
        #Else
            Call mTooltip.Render(TileX - 1, TileY, PixelOffSetX, PixelOffSetY)

        #End If

        If mTooltip.Alpha = 0 Then
            Tooltips.Remove (I - Eliminados)
            Eliminados = Eliminados + 1

        End If

    Next I



    If Not bTecho Then
        If bAlpha < 255 Then
            If tTick < (GetTickCount() And &H7FFFFFFF) - 30 Then
                bAlpha = bAlpha + IIf(bAlpha + 8 < 255, 8, 255 - bAlpha)
                ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, bAlpha)
                tTick = (GetTickCount() And &H7FFFFFFF)

            End If

        End If

    Else

        If bAlpha > 0 Then
            If tTick < (GetTickCount() And &H7FFFFFFF) - 30 Then
                bAlpha = bAlpha - IIf(bAlpha - 8 > 0, 8, bAlpha)
                ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, bAlpha)
                tTick = (GetTickCount() And &H7FFFFFFF)

            End If

        End If

    End If

    For Y = TileY - BufferY4 To TileY + BufferY4
        For X = TileX - BufferX4 To TileX + BufferX4

            If X > 0 And Y > 0 And X <= MapInfo.Width And Y <= MapInfo.Height Then
                ScreenX = X - TileX + BufferX1
                ScreenY = Y - TileY + BufferY1

                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex And bAlpha > 0 Then
                    'Draw
                    Call DrawGrhIndex(MapData(X, Y).Graphic(4).GrhIndex, ScreenX * TilePixelWidth + PixelOffSetX, ScreenY * TilePixelHeight + PixelOffSetY, 1, ColorTecho)

                End If

            End If

        Next X
    Next Y

    'TODO : Check this!!
    Dim ColorLluvia As Long

    If ZonaActual > 0 Then
        If Zonas(ZonaActual).Terreno <> eTerreno.Dungeon Then
            If bRain Then

                'Figure out what frame to draw
                If llTick < (GetTickCount() And &H7FFFFFFF) - 50 Then
                    iFrameIndex = iFrameIndex + 1
                    If iFrameIndex > 7 Then iFrameIndex = 0
                    llTick = (GetTickCount() And &H7FFFFFFF)

                End If

                ColorLluvia = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 140)

                'fix Lluvia idea SenSui, Helios 06/06/2021
                For Y = 0 To 6
                    For X = 0 To 7
                        ' Call Engine_Render_Rectangle(LTLluvia(x), LTLluvia(y) + 40, RLluvia(iFrameIndex).Right + 30 - RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Bottom - RLluvia(iFrameIndex).Top, RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Top, RLluvia(iFrameIndex).Right - RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Bottom - RLluvia(iFrameIndex).Top, , , , 5556, ColorLluvia, ColorLluvia, ColorLluvia, ColorLluvia)
                        Call Engine_Render_Rectangle(LTLluvia(X) - 256, LTLluvia(Y) + 40 - 256, RLluvia(iFrameIndex).right + 30 - RLluvia(iFrameIndex).left, RLluvia(iFrameIndex).bottom - RLluvia(iFrameIndex).top, RLluvia(iFrameIndex).left, RLluvia(iFrameIndex).top, RLluvia(iFrameIndex).right - RLluvia(iFrameIndex).left, RLluvia(iFrameIndex).bottom - RLluvia(iFrameIndex).top, , , , 5556, ColorLluvia, ColorLluvia, ColorLluvia, ColorLluvia)
                    Next X
                Next Y

            End If

        End If

    End If

    If ZonaActual = 62 Then
        NieveOn = True
    Else
        NieveOn = False

    End If

    'NIEBLA EN BOSQUE DORK BLANCA Y EN INFIERNO ROJA
    'If ZonaActual = 11 Or ZonaActual = 31 Then

    '   Dim Mueve As Single
    '  Dim tmpColor As Long
    ' Dim t As Single
    'GTCPres = Abs((GetTickCount() And &H7FFFFFFF) - GTCInicial)
    't = (GTCPres - 4000) / 1000
    'For t = 1 To 512
    ' Mueve = (t * 20) Mod 512
    ' Mueve = t
    'If ZonaActual = 11 Then

    '   tmpColor = D3DColorRGBA(253, 255, 255, CalcAlpha(GTCPres, 4000, 150, 15))
    'Else
    '   tmpColor = D3DColorRGBA(251, 83, 108, CalcAlpha(GTCPres, 4000, 150, 15))
    'End If
    'Call Engine_Render_D3DXSprite(255, 255, 512 - Mueve, 512, Mueve, 0, tmpColor, 14706, 0)
    'Call Engine_Render_D3DXSprite(255, 767, 512 - Mueve, 256, Mueve, 0, tmpColor, 14706, 0)

    'Call Engine_Render_D3DXSprite(767 - Mueve, 255, 512, 512, 0, 0, tmpColor, 14706, 0)
    'Call Engine_Render_D3DXSprite(767 - Mueve, 767, 512, 256, 0, 0, tmpColor, 14706, 0)

    '        Call Engine_Render_D3DXSprite(1279 - Mueve, 255, Mueve, 512, 0, 0, tmpColor, 14706, 0)
    '       Call Engine_Render_D3DXSprite(1279 - Mueve, 767, Mueve, 256, 0, 0, tmpColor, 14706, 0)
    'Next t

    '   End If
    '

    Call Dialogos.Render
    Call DibujarCartel
    Call DialogosClanes.Draw

    If CambioZona > 0 And ZonaActual > 0 Then
        If CambioZona > 300 Then
            tmpInt = 500 - CambioZona
        ElseIf CambioZona < 200 Then
            tmpInt = CambioZona
        Else
            tmpInt = 200

        End If

        If zTick < (GetTickCount() And &H7FFFFFFF) - 50 Then
            CambioZona = CambioZona - 5
            zTick = (GetTickCount() And &H7FFFFFFF)

        End If

        'Mensaje al cambiar de zona
        #If RenderFull = 0 Then
            If ZonaActual <> 23 Then
                Call D3DX.DrawText(MainFont, D3DColorRGBA(0, 0, 0, tmpInt), Zonas(ZonaActual).nombre, DDRect(0, 140, 1024, 220), DT_CENTER)
                Call D3DX.DrawText(MainFont, D3DColorRGBA(220, 215, 215, tmpInt), Zonas(ZonaActual).nombre, DDRect(0, 145, 1024, 220), DT_CENTER)
            End If
            If CambioSegura Then
                Call DrawFont(IIf(Zonas(ZonaActual).Segura = 1, "Entraste a una zona segura", "Saliste de una zona segura"), 420, 214, D3DColorRGBA(255, 0, 0, tmpInt))

            End If

        #Else
            Call D3DX.DrawText(MainFont, D3DColorRGBA(0, 0, 0, tmpInt), Zonas(ZonaActual).nombre, DDRect(0, 10, 814, 220), DT_CENTER)
            Call D3DX.DrawText(MainFont, D3DColorRGBA(220, 215, 215, tmpInt), Zonas(ZonaActual).nombre, DDRect(5, 15, 814, 220), DT_CENTER)

            If CambioSegura Then
                Call DrawFont(IIf(Zonas(ZonaActual).Segura = 1, "Entraste a una zona segura", "Saliste de una zona segura"), 318, 89, D3DColorRGBA(255, 0, 0, tmpInt))

            End If

        #End If

    End If
    
    
   
    If UseMotionBlur And mOpciones.BlurEffects = True Then

        AngMareoMuerto = AngMareoMuerto + timerElapsedTime * 0.002

        If AngMareoMuerto >= 6.28318530717959 Then
            AngMareoMuerto = 0

            'GoingHome = 0
        End If

        If GoingHome = 1 Then
            RadioMareoMuerto = RadioMareoMuerto + timerElapsedTime * 0.01

            If RadioMareoMuerto > 50 Then RadioMareoMuerto = 50
        ElseIf GoingHome = 2 Then
            RadioMareoMuerto = RadioMareoMuerto - timerElapsedTime * 0.02

            If RadioMareoMuerto <= 0 Then
                RadioMareoMuerto = 0
                GoingHome = 0

            End If

        End If

        If FrameUseMotionBlur Then
            FrameUseMotionBlur = False

            With D3DDevice

                'Dim ValueEffect As Long
                'ValueEffect = 2048

                'Perform the zooming calculations
                ' * 1.333... maintains the aspect ratio
                ' ... / 1024 is to factor in the buffer size
                BlurTA(0).tu = ZoomLevel + RadioMareoMuerto / 2048 * Sin(AngMareoMuerto) + RadioMareoMuerto / 2048
                BlurTA(0).tv = ZoomLevel + RadioMareoMuerto / 2048 * Cos(AngMareoMuerto) + RadioMareoMuerto / 2048
                BlurTA(1).tu = ((ScreenWidth + 1 + Cos(AngMareoMuerto) * RadioMareoMuerto / 2 - RadioMareoMuerto / 2) / 1024) - ZoomLevel
                BlurTA(1).tv = ZoomLevel + RadioMareoMuerto / 2048 * Sin(AngMareoMuerto) + RadioMareoMuerto / 2048
                BlurTA(2).tu = ZoomLevel + RadioMareoMuerto / 2048 * Cos(AngMareoMuerto) + RadioMareoMuerto / 2048
                BlurTA(2).tv = ((ScreenHeight + 1 + Sin(AngMareoMuerto) * RadioMareoMuerto / 2 - RadioMareoMuerto / 2) / 768) - ZoomLevel
                BlurTA(3).tu = BlurTA(1).tu
                BlurTA(3).tv = BlurTA(2).tv

                'Draw what we have drawn thus far since the last .Clear
                'LastTexture = -100
                D3DDevice.EndScene
                .SetRenderTarget pBackbuffer, Nothing, ByVal 0

                D3DDevice.BeginScene

                .SetTexture 0, BlurTexture
                .SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(BlurIntensity, 255, 255, 255)
                .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
                .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BlurTA(0), Len(BlurTA(0))
                .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE

            End With

        End If

    End If

    If VerMapa Then

        '420
        '0.46545454545454545454545454545455

        If UserMap = 1 Then
            PosMapX = -Int(UserPos.X * RelacionMiniMapa) + 32 + 398
            PosMapY = -Int(UserPos.Y * RelacionMiniMapa) + 32 + 292

            If PosMapX > 0 Then PosMapX = 0
            If PosMapX < -1247 Then PosMapX = -1247
            If PosMapY > 0 Then PosMapY = 0
            If PosMapY < -2210 Then PosMapY = -2210

            Color = D3DColorRGBA(255, 255, 255, 225)

            If PosMapX > -1024 Then    'Dibujo primera columna
                If PosMapY <= 0 And PosMapY > -1024 Then
                    Call Engine_Render_Rectangle(0, 0, 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY < -480, PosMapY + 480, 0), -PosMapX, -PosMapY, 600 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY < -480, PosMapY + 480, 0), , , , 14763, Color, Color, Color, Color)

                End If

                If PosMapY <= -480 And PosMapY > -2048 Then
                    Call Engine_Render_Rectangle(0, PosMapY + 1024, 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, -PosMapX, 0, 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, , , , 14765, Color, Color, Color, Color)

                End If

                If PosMapY <= -1504 Then
                    Call Engine_Render_Rectangle(0, PosMapY + 2048, 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), -PosMapX, 0, 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), , , , 14767, Color, Color, Color, Color)

                End If

            End If

            If PosMapX < -288 Then    'Dibujo segunda columna
                If PosMapY <= 0 And PosMapY > -1024 Then
                    Call Engine_Render_Rectangle(IIf(PosMapX > -1024, 736 + 288 + PosMapX, 0), 0, 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY < -480, PosMapY + 480, 0), IIf(PosMapX < -1024, -PosMapX - 1024, 0), -PosMapY, 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY < -480, PosMapY + 480, 0), , , , 14764, Color, Color, Color, Color)

                End If

                If PosMapY <= -480 And PosMapY > -2048 Then
                    Call Engine_Render_Rectangle(IIf(PosMapX > -1024, 736 + 288 + PosMapX, 0), PosMapY + 1024, 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, IIf(PosMapX < -1024, -PosMapX - 1024, 0), 0, 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, , , , 14766, Color, Color, Color, Color)

                End If

                If PosMapY <= -1504 Then
                    Call Engine_Render_Rectangle(IIf(PosMapX > -1024, 736 + 288 + PosMapX, 0), PosMapY + 2048, 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), IIf(PosMapX < -1024, -PosMapX - 1024, 0), 0, 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), , , , 14768, Color, Color, Color, Color)

                End If

            End If

            'If PosMap <= 210 Then '488
            '    MapaY = 0
            'ElseIf PosMap > 210 And PosMap < 492 Then
            '    MapaY = -(PosMap - 210)
            'Else
            '    MapaY = -282
            'End If

            'Call Engine_Render_Rectangle(256 + 0, 256 + MapaY, 512, 512, 0, 0, 512, 512, , , , 14404, color, color, color, color)
            'Call Engine_Render_Rectangle(256 + 0, 256 + 512 + MapaY, 512, 186, 0, 0, 512, 186, , , , 14405, color, color, color, color)
            Color = D3DColorRGBA(255, 255, 255, 255)
            'Call Engine_Render_Rectangle(256 + UserPos.x * RelacionMiniMapa - 35 + PosMapX, 256 + UserPos.Y * RelacionMiniMapa - 35 + PosMapY, 5, 5, 0, 0, 5, 5, , , , 1, color, color, color, color)
            Call Engine_Render_Rectangle(UserPos.X * RelacionMiniMapa - 35 + PosMapX, UserPos.Y * RelacionMiniMapa - 35 + PosMapY, 5, 5, 0, 0, 5, 5, , , , 1, Color, Color, Color, Color)

            X = Int((frmMain.MouseX - PosMapX + 32) / RelacionMiniMapa)
            Y = Int((frmMain.MouseY - PosMapY + 32) / RelacionMiniMapa)

            If X > 1 And X < 1100 And Y > 1 And Y < 1500 Then
                Call DrawFont("(" & X & "," & Y & ")", frmMain.MouseX + 12, frmMain.MouseY + 12, D3DColorRGBA(255, 255, 255, 200))
                I = BuscarZona(X, Y)

                If I > 0 Then
                    Call DrawFont(Zonas(I).nombre, frmMain.MouseX - 10, frmMain.MouseY + 26, D3DColorRGBA(255, 255, 255, 200))

                End If

            End If

        ElseIf ZonaActual = 33 Or ZonaActual = 34 Or ZonaActual = 35 Then    'Dungeon Newbie
            Color = D3DColorRGBA(255, 255, 255, 190)
            Call Engine_Render_Rectangle(60, 3, 512, 512, 0, 0, 512, 512, , , , 14406, Color, Color, Color, Color)

            Color = D3DColorRGBA(255, 255, 255, 255)
            Call Engine_Render_Rectangle(60 + (UserPos.X - 571) * 2.21105527638191, 5 + (UserPos.Y - 311) * 2.21105527638191, 5, 5, 0, 0, 5, 5, , , , 1, Color, Color, Color, Color)
        Else
            'Mensaje al cambiar de zona
            Call D3DX.DrawText(MainFont, D3DColorRGBA(0, 0, 0, 200), Zonas(ZonaActual).nombre, DDRect(0, 10, 736, 200), DT_CENTER)
            Call D3DX.DrawText(MainFont, D3DColorRGBA(220, 215, 215, 200), Zonas(ZonaActual).nombre, DDRect(5, 15, 736, 200), DT_CENTER)

        End If

    End If

    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Particle_Group_Render MapData(150, 800).particle_group_index, MouseX, MouseY

    'Dim tmplng As Long
    'Dim tmblng2 As Long
    'ScreenY = minYOffset '- TileBufferSize
    'For y = minY To maxY
    '    ScreenX = minXOffset '- TileBufferSize
    '    For x = minX To maxX
    '        With MapData(x, y)
    '            '*** Start particle effects ***
    '            If MapData(x, y).particle_group_index Then
    '                Particle_Group_Render MapData(x, y).particle_group_index, ScreenX, ScreenY
    '            End If
    '            '*** End particle effects ***
    '        End With
    '        ScreenX = ScreenX + 1
    '    Next x
    '    ScreenY = ScreenY + 1
    'Next y
    'Call Engine_Render_Rectangle(frmMain.MouseX, frmMain.MouseY, 128, 128, 0, 256, 128, 128, , , 0, 14332)

    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    If TiempoRetos > 0 Then
        '10 segundos de espera para empezar la ronda
        tmpLong = Abs((GetTickCount() And &H7FFFFFFF) - TiempoRetos)
        tmpInt = 10 - Int(tmpLong / 1000)

        If tmpLong < 10000 Then
            Call D3DX.DrawText(MainFontBig, D3DColorRGBA(0, 0, 0, 200), CStr(tmpInt), DDRect(0, 30, 736, 230), DT_CENTER)
            Call D3DX.DrawText(MainFontBig, D3DColorRGBA(220, 215, 215, 200), CStr(tmpInt), DDRect(5, 35, 736, 230), DT_CENTER)
        Else
            'Termino el tiempo de espera que empieze el reto
            TiempoRetos = 0

        End If

    End If

    If Entrada > 0 Then
        Color = D3DColorRGBA(255, 255, 255, Entrada)
        Call Engine_Render_Rectangle(0, 0, 544, 416, 0, 0, 512, 416, , , , 14325, Color, Color, Color, Color)

        If zTick2 < (GetTickCount() And &H7FFFFFFF) - 75 Then
            Entrada = Entrada - 15
            zTick2 = (GetTickCount() And &H7FFFFFFF)

        End If

    End If

    If PergaminoDireccion > 0 Then
        If PergaminoTick < (GetTickCount() And &H7FFFFFFF) - 20 Then
            If PergaminoDireccion = 1 And AperturaPergamino < 240 Then
                AperturaPergamino = AperturaPergamino + (5 + Sqr(240 - AperturaPergamino) / 2)

                If AperturaPergamino > 240 Then AperturaPergamino = 240
            ElseIf PergaminoDireccion = 2 And AperturaPergamino > 0 Then
                AperturaPergamino = AperturaPergamino - (5 + Sqr(AperturaPergamino) / 2)

                If AperturaPergamino < 0 Then AperturaPergamino = 0

            End If

            PergaminoTick = (GetTickCount() And &H7FFFFFFF)

        End If

    End If

    If AperturaPergamino > 0 Then
        If DateDiff("s", TiempoAbierto, Now) > 10 Then
            PergaminoDireccion = 2
            TiempoAbierto = Now

        End If

        Color = D3DColorRGBA(255, 255, 255, AperturaPergamino * 175 / 240)

        Call Engine_Render_Rectangle(10 - 5 + 240 - AperturaPergamino, 309 + 2, 28, 107, 0, 0, 28, 107, , , , 14687, Color, Color, Color, Color)
        Call Engine_Render_Rectangle(38 - 5 + 240 - AperturaPergamino, 336 + 2, AperturaPergamino, 74, 240 - AperturaPergamino, 108, AperturaPergamino, 74, , , , 14687, Color, Color, Color, Color)
        Call Engine_Render_Rectangle(517 - 5 - 240 + AperturaPergamino, 309 + 2, 26, 107, 29, 0, 26, 107, , , , 14687, Color, Color, Color, Color)
        Call Engine_Render_Rectangle(278 - 5, 335 + 2, AperturaPergamino, 74, 0, 182, AperturaPergamino, 74, , , , 14687, Color, Color, Color, Color)

        'If AperturaPergamino >= 232 Then
        '    Call Engine_Render_Rectangle(256 + 40, 256 + 335 + 11, 56, 56, 56, 0, 56, 56, , , , 14687, color, color, color, color)
        'ElseIf AperturaPergamino < 232 And AperturaPergamino >= 176 Then
        '    Call Engine_Render_Rectangle(256 + 40 + 232 - AperturaPergamino, 256 + 335 + 11, AperturaPergamino - 176, 56, 56 + 232 - AperturaPergamino, 0, AperturaPergamino - 176, 56, , , , 14687, color, color, color, color)
        'End If
        Call Engine_Render_D3DXTexture(38 - 5 + 240 - Int(AperturaPergamino), 342, Int(AperturaPergamino) * 2, 80, 240 - Int(AperturaPergamino), 0, Color, pRenderTexture, 0)

    End If

    If mOpciones.Niebla = True Then

        Call RenderNiebla

    End If

    If AlphaCuenta > 0 Then
        Call RenderCuentaRegresiva

    End If

    If AlphaBlood > 0 Then
        Call RenderBlood

    End If

    If AlphaBloodUserDie > 0 Then
        Call RenderUserDieBlood

    End If

    If AlphaTextKills > 0 Then
        Call RenderTextKills

    End If

    If bRain = True Then
        Call RenderRelampago

    End If

    If UserCiego = True Then
        Call RenderCeguera

    End If

    If AlphaSalir > 0 Then
        Call RenderSaliendo

    End If

    FrameTime = (timeGetTime() And &H7FFFFFFF)

    #If RenderFull = 0 Then

        If FPSFLAG Then Call DrawFont("FPS: ", 484, 34, D3DColorRGBA(240, 34, 37, 200))
        If FPSFLAG Then Call DrawFont("        " & FPS, 484, 34, D3DColorRGBA(101, 209, 238, 160))
        If Resolucion = True Then
            Call Engine_Render_Rectangle(0, 0, 1024, 768, 0, 0, 1024, 782, , , 0, 14941)
        Else
            Call Engine_Render_Rectangle(1, 1, 1024, 782, 0, 0, 1024, 782, , , 0, 14941)
        End If

        If MmenuBarras = True Then
            'Call Engine_Render_Rectangle(13, 0, 250, 250, 0, 0, 250, 250, , , 0, 14936) ' vida Helios
            'Call DrawFont(CStr(UserLvl), 57, 39, D3DColorRGBA(255, 255, 0, 190)) 'Helios UserNivel
            'Call DrawFont(CStr(UserName), 54, 88, D3DColorRGBA(255, 255, 255, 255)) ' user Name Helios
            'Call DrawFont(" " & CStr(UserGLD), 174, 88, D3DColorRGBA(256, 239, 239, 160), True) 'ORO Helios




            Call Engine_Render_Rectangle(2, -10, 400, 300, 0, 0, 400, 300, , , 0, 14936)


            Dim CantAgua As Integer
            Dim CantHam As Integer
            Dim Cantvida As Integer
            Dim CantMana As Integer
            Dim CantStamina As Integer
            Dim CantExp As Integer
            Dim CantExp2 As Integer
            Dim CantAgilidad As Integer
            Dim CantFuerza As Integer
            'experiencia
            If UserExp <> 0 And UserPasarNivel <> 0 Then
            CantExp = 67 * Round(CDbl(UserExp) / CDbl(UserPasarNivel / 2), 2)

            'Call Engine_Render_Rectangle(309, 275, 91, 89, 0, 0, 91, 89, , , 0, 14801)
            If CantExp >= 67 Then
                CantExp2 = CantExp - 67
                CantExp = 67



            End If
End If
            Call Engine_Render_Rectangle(96, 87, -33, -CantExp2, 0, 0, -33, -CantExp2, , , 0, 14948)
            Call Engine_Render_Rectangle(31, 20, 33, CantExp, 0, 0, 33, CantExp, , , 0, 14948)
 If UserExp <> 0 And UserPasarNivel <> 0 Then
            Call DrawFont(Round((UserExp / UserPasarNivel) * 100) & "%", 68, 60, D3DColorRGBA(255, 255, 255, 200), True)
           Else
            Call DrawFont("0%", 68, 60, D3DColorRGBA(255, 255, 255, 200), True)
           
           End If
            Call D3DX.DrawText(FontRender, D3DColorRGBA(255, 255, 255, 200), CStr(UserLvl), DDRect(0, 30, 130, 0), DT_CENTER)

            'experiencia

            'vida
            Cantvida = (((UserMinHP / 211) / (UserMaxHP / 211)) * 211)

            Call Engine_Render_Rectangle(105, 13, Cantvida, 33, 0, 0, Cantvida, 33, , , 0, 14945)
            Call DrawFont(UserMinHP & " / " & UserMaxHP, 215, 22, D3DColorRGBA(255, 255, 255, 200), True)
            'vida

            'mana
            If UserMaxMAN > 0 Then
                CantMana = (((UserMinMAN / 211) / (UserMaxMAN / 211)) * 211)
            End If
            Call Engine_Render_Rectangle(105, 43, CantMana, 33, 0, 0, CantMana, 33, , , 0, 14946)
            Call DrawFont(UserMinMAN & " / " & UserMaxMAN, 215, 52, D3DColorRGBA(255, 255, 255, 200), True)
            'mana

            'agilidad
            CantAgilidad = -(((UserAtributos(2) / 42) / (42 / 42)) * 42)

            Call Engine_Render_Rectangle(227, 136, 45, CantAgilidad, 0, 0, 45, CantAgilidad, , , 0, 14942)
            Call DrawFont(Round(UserAtributos(2)), 248, 108, D3DColorRGBA(255, 255, 255, 160), True)

            'agilidad
            'fuerza
            CantFuerza = -(((UserAtributos(1) / 42) / (42 / 42)) * 42)

            Call Engine_Render_Rectangle(268, 136, 45, CantFuerza, 0, 0, 45, CantFuerza, , , 0, 14807)
            Call DrawFont(Round(UserAtributos(1)), 289, 108, D3DColorRGBA(255, 255, 255, 160), True)
            'fuerza

            'stamina
            CantStamina = (((UserMinSTA / 210) / (UserMaxSTA / 210)) * 210)

            Call Engine_Render_Rectangle(105, 74, CantStamina, 21, 0, 0, CantStamina, 21, , , 0, 14947)
            Call DrawFont(UserMinSTA & " / " & UserMaxSTA, 215, 77, D3DColorRGBA(255, 255, 255, 200), True)
            'stamina

            'agua
            CantAgua = -(((UserMinAGU / 41) / (UserMaxAGU / 41)) * 41)

            Call Engine_Render_Rectangle(145, 136, 45, CantAgua, 0, 0, 45, CantAgua, , , 0, 14943)
            Call DrawFont(Round((-CantAgua * 100) / 41), 167, 108, D3DColorRGBA(255, 255, 255, 160), True)
            'agua


            'comida
            CantHam = -(((UserMinHAM / 41) / (UserMaxHAM / 41)) * 41)

            Call Engine_Render_Rectangle(187, 136, 45, CantHam, 0, 0, 45, CantHam, , , 0, 14944)
            Call DrawFont(Round((-CantHam * 100) / 41), 208, 108, D3DColorRGBA(255, 255, 255, 160), True)
            'comida
            Call Engine_Render_Rectangle(10, 115, 32, 32, 0, 0, 32, 32, , , 0, 510)
            If UserGLD <> 0 Then
                Call DrawFont(Format$(UserGLD, "##,##"), 70, 123, D3DColorRGBA(255, 255, 0, 160), True)
            Else
                Call DrawFont(Round(UserGLD), 55, 123, D3DColorRGBA(255, 255, 0, 160), True)
            End If
            Call DrawFont(UserName, 65, 102, D3DColorRGBA(255, 255, 0, 160), True)
        End If

        Call Engine_Render_Rectangle(550, 0, 76, 35, 0, 0, 76, 35, , , 0, 14954)    'user onlie Helios

        Call Engine_Render_Rectangle(652, 0, 343, 36, 0, 0, 343, 36, , , 0, 14955)    'Menu Helios
        Call Engine_Render_Rectangle(627, 0, 369, 35, 0, 0, 369, 35, , , 0, 14809)    'Menu Helios
        Call Engine_Render_Rectangle(992, -5, 34, 35, 0, 0, 34, 35, , , 0, 14953)

        If frmMain.imgMiniMapa.Visible = True Then
            Call Engine_Render_Rectangle(892, 34, 110, 111, 0, 0, 110, 111, , , 0, 14958)    'Marco Minimapa

        End If

        If frmMain.invHechisos.Visible = True Then
            Call Engine_Render_Rectangle(894, 364, 40, 45, 0, 0, 40, 45, , , 0, 14956)    'Lanzar Hechizos

        End If

        If MostrarMenuInventario = True Then
            'Call Engine_Render_Rectangle(927, 209, 75, 36, 0, 0, 75, 36, , , 0, 14957) ' Helios menu elegir inventario hechizos
            Call Engine_Render_Rectangle(963, 209, 52, 408, 0, 0, 52, 408, , , 0, 14959)
            If RecuadroInv = True Then
                Call Engine_Render_Rectangle(RecuadroX, RecuadroY, 16, 14, 0, 0, 16, 14, , , 0, 14815)
            End If
        End If

        'Call Engine_Render_Rectangle(262, 690, 50, 39, 0, 0, 50, 39, , , 0, 14942) ' engranaje

        Call DrawFont("      " & CStr(UsersOn), 560, 14, D3DColorRGBA(240, 34, 37, 200))    'Useron Helios

        ' Call Engine_Render_Rectangle(1150, 450, 32, 32, 0, 0, 32, 32, , , 0, 510) ' ORO Helios

        Call DrawFont("      " & CStr(Time), 901, 14, D3DColorRGBA(101, 209, 238, 160))    'Hora Helios

        If frmMain.imgMiniMapa.Visible = True Then    'Helios escondo letras 04/06/2021 0:08
            If Zonas(ZonaActual).Segura = 1 Then
                Call DrawFont(Zonas(ZonaActual).nombre, 938, 143, D3DColorRGBA(0, 255, 0, 160), True)
            Else
                Call DrawFont(Zonas(ZonaActual).nombre, 938, 143, D3DColorRGBA(255, 0, 0, 160), True)

            End If

            ' Call DrawFont("Mapa: " & Zonas(ZonaActual).Mapa & "(X:" & UserPos.X & ", Y:" & UserPos.Y & ")", 1124, 425, D3DColorRGBA(255, 255, 255, 160))
            Call DrawFont("(X:" & UserPos.X & ", Y:" & UserPos.Y & ")", 938, 159, D3DColorRGBA(255, 255, 255, 160), True)

        End If
        If RecuadroON = True Then
            Call Engine_Render_Rectangle(RecuadroX, RecuadroY, 26, 23, 0, 0, 26, 23, , , 0, 14806)
        End If
        If RecuadroSON = True Then
            Call Engine_Render_Rectangle(RecuadroX, RecuadroY, 23, 21, 0, 0, 23, 21, , , 0, 14808)
        End If

        If Resolucion = True Then
            Call Engine_Render_Rectangle(950, 725, 27, 28, 0, 0, 27, 28, , , 0, SeguroResu)
            Call Engine_Render_Rectangle(980, 725, 27, 28, 0, 0, 27, 28, , , 0, SeguroConIma)
            If Consolacom = 0 Then
            Consolacom = 14816
            End If
            Call Engine_Render_Rectangle(20, 725, 27, 28, 0, 0, 27, 28, , , 0, Consolacom)
        

        Else
            Call Engine_Render_Rectangle(950, 738, 27, 28, 0, 0, 27, 28, , , 0, SeguroResu)
            Call Engine_Render_Rectangle(980, 738, 27, 28, 0, 0, 27, 28, , , 0, SeguroConIma)
            If Consolacom = 0 Then
            Consolacom = 14816
            End If
            Call Engine_Render_Rectangle(20, 738, 27, 28, 0, 0, 27, 28, , , 0, Consolacom)
        End If
        
        
        If frmMain.macrotrabajo Then
         
        Call DrawFont("¡¡Trabajando...!!", 895, 745, D3DColorRGBA(0, 255, 0, 160), True)
        End If
    #Else

        If FPSFLAG Then Call DrawFont("     " & FPS, 484, 34, D3DColorRGBA(101, 209, 238, 160))
    #End If

End Sub

'End Sub

Function CalcAlpha(Tiempo As Long, _
                   STiempo As Long, _
                   MaxAlpha As Byte, _
                   Tempo As Single) As Byte

    Dim tmpInt As Long

    tmpInt = (Tiempo - STiempo) / Tempo

    If tmpInt >= 0 Then
        CalcAlpha = IIf(tmpInt > MaxAlpha, MaxAlpha, tmpInt)

    End If

End Function

Public Function RenderSounds()

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 3/30/2008
    'Actualiza todos los sonidos del mapa.
    '**************************************************************
    If ZonaActual > 0 Then
        If Zonas(ZonaActual).Terreno <> eTerreno.Dungeon Then
            If bRain Then
                If bTecho Then
                    If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                        If RainBufferIndex Then Call Audio.StopWave(RainBufferIndex)
                        RainBufferIndex = Audio.PlayWave(SND_LLUVIAIN, 0, 0, LoopStyle.Enabled)
                        frmMain.IsPlaying = PlayLoop.plLluviain

                    End If

                Else

                    If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                        If RainBufferIndex Then Call Audio.StopWave(RainBufferIndex)
                        RainBufferIndex = Audio.PlayWave(SND_LLUVIAOUT, 0, 0, LoopStyle.Enabled)
                        frmMain.IsPlaying = PlayLoop.plLluviaout

                    End If

                End If

            End If

        Else

            If frmMain.IsPlaying <> PlayLoop.plNone Then
                Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone

            End If
        
        End If
    
        Call ReproducirSonidosDeAmbiente
        DoFogataFx
    
    End If

End Function

Function HayUserAbajo(ByVal X As Integer, _
                      ByVal Y As Integer, _
                      ByVal GrhIndex As Integer) As Boolean

    If GrhIndex > 0 Then
        HayUserAbajo = charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) And charlist(UserCharIndex).Pos.Y <= Y

    End If

End Function

Sub LoadGraphics()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero - complete rewrite
    'Last Modify Date: 11/03/2006
    'Initializes the SurfaceDB and sets up the rain rects
    '**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(D3DDevice, D3DX, ClientSetup.bUseVideo, DirRecursos & "Graphics.AO", ClientSetup.byMemory)
    
    'Set up te rain rects
    RLluvia(0).top = 0:      RLluvia(1).top = 0:      RLluvia(2).top = 0:      RLluvia(3).top = 0
    RLluvia(0).left = 0:     RLluvia(1).left = 128:   RLluvia(2).left = 256:   RLluvia(3).left = 384
    RLluvia(0).right = 128:  RLluvia(1).right = 256:  RLluvia(2).right = 384:  RLluvia(3).right = 512
    RLluvia(0).bottom = 128: RLluvia(1).bottom = 128: RLluvia(2).bottom = 128: RLluvia(3).bottom = 128
    
    RLluvia(4).top = 128:    RLluvia(5).top = 128:    RLluvia(6).top = 128:    RLluvia(7).top = 128
    RLluvia(4).left = 0:     RLluvia(5).left = 128:   RLluvia(6).left = 256:   RLluvia(7).left = 384
    RLluvia(4).right = 128:  RLluvia(5).right = 256:  RLluvia(6).right = 384:  RLluvia(7).right = 512
    RLluvia(4).bottom = 256: RLluvia(5).bottom = 256: RLluvia(6).bottom = 256: RLluvia(7).bottom = 256

End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, _
                               ByVal setMainViewTop As Integer, _
                               ByVal setMainViewLeft As Integer, _
                               ByVal setTilePixelHeight As Integer, _
                               ByVal setTilePixelWidth As Integer, _
                               ByVal setWindowTileHeight As Integer, _
                               ByVal setWindowTileWidth As Integer, _
                               ByVal setTileBufferSize As Integer, _
                               ByVal pixelsToScrollPerFrameX As Integer, _
                               pixelsToScrollPerFrameY As Integer, _
                               ByVal EngineSpeed As Single) As Boolean
    '***************************************************
    'Author: Aaron Perkins
    'Last Modification: 08/14/07
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Creates all DX objects and configures the engine to start running.
    '***************************************************
    
    IniPath = PathInit & "\"
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = Round(setWindowTileHeight \ 2, 0)
    HalfWindowTileWidth = Round(setWindowTileWidth \ 2, 0)
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = EngineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    'MinXBorder = 1 + (WindowTileWidth \ 2)
    'MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    'MinYBorder = 1 + (WindowTileHeight \ 2)
    'MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    'ReDim MapData(1 To XMaxMapSize, 1 To YMaxMapSize, 1 To 2) As MapBlock
    
    'Set intial user position
    UserPos.X = 1
    UserPos.Y = 1
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the view rect
    With MainViewRect
        .left = MainViewLeft
        .top = MainViewTop
        .right = .left + MainViewWidth
        .bottom = .top + MainViewHeight

    End With
    
    'Set the dest rect
    With MainDestRect
        .left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .right = .left + MainViewWidth
        .bottom = .top + MainViewHeight

    End With
    
    IniciarD3D

    Call CargarFont
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call CargarParticulas
    Call CargarAlas
    'Call General_Particle_Create(1, 150, 800, -1, 20, -15)
    
    Set TestPart = New clsParticulas
    TestPart.texture = 14386
    TestPart.ParticleCounts = 35
    TestPart.ReLocate 400, 400
    TestPart.Begin
    
    'Actual = 2
    'Particle_Group_Make Actual, 1, 150, 850, Particula(Actual).VarZ, Particula(Actual).VarX, Particula(Actual).VarY, Particula(Actual).AlphaInicial, Particula(Actual).RedInicial, Particula(Actual).GreenInicial, _
    'Particula(Actual).BlueInicial, Particula(Actual).AlphaFinal, Particula(Actual).RedFinal, Particula(Actual).GreenFinal, Particula(Actual).BlueFinal, Particula(Actual).NumOfParticles, Particula(Actual).gravity, Particula(Actual).Texture, Particula(Actual).Zize, Particula(Actual).Life
    
    LTLluvia(0) = 224
    LTLluvia(1) = 352
    LTLluvia(2) = 480
    LTLluvia(3) = 608
    LTLluvia(4) = 736
    LTLluvia(5) = 864
    LTLluvia(6) = 992
    LTLluvia(7) = 1120
    Call LoadGraphics
    'Particulas
    '**********
    Call Engine_Init_ParticleEngine
    'Particulas
    InitTileEngine = True

End Function

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, _
                  ByVal DisplayFormLeft As Integer, _
                  ByVal MouseViewX As Integer, _
                  ByVal MouseViewY As Integer, _
                  Optional ByVal Update As Boolean = False, _
                  Optional ByVal X As Integer = 0, _
                  Optional ByVal Y As Integer = 0)

'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single

    Static OffsetCounterY As Single

    '****** Set main view rectangle ******
    MainViewRect.left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
    MainViewRect.top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
    MainViewRect.right = MainViewRect.left + MainViewWidth
    MainViewRect.bottom = MainViewRect.top + MainViewHeight

    If EngineRun Then

        If Update = True Then
            ' D3DDevice.BeginScene
            Debug.Print " UPDATE SCREEEN"
            'Call UpdateRenderScreen(X, Y, OffsetCounterX, OffsetCounterY)
            'D3DDevice.EndScene
            Exit Sub

        End If

        If UserEmbarcado Then
            OffsetCounterX = -BarcoOffSetX
            OffsetCounterY = -BarcoOffSetY
        ElseIf UserMoving Then

            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame * 1.2

                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False

                End If

            End If

            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame * 1.2

                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False

                End If

            End If

        End If

        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)

        If GoingHome = 1 Or GoingHome = 2 Then
            BlurIntensity = 5
        Else
            BlurIntensity = 0

        End If

        'Set the motion blur if needed
        'If UseMotionBlur Then
        '   If BlurIntensity > 0 And BlurIntensity < 255 Or ZoomLevel > 0 Then
        '      FrameUseMotionBlur = True
        '     D3DDevice.SetRenderTarget BlurSurf, Nothing, ByVal 0
        ' Else
        '    FrameUseMotionBlur = False

        'End If

        '  End If

        ' If UseMotionBlur Then
        '    If BlurIntensity < 255 Then
        '       BlurIntensity = BlurIntensity + (timerElapsedTime * 0.01)

        '      If BlurIntensity > 255 Then BlurIntensity = 255

        ' End If

        '       End If

        D3DDevice.BeginScene

        'Clear the screen with a solid color (to prevent artifacts)
        '     If Not FrameUseMotionBlur Then
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0

        '    End If

        '****** Update screen ******
        #If RenderFull = 0 Then

            If Conectar Then
                TT2.Destroy
                Nombres = True
                'frmMain.picHechiz.Visible = False
                frmMain.BarraHechiz.Visible = False
                frmMain.invHechisos.Visible = False
                frmMain.cmdinfo.Visible = False
                frmMain.picture1.Visible = False
                'frmMain.Menu.Visible = False
                'Helios Barras
                '                frmMain.bar_salud(0).Visible = False
                '                frmMain.bar_salud(1).Visible = False
                '                frmMain.Bar_Mana(0).Visible = False
                '                frmMain.Bar_Mana(1).Visible = False
                '                frmMain.bar_sta.Visible = False
                '                frmMain.bar_comida.Visible = False
                'MostrarMenuInventario = False
                'Fin Helios Barras
                'frmMain.picfondoinve.Visible = False Helios elije Menuinventario
                '                frmMain.Bar_Agua.Visible = False    'Helios Barras
                'Helios Barra exp
                frmMain.picInv.Visible = False
                frmMain.PicSpells.Visible = False
                frmMain.barritaa.Visible = False
                frmMain.imgMiniMapa.Visible = False
                frmMain.LanzarImg.Visible = False
                Call RenderConectar
                'ElseIf UserCiego Then
                '    Call CleanViewPort
            Else
                If MostrarMenuInventario = True Then
                ContarClip = 1
                    frmMain.picInv.Visible = True
                    frmMain.PicSpells.Visible = True
                    frmMain.barritaa.Visible = True
                End If
                
               

                'frmMain.picInv.Visible = True
                ' frmMain.imgMiniMapa.Visible = True


                ' frmMain.Menu.Visible = False
                '                frmMain.bar_salud(0).Visible = True
                '                frmMain.Bar_Mana(0).Visible = True
                '
                '                frmMain.bar_sta.Visible = True
                '                frmMain.bar_comida.Visible = True
                '                'frmMain.picfondoinve.Visible = True
                '                frmMain.Bar_Agua.Visible = True

                Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX - 16, OffsetCounterY - 16)
                RenderConsola

                ' Form1.BarraCir.ChangeDefaults UserPasarNivel, RGB(200, 15, 19), 0.25, 0.8, &H777777, "Times New Roman", RGB(255, 255, 255)
            End If

            'End the rendering (scene)
            D3DDevice.EndScene

            'Flip the backbuffer to the screen
            If Conectar Then
                D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
            Else
                D3DDevice.Present RectJuego, ByVal 0, 0, ByVal 0

            End If

        #Else

            If Conectar Then

                Call RenderConectar
                'ElseIf UserCiego Then
                '    Call CleanViewPort
            Else
                Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)

            End If

            'End the rendering (scene)
            D3DDevice.EndScene

            'Flip the backbuffer to the screen
            If Conectar Then
                D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
            Else
                D3DDevice.Present RectJuego, ByVal 0, 0, ByVal 0

            End If

        #End If

        'Screen
        If ScreenShooterCapturePending Then
            DoEvents
            Call ScreenCapture(True)
            ScreenShooterCapturePending = False

        End If

        'Limit FPS to 60 (an easy number higher than monitor's vertical refresh rates)
        'While General_Get_Elapsed_Time2() < 15.5
        '    DoEvents
        'Wend

        'timer_ticks_per_frame = General_Get_Elapsed_Time() * 0.029

        'FPS update
        If fpsLastCheck + 1000 < (GetTickCount() And &H7FFFFFFF) Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = (GetTickCount() And &H7FFFFFFF)
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1

        End If

        'Get timing info
        timerElapsedTime = GetElapsedTime()

        If timerElapsedTime <= 0 Then timerElapsedTime = 1
        timerTicksPerFrame = timerElapsedTime * EngineSpeed()

    End If

End Sub

Public Sub RenderText(ByVal lngXPos As Integer, _
                      ByVal lngYPos As Integer, _
                      ByRef strText As String, _
                      ByVal lngColor As Long)

    If strText <> "" Then
        Call DrawFont(strText, lngXPos, lngYPos, lngColor)

    End If

End Sub

Public Sub RenderTextCentered(ByVal lngXPos As Integer, _
                              ByVal lngYPos As Integer, _
                              ByRef strText As String, _
                              ByVal lngColor As Long)

    If strText <> "" Then
        Call DrawFont(strText, lngXPos, lngYPos, lngColor, True)

    End If

End Sub

Private Function GetElapsedTime() As Single

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim Start_Time    As Currency

    Static End_Time   As Currency

    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq

    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - End_Time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(End_Time)

End Function

Public Sub CharRender(ByRef rChar As Char, _
                      ByVal CharIndex As Integer, _
                      ByVal PixelOffSetX As Integer, _
                      ByVal PixelOffSetY As Integer)

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean

    Dim Pos As Integer

    Dim line As String

    Dim Color As Long

    Dim VelChar As Single

    Dim ColorPj As Long

    Dim ShowPJ As Byte

    Dim ShowPJ_Alpha As Byte

    With rChar

        If .Moving Then
            If .nombre = "" Then
                VelChar = 0.75
            ElseIf left(.nombre, 1) = "!" Then
                VelChar = 0.75
            Else
                VelChar = 1.2

            End If

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * VelChar

                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                If .alaIndex > 0 Then
                    .Alas.GrhIndex(.Heading).Started = 1
                End If
                'Char moved
                moved = True

                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If

            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * VelChar

                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                If .alaIndex > 0 Then
                    .Alas.GrhIndex(.Heading).Started = 1
                End If

                'Char moved
                moved = True

                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If

        If .simbolo <> 0 Then
            'frmMain.TimerSimbolo.Enabled = True
            'Call DrawGrhIndex(3072 & .simbolo, PixelOffSetX, PixelOffSetY + .Body.HeadOffset.y - 61 + SimboloY + 5, 1, D3DColorRGBA(255, 0, 0, 255))
           If .Body.HeadOffset.Y <> 0 Then
            Call DrawGrhIndex(3072 & .simbolo, PixelOffSetX + 2, PixelOffSetY + .Body.HeadOffset.Y - 10 - 10 * Sin((FrameTime Mod 31415) * 0.002) ^ 2, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255))
  Else
             Call DrawGrhIndex(3072 & .simbolo, PixelOffSetX, (PixelOffSetY - 60) + 20 * Sin((FrameTime Mod 31415) * 0.002) ^ 2, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255))
 
     End If

            'frmMain.TimerSimbolo.Enabled = False
        End If

        'If done moving stop animation
        If Not moved And True Then
            'Stop animations
            .Quieto = .Quieto + 1

            If .Quieto >= FPS / 35 Then    'Esto es para que las animacion sean continuas mientras se camine, por ejemplo sin esto el andar del golum se ve feo
                .Quieto = 0

                If .Heading = 0 Then Exit Sub
                .Body.Walk(.Heading).Started = 0
                .Body.Walk(.Heading).FrameCounter = 1
                If .alaIndex > 0 Then
                    .Alas.GrhIndex(.Heading).Started = 0
                    .Alas.GrhIndex(.Heading).FrameCounter = 1
                End If

                If .Arma.WeaponAttack = 0 Then
                    .Arma.WeaponWalk(.Heading).Started = 0
                    .Arma.WeaponWalk(.Heading).FrameCounter = 1
                Else

                    If .Arma.WeaponWalk(.Heading).Started = 0 Then
                        .Arma.WeaponAttack = 0
                        .Arma.WeaponWalk(.Heading).FrameCounter = 1

                    End If

                End If

                If .Escudo.ShieldAttack = 0 Then
                    .Escudo.ShieldWalk(.Heading).Started = 0
                    .Escudo.ShieldWalk(.Heading).FrameCounter = 1
                Else

                    If .Escudo.ShieldWalk(.Heading).Started = 0 Then
                        .Escudo.ShieldAttack = 0
                        .Escudo.ShieldWalk(.Heading).FrameCounter = 1

                    End If

                End If

            End If

            .Moving = False
        Else
            .Quieto = 0

        End If

        PixelOffSetX = PixelOffSetX + .MoveOffsetX
        PixelOffSetY = PixelOffSetY + .MoveOffsetY

        'Verificamos si vamos a mostrar el PJ
        ShowPJ = 0
        ShowPJ_Alpha = 0

        If Not .invisible Then
            ShowPJ = 1
            ShowPJ_Alpha = 255
        ElseIf UserCharIndex = CharIndex Then
            ShowPJ = 2
            ShowPJ_Alpha = 120
        ElseIf .invisible = True Then

            If charEsClan(CharIndex) Then
                ShowPJ = 3
                ShowPJ_Alpha = 120
            ElseIf .oculto = True Then
                ShowPJ = 0
                ShowPJ_Alpha = 0

            End If

        End If

        ColorPj = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, .Alpha)

        If .Heading = 0 Then Exit Sub
        If .Head.Head(.Heading).GrhIndex Then
            If .invisible Then
                If ShowPJ = 2 Or ShowPJ = 3 Then
                    .Alpha = ShowPJ_Alpha
                ElseIf .ContadorInvi > 0 Then

                    If .iTick < (GetTickCount() And &H7FFFFFFF) - 35 Then
                        If .ContadorInvi > 30 And .ContadorInvi <= 60 And .Alpha < 255 Then
                            .Alpha = .Alpha + 5
                        ElseIf .ContadorInvi <= 30 And .Alpha > 0 Then
                            .Alpha = .Alpha - 5

                        End If

                        .ContadorInvi = .ContadorInvi - 1
                        .iTick = (GetTickCount() And &H7FFFFFFF)

                    End If

                Else
                    .ContadorInvi = INTERVALO_INVI

                End If

            End If

            'auras
            'Call Effect_Fire_Begin(Engine_PixelPosX(291), Engine_PixelPosY(855), 8, 100, 180)

            If ActivarAuras = "1" Then

                Dim loopxx As Long

                For loopxx = 0 To 5

                    If .aura(loopxx).AuraGrh Then

                        Dim tmpColor As Long

                        tmpColor = D3DColorRGBA(.aura(loopxx).R, .aura(loopxx).G, .aura(loopxx).b, -1)
                        If tmpColor = D3DColorRGBA(0, 0, 0, -1) Then
                            tmpColor = -1
                        End If
                        If .aura(loopxx).Giratoria = True And RotarActivado = "1" Then
                            Rotacion = Rotacion + 0.5
                            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
                            Call Engine_Render_Rectangle(PixelOffSetX - 35, PixelOffSetY - 40, 100, 100, .aura(loopxx).OffSetX, .aura(loopxx).OffSetX, 128, 128, , , Rotacion, .aura(loopxx).AuraGrh, tmpColor, tmpColor, tmpColor, tmpColor)
                            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                        Else
                            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
                            Call Engine_Render_Rectangle(PixelOffSetX - 35, PixelOffSetY - 40, 100, 100, .aura(loopxx).OffSetX, .aura(loopxx).OffSetX, 128, 128, , , 0, .aura(loopxx).AuraGrh, tmpColor, tmpColor, tmpColor, tmpColor)
                            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

                        End If

                    End If

                Next loopxx

            End If

            'auras

            If .Alpha > 0 Then
                If .priv = 9 Then
                    ColorPj = D3DColorRGBA(10, 10, 10, 255)
                Else
                    ColorPj = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, .Alpha)

                    If .congelado = True Then
                        ColorPj = D3DColorRGBA(0, 175, 255, 225)

                    End If

                End If

                Dim Sombra As Boolean

                If ZonaActual > 0 Then
                    Sombra = .invisible Or .muerto Or Zonas(ZonaActual).Terreno = eTerreno.Dungeon Or .priv = 10

                End If

                Dim TempBodyOffsetX As Integer

                Dim TempBodyOffsetY As Integer

                Dim TempHeadOffsetY As Integer

                Dim TempHeadOffsetX As Integer

                TempHeadOffsetY = .Body.HeadOffset.Y
                TempHeadOffsetX = .Body.HeadOffset.X
                TempBodyOffsetY = PixelOffSetY
                TempBodyOffsetX = PixelOffSetX

                If .Chiquito = True Then    'CHIQUITOLIN
                    TempHeadOffsetY = TempHeadOffsetY + 3
                    TempHeadOffsetX = TempHeadOffsetX - 2
                    TempBodyOffsetY = TempBodyOffsetY + 10
                    TempBodyOffsetX = TempBodyOffsetX + 3

                End If

                '                If .equitando = True Then 'CHIQUITOLIN
                '                    TempHeadOffsetY = TempHeadOffsetY - 67
                '                    TempHeadOffsetX = TempHeadOffsetX + 1
                '                End If

                ' Reflejos en el agua
                Call RenderReflejos(CharIndex, PixelOffSetX, PixelOffSetY)

                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then
                    'If EsNPC(Val(CharIndex)) Then

                    ' End If

                    Call DrawGrhShadow(.Body.Walk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, IIf(Sombra, 0, 1), ColorPj, 255, .Chiquito)

                    'Draw Head
                    Call DrawGrhShadow(.Head.Head(.Heading), TempBodyOffsetX + TempHeadOffsetX, TempBodyOffsetY + TempHeadOffsetY, 1, 0, IIf(Sombra, 0, 2), ColorPj, 255, .Chiquito)

                End If

                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then Call DrawGrhShadow(.Casco.Head(.Heading), TempBodyOffsetX + TempHeadOffsetX, TempBodyOffsetY + TempHeadOffsetY, 1, 0, IIf(Sombra, 0, 2), ColorPj, 255, .Chiquito)

                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then Call DrawGrhShadow(.Arma.WeaponWalk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, IIf(Sombra, 0, 1), ColorPj, 255, .Chiquito)

                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call DrawGrhShadow(.Escudo.ShieldWalk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, IIf(Sombra, 0, 1), ColorPj, 255, .Chiquito)

                'Draw Body
                If .Heading <> north Then
                    If .alaIndex > 0 Then
                        Call DrawGrh(.Alas.GrhIndex(.Heading), TempBodyOffsetX, TempBodyOffsetY - 5, 1, 0.5, 0, ColorPj, 0)
                        Call DrawGrhShadow(.Alas.GrhIndex(.Heading), TempBodyOffsetX - 14, TempBodyOffsetY - 5, 1, 0.5, IIf(Sombra, 0, 1), ColorPj, 255, .Chiquito)
                    End If
                End If
                If .Body.Walk(.Heading).GrhIndex Then Call DrawGrhShadowOff(.Body.Walk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, ColorPj, .Chiquito)
                If .Heading = north Then
                    If .alaIndex > 0 Then
                        Call DrawGrh(.Alas.GrhIndex(.Heading), TempBodyOffsetX, TempBodyOffsetY - 5, 1, 0.5, 0, ColorPj, 0)
                        Call DrawGrhShadow(.Alas.GrhIndex(.Heading), TempBodyOffsetX - 14, TempBodyOffsetY - 5, 1, 0.5, IIf(Sombra, 0, 1), ColorPj, 255, .Chiquito)
                    End If
                End If
                'Draw Head
                Call DrawGrhShadowOff(.Head.Head(.Heading), TempBodyOffsetX + TempHeadOffsetX, TempBodyOffsetY + TempHeadOffsetY, 1, 0, ColorPj, .Chiquito)

                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then Call DrawGrhShadowOff(.Casco.Head(.Heading), TempBodyOffsetX + TempHeadOffsetX, TempBodyOffsetY + TempHeadOffsetY, 1, 0, ColorPj, .Chiquito)

                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then Call DrawGrhShadowOff(.Arma.WeaponWalk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, ColorPj, .Chiquito)

                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call DrawGrhShadowOff(.Escudo.ShieldWalk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, ColorPj, .Chiquito)

                'Draw name over head
                If LenB(.nombre) > 0 And (ShowPJ = 2 Or ShowPJ = 1) And .priv <> 10 Then
                    If Nombres Then
                        Pos = InStr(.nombre, "<")

                        If Pos = 0 Then Pos = Len(.nombre) + 2

                        If .invisible = True Then
                            Color = D3DColorRGBA(200, 200, 200, ShowPJ_Alpha)
                        ElseIf .priv = 0 Then

                            If .Criminal Then
                                Color = D3DColorRGBA(ColoresPJ(50).R, ColoresPJ(50).G, ColoresPJ(50).b, 200)
                            Else
                                Color = D3DColorRGBA(ColoresPJ(49).R, ColoresPJ(49).G, ColoresPJ(49).b, 200)

                            End If

                        Else
                            Color = D3DColorRGBA(ColoresPJ(.priv).R, ColoresPJ(.priv).G, ColoresPJ(.priv).b, 200)

                        End If

                        'Nick
                        line = left$(.nombre, Pos - 2)

                        If left(line, 1) = "!" Then
                            line = right(line, Len(line) - 1)
                            Pos = Pos - 1

                        End If
                        

                        Call RenderTextCentered(PixelOffSetX + TilePixelWidth \ 2, PixelOffSetY + 30, line, Color)

                        'Clan
                        line = mid$(.nombre, Pos)
                        Call RenderTextCentered(PixelOffSetX + TilePixelWidth \ 2, PixelOffSetY + 45, line, Color)

                        '                            If .logged Then
                        '                                color = D3DColorRGBA(10, 200, 10, 200)
                        '                                Call RenderTextCentered(PixelOffSetX + TilePixelWidth \ 2, PixelOffSetY + 45, "(Online)", color)
                        '                            End If
                        '
                    End If

                End If

            End If

        Else

            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then Call DrawGrh(.Body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, VelChar, IIf(Sombra, 0, 1), ColorPj)

        End If

        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffSetX + .Body.HeadOffset.X + 16, PixelOffSetY + .Body.HeadOffset.Y, CharIndex)
        'particulas ore
        Dim I As Integer
        If .particle_count > 0 Then
            For I = 1 To .particle_count
                If .particle_group(I) > 0 Then
                    ParticlesORE.Particle_Group_Render .particle_group(I), PixelOffSetX, PixelOffSetY
                End If
            Next I
        End If
        'particulas ore
        'Draw FX
        If .FxIndex <> 0 Then
            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
            Call DrawGrh(.fX, PixelOffSetX + FxData(.FxIndex).OffSetX, PixelOffSetY + FxData(.FxIndex).OffSetY, 1, 1, 0, D3DColorRGBA(255, 255, 255, 170))
            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

            'Check if animation is over
            If .fX.Started = 0 Then
                .FxIndex = 0

            End If

        End If

    End With

End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, _
                          ByVal fX As Integer, _
                          ByVal Loops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
  With charlist(CharIndex)

        'If fX > 0 Then
        'If CharIndex = UserCharIndex Then  ' And Not UserMeditar
        .FxIndex = fX
        Select Case fX
        Case 0
            .particle_count = 0

        Case 4


            charlist(CharIndex).particle_count = fX
            Call General_Char_Particle_Create(81, CharIndex)
        Case 5

            charlist(CharIndex).particle_count = fX
            Call General_Char_Particle_Create(charlist(CharIndex).particle_count, CharIndex)
        Case Else
            Call InitGrh(.fX, FxData(fX).Animacion)

            .fX.Loops = Loops

        End Select
End With

End Sub

Public Sub SetAreaFx(ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal fX As Integer, _
                     ByVal Loops As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    
    If fX > 0 Then
        Call InitGrh(MapData(X, Y).fXGrh, FxData(fX).Animacion)
        MapData(X, Y).fX = fX
        MapData(X, Y).fXGrh.Loops = Loops

    End If
 
End Sub

Private Sub CleanViewPort()

    'Limpiar
End Sub

Public Function Char_Pos_Get(ByVal CharIndex As Integer, _
                             ByRef X As Integer, _
                             ByRef Y As Integer)
    
    If CharIndex < 1 Then Exit Function

    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        If X > 0 And Y > 0 Then
            Char_Pos_Get = True
        Else
            Char_Pos_Get = False

        End If

    End With

End Function

Public Function charEsClan(ByVal Char As Integer) As Boolean
    charEsClan = False

    If Char > 0 Then

        Dim tempTag   As String

        Dim tempPos   As Integer

        Dim miTag     As String

        Dim miTempPos As Integer

        With charlist(Char)
            miTempPos = getTagPosition(charlist(UserCharIndex).nombre)
            miTag = mid$(charlist(UserCharIndex).nombre, miTempPos)
            tempPos = getTagPosition(.nombre)
            tempTag = mid$(.nombre, tempPos)

            If tempTag = miTag And miTag <> "" And tempTag <> "" Then
                charEsClan = True
                Exit Function

            End If
        
        End With

    End If

End Function

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
    '***************************************************
    'Author: Ezequiel Juarez (Standelf)
    'Last Modification: 16/05/10
    'Blisse-AO | Set a Long Color to a RGB List
    '***************************************************
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)

End Sub

Private Sub RenderReflejos(ByVal CharIndex As Integer, _
                           ByVal PixelOffSetX As Integer, _
                           ByVal PixelOffSetY As Integer)

    '****************************************************
    ' Renderizamos el char reflejado en el agua
    '****************************************************
    On Error GoTo err

    With charlist(CharIndex)
    
        Movement_Speed = 0.5
        
        If HayAgua(.Pos.X, .Pos.Y + 1) Then
                    
            Dim GetInverseHeading  As Byte

            Dim ColorFinal(0 To 3) As Long
            
            'Se anula el viejo reflejo usando Alpha para remplazarlo por transparencia (50%)
            Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(100, 128, 128, 128))

            Select Case .Heading
    
                Case E_Heading.west
                    GetInverseHeading = E_Heading.east

                Case E_Heading.east
                    GetInverseHeading = E_Heading.west

                Case Else
                    GetInverseHeading = .Heading
    
            End Select
                    
            '************ Renderizamos animaciones en los reflejos ************
            If .Moving Then
                .Body.Walk(GetInverseHeading).Started = 1
                .Arma.WeaponWalk(GetInverseHeading).Started = 1
                .Escudo.ShieldWalk(GetInverseHeading).Started = 1
            Else

                '.Body.Walk(GetInverseHeading).Started = 0
                ' .Escudo.ShieldWalk(GetInverseHeading).Started = 0
                '
            End If
                    
            'Animacion del reflejo del arma.
            If .Moving = False Then

                '.Arma.WeaponWalk(GetInverseHeading).Started = 0
            End If
            
            'If .Arma.WeaponWalk(GetInverseHeading).Started = 0 Then
            '   .Arma.WeaponWalk(GetInverseHeading).Started = 1
            '  .Arma.WeaponWalk(GetInverseHeading).FrameCounter = 1
              
            'ElseIf .Arma.WeaponWalk(GetInverseHeading).FrameCounter > 4 Then
            '.attacking = False
    
            'End If
            '************ Renderizamos animaciones en los reflejos ************
                    
            If Not EsNPC(Val(CharIndex)) Then

                'Se anulo el uso de UserNavegando ya que los reflejos de todos los personajes variaban dependiendo de si el usuario navegaba o no.
                If ((.iHead = 0) Or (.iBody = FRAGATA_FANTASMAL)) Then

                    'Reflejo Body Navegando
                    Call Draw_Grh(.Body.Walk(GetInverseHeading), PixelOffSetX, PixelOffSetY + 80, 1, ColorFinal(), 1, False, 360)
                    'Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 0.5, ColorPj)
                            
                    'ElseIf .iBody = 604 Or .iBody = 617 Or .iBody = 612 Or .iBody = 614 Or .iBody = 616 Then 'Define Body Montado
                    
                    'Si ademÃ¡s de estar montado estÃ¡ mirando para arriba o abajo
                    '   If .Heading = E_Heading.SOUTH Or .Heading = E_Heading.NORTH Then
                    'Call DrawGrhShadowOff(.Body.Walk(GetInverseHeading), PixelOffsetX, PixelOffsetY + 80, 1, ColorFinal(), 1, False, 360)
                    'Call DrawGrhShadowOff(.Head.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + 76, 1, ColorFinal(), 0, False, 360)
                    ' Call DrawGrhShadowOff(.Casco.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X - 1, PixelOffsetY + .Body.HeadOffset.Y + 116, 1, ColorFinal(), 0, False, 360)
                    
                    '  Else 'Si estÃ¡ mirando para izquierda o derecha entonces:
                    '  Call DrawGrhShadowOff(.Body.Walk(GetInverseHeading), PixelOffsetX, PixelOffsetY + 70, 1, ColorFinal(), 1, False, 360)
                    '   Call DrawGrhShadowOff(.Head.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + 76, 1, ColorFinal(), 0, False, 360)
                    '    Call DrawGrhShadowOff(.Casco.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X - 1, PixelOffsetY + .Body.HeadOffset.Y + 116, 1, ColorFinal(), 0, False, 360)
                    
                    ' End If
                
                Else
                            
                    'Reflejo completo si no estÃ¡ ni montado ni navegando
                    Call Draw_Grh(.Body.Walk(GetInverseHeading), PixelOffSetX, PixelOffSetY + 44, 1, ColorFinal(), 1, False, 360)
                    Call Draw_Grh(.Head.Head(GetInverseHeading), PixelOffSetX + .Body.HeadOffset.X, PixelOffSetY - .Body.HeadOffset.Y + 15, 1, ColorFinal(), 1, False, 360)
                    Call Draw_Grh(.Casco.Head(GetInverseHeading), PixelOffSetX + .Body.HeadOffset.X, PixelOffSetY + 57, 1, ColorFinal(), 1, False, 360)
                    Call Draw_Grh(.Arma.WeaponWalk(GetInverseHeading), PixelOffSetX, PixelOffSetY + 44, 1, ColorFinal(), 1, False, 360)
                    Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffSetX, PixelOffSetY + 44, 1, ColorFinal(), 0, False, 360)
                
                End If
                        
            End If

        End If
        
    End With

err:

End Sub

Sub Draw_Grh(ByRef Grh As Grh, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal Center As Byte, _
             ByRef Color_List() As Long, _
             ByVal Animate As Byte, _
             Optional ByVal Alpha As Boolean = False, _
             Optional ByVal angle As Single = 0, _
             Optional ByVal ScaleX As Single = 1!, _
             Optional ByVal ScaleY As Single = 1!)

    '*****************************************************************
    'Draws a GRH transparently to a X and Y position
    '*****************************************************************
    Dim CurrentGrhIndex As Long

    Dim FrameDuration   As Single
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
    On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            FrameDuration = Grh.Speed / GrhData(Grh.GrhIndex).NumFrames
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime / FrameDuration) * Movement_Speed
    
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0

                    End If

                End If

            End If

        ElseIf Grh.FrameCounter > 1 Then
            FrameDuration = Grh.Speed / GrhData(Grh.GrhIndex).NumFrames
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime / FrameDuration) * Movement_Speed
    
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = 1

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.PixelWidth * ScaleX - TilePixelWidth) \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If

        ' Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle ScaleX, ScaleY)
            
        Call Engine_Render_Rectangle(X, Y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 180, .FileNum, D3DColorRGBA(137, 200, 200, CalcAlpha(GTCPres, 4000, 150, 15)), D3DColorRGBA(137, 200, 200, CalcAlpha(GTCPres, 4000, 150, 15)), D3DColorRGBA(137, 200, 200, CalcAlpha(GTCPres, 4000, 150, 15)), D3DColorRGBA(137, 200, 200, CalcAlpha(GTCPres, 4000, 150, 15)))
    
    End With
    
    Exit Sub

Error:

    If err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in Draw_Grh, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call CloseClient

    End If

End Sub

Public Sub Grh_Render_To_Hdc(ByRef Pic As PictureBox, _
                             ByVal GrhIndex As Long, _
                             ByVal screen_x As Integer, _
                             ByVal screen_y As Integer, _
                             Optional ByVal Alpha As Integer = False, _
                             Optional ByVal ClearColor As Long = &O0)
    
    On Error GoTo Grh_Render_To_Hdc_Err

    If GrhIndex = 0 Then Exit Sub

    Static Picture As RECT

    With Picture
        .left = 0
        .top = 0

        .bottom = Pic.ScaleHeight
        .right = Pic.ScaleWidth

    End With

    Call D3DDevice.BeginScene
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, ClearColor, 1#, 0)
    
    DrawGrhIndex GrhIndex, screen_x, screen_y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)
    
    Call D3DDevice.EndScene
    Call D3DDevice.Present(Picture, ByVal 0, Pic.hwnd, ByVal 0)
    
    Exit Sub

Grh_Render_To_Hdc_Err:

    ' Call RegistrarError(Err.Number, Err.Description, "TileEngine.Grh_Render_To_Hdc", Erl)
    Resume Next
    
End Sub

Public Sub Grh_Render_To_HdcSinBorrar(ByRef Pic As PictureBox, _
                             ByVal GrhIndex As Long, _
                             ByVal screen_x As Integer, _
                             ByVal screen_y As Integer, _
                             Optional ByVal Alpha As Integer = False, _
                             Optional ByVal ClearColor As Long = &O0)
    
    On Error GoTo Grh_Render_To_Hdc_Err

    If GrhIndex = 0 Then Exit Sub

    Static Picture As RECT

    With Picture
        .left = 0
        .top = 0

        .bottom = Pic.ScaleHeight
        .right = Pic.ScaleWidth

    End With

    Call D3DDevice.BeginScene

    
    DrawGrhIndex GrhIndex, screen_x, screen_y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)
    
    Call D3DDevice.EndScene
    Call D3DDevice.Present(Picture, ByVal 0, Pic.hwnd, ByVal 0)
    
    Exit Sub

Grh_Render_To_Hdc_Err:

    ' Call RegistrarError(Err.Number, Err.Description, "TileEngine.Grh_Render_To_Hdc", Erl)
    Resume Next
    
End Sub

Sub RenderConsola()

    Dim I As Byte

    If OffSetConsola > 0 Then OffSetConsola = OffSetConsola - 1
    If OffSetConsola = 0 Then UltimaLineavisible = True

    For I = 1 To MaxLineas - 1

        RenderText 44, ComienzoY + (I * 15) + OffSetConsola, Con(I).T, D3DColorRGBA(Con(I).R, Con(I).G, Con(I).b, I * (255 / MaxLineas))

    Next I

    If UltimaLineavisible = True Then RenderText 44, ComienzoY + (MaxLineas * 15) + OffSetConsola, Con(I).T, D3DColorRGBA(Con(MaxLineas).R, Con(MaxLineas).G, Con(I).b, 255)

End Sub

Public Function ParticulaX(ByVal PosUserX As Integer, ByVal PosConX As Integer) As Integer


'On Error GoTo ParticulaX_Err
    Dim Medio As Long
    Dim Parcial As Long
    Medio = Round(frmMain.pRender.Width / 2)

    If PosUserX > PosConX Then
        Parcial = (PosUserX - PosConX) * 32
        ParticulaX = Medio + Parcial
        Exit Function
    Else
        Parcial = (PosConX - PosUserX) * 32
        ParticulaX = Medio - Parcial
        Exit Function
    End If





    'ParticulaX_Err:

    'Resume Next

End Function

Public Function ParticulaY(ByVal PosUserY As Integer, ByVal PosConY As Integer) As Integer

    'On Error GoTo ParticulaY_Err
    Dim Medio As Long
    Dim Parcial As Long
    Medio = Round(frmMain.pRender.Height / 2)
     If PosUserY > PosConY Then
        Parcial = (PosUserY - PosConY) * 32
        ParticulaY = Medio + Parcial
        Exit Function
    Else
        Parcial = (PosConY - PosUserY) * 32
        ParticulaY = Medio - Parcial
        Exit Function
    End If
   
  
   

    Exit Function

'ParticulaY_Err:

    Resume Next

End Function



