Attribute VB_Name = "Mod_General"
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

'Particulas
'********************************

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
'Particulas
'***************************
Option Explicit
#If RenderFull = 0 Then

    Public frmMain As frmMain2
#Else

    Public frmMain As frmMain1
#End If

Public PCred()             As Integer

Public PCgreen()           As Integer

Public PCblue()            As Integer

Public iplst               As String

Public bFogata             As Boolean

Private lFrameTimer        As Long

Public IpServidor          As String

Public PuertoServidor      As Long

Public PathGraficos        As String

Public PathRecursosCliente As String

Public PathWav             As String

Public PathInterface       As String

Public PathInit            As String

Private Type TConsola

    Texto As String
    Color As Long
    bold As Byte
    italic As Byte

End Type

Public Consola()     As TConsola

Public LineasConsola As Integer

Public ArchivoMapa   As Integer

Public DataMap1()    As Byte

Public DataMap2()    As Byte

Public Map1Loaded    As Boolean

Public Map2Loaded    As Integer

Public MapInfo       As MapInformation

Public Function DirInterface() As String
    DirInterface = App.path & "\" & Config_Inicio.DirGraficos & "\Interface\"

End Function

Public Function DirGraficos() As String
    DirGraficos = App.path & "\" & Config_Inicio.DirGraficos & "\"

End Function

Public Function DirSound() As String
    DirSound = App.path & "\" & Config_Inicio.DirSonidos & "\"

End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\" & Config_Inicio.DirMusica & "\"

End Function

Public Function DirRecursos() As String
    DirRecursos = PathRecursosCliente & "\Recursos\"

End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

Sub CargarAnimArmas()

    On Error Resume Next

    Dim LoopC      As Long

    Dim N          As Integer

    Dim MisArmas() As tIndiceArma

    N = FreeFile()
    Open PathInit & "\Armas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumWeaponAnims

    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    ReDim MisArmas(1 To NumWeaponAnims) As tIndiceArma
    
    For LoopC = 1 To NumWeaponAnims
        Get #N, , MisArmas(LoopC)
    
        InitGrh WeaponAnimData(LoopC).WeaponWalk(1), MisArmas(LoopC).Arma(1), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(2), MisArmas(LoopC).Arma(2), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(3), MisArmas(LoopC).Arma(3), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(4), MisArmas(LoopC).Arma(4), 0
    Next LoopC
    
    Close #N
    
End Sub

Sub CargarColores()

    On Error Resume Next

    Dim archivoC As String
    
    archivoC = PathInit & "\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
        'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub

    End If
    
    Dim I As Long
    
    For I = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(I).R = CByte(GetVar(archivoC, CStr(I), "R"))
        ColoresPJ(I).G = CByte(GetVar(archivoC, CStr(I), "G"))
        ColoresPJ(I).b = CByte(GetVar(archivoC, CStr(I), "B"))
    Next I
    
    ColoresPJ(50).R = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).G = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).R = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).G = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))

End Sub

Sub CargarZonas()

    On Error Resume Next

    Dim archivoC As String
    
    archivoC = PathInit & "\zonas.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
        'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar las zonas. Falta el archivo zonas.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub

    End If
    
    Dim I As Integer

    Dim e As Integer
    
    NumZonas = GetVar(archivoC, "Config", "Cantidad")
    
    ReDim Zonas(1 To NumZonas)

    For I = 1 To NumZonas

        With Zonas(I)
            .nombre = GetVar(archivoC, "Zona" & CStr(I), "Nombre")
            .Mapa = CByte(GetVar(archivoC, "Zona" & CStr(I), "Mapa"))
            .x1 = CInt(GetVar(archivoC, "Zona" & CStr(I), "X1"))
            .y1 = CInt(GetVar(archivoC, "Zona" & CStr(I), "Y1"))
            .x2 = CInt(GetVar(archivoC, "Zona" & CStr(I), "X2"))
            .y2 = CInt(GetVar(archivoC, "Zona" & CStr(I), "Y2"))
            .Segura = CByte(GetVar(archivoC, "Zona" & CStr(I), "Segura"))
            .Acoplar = CByte(Val(GetVar(archivoC, "Zona" & CStr(I), "Acoplar")))
            .Terreno = CByte(Val(GetVar(archivoC, "Zona" & CStr(I), "Terreno")))
            .Niebla = CByte(Val(GetVar(archivoC, "Zona" & CStr(I), "Niebla")))
            .NieblaR = CByte(Val(GetVar(archivoC, "Zona" & CStr(I), "NieblaR")))
            .NieblaG = CByte(Val(GetVar(archivoC, "Zona" & CStr(I), "NieblaG")))
            .NieblaB = CByte(Val(GetVar(archivoC, "Zona" & CStr(I), "NieblaB")))
            .Musica(1) = Val(GetVar(archivoC, "Zona" & CStr(I), "Musica1"))
            .Musica(2) = Val(GetVar(archivoC, "Zona" & CStr(I), "Musica2"))
            .Musica(3) = Val(GetVar(archivoC, "Zona" & CStr(I), "Musica3"))
            .Musica(4) = Val(GetVar(archivoC, "Zona" & CStr(I), "Musica4"))
            .Musica(5) = Val(GetVar(archivoC, "Zona" & CStr(I), "Musica5"))
            
            .Sonido(1) = Val(GetVar(archivoC, "Zona" & CStr(I), "Sonido1"))
            .Sonido(2) = Val(GetVar(archivoC, "Zona" & CStr(I), "Sonido2"))
            .Sonido(3) = Val(GetVar(archivoC, "Zona" & CStr(I), "Sonido3"))
            .Sonido(4) = Val(GetVar(archivoC, "Zona" & CStr(I), "Sonido4"))
            .Sonido(5) = Val(GetVar(archivoC, "Zona" & CStr(I), "Sonido5"))
                       
            If .NieblaR = 0 And .NieblaG = 0 And .NieblaB = 0 Then
                .NieblaR = 255
                .NieblaG = 200
                .NieblaB = 200

            End If
                       
            For e = 1 To 5

                If .Musica(e) > 0 Then .CantMusica = .CantMusica + 1
                If .Sonido(e) > 0 Then .CantSonidos = .CantSonidos + 1
            Next e

        End With

    Next I

End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()

    On Error Resume Next

    Dim LoopC        As Long

    Dim N            As Integer

    Dim MisEscudos() As tIndiceArma

    N = FreeFile()
    Open PathInit & "\Escudos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumEscudosAnims

    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    ReDim MisEscudos(1 To NumEscudosAnims) As tIndiceArma
    
    For LoopC = 1 To NumEscudosAnims
        Get #N, , MisEscudos(LoopC)
        
        InitGrh ShieldAnimData(LoopC).ShieldWalk(1), MisEscudos(LoopC).Arma(1), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(2), MisEscudos(LoopC).Arma(2), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(3), MisEscudos(LoopC).Arma(3), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(4), MisEscudos(LoopC).Arma(4), 0
    Next LoopC
    
    Close #N

End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()

    '*****************************************************************
    'Goes through the charlist and replots all the characters on the map
    'Used to make sure everyone is visible
    '*****************************************************************
    Dim LoopC As Long
    
    For LoopC = 1 To LastChar

        If charlist(LoopC).ACTIVE = 1 Then
            MapData(charlist(LoopC).Pos.X, charlist(LoopC).Pos.Y).CharIndex = LoopC

        End If

    Next LoopC

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean

    Dim car As Byte

    Dim I   As Long
    
    cad = LCase$(cad)
    
    For I = 1 To Len(cad)
        car = Asc(mid$(cad, I, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function

        End If

    Next I
    
    AsciiValidos = True

End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean

    'Validamos los datos del user
    Dim LoopC     As Long

    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MessageBox ("Dirección de email invalida")
        Exit Function

    End If
    
    If UserPassword = "" Then
        MessageBox ("Ingrese un password.")
        Exit Function

    End If
    
    For LoopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, LoopC, 1))

        If Not LegalCharacter(CharAscii) Then
            MessageBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function

        End If

    Next LoopC
    
    If UserName = "" Then
        MessageBox ("Ingrese un nombre de personaje.")
        Exit Function

    End If
    
    If Len(UserName) > 30 Then
        MessageBox ("El nombre debe tener menos de 30 letras.")
        Exit Function

    End If
    
    For LoopC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, LoopC, 1))

        If Not LegalCharacter(CharAscii) Then
            MessageBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function

        End If

    Next LoopC
    
    CheckUserData = True

End Function

Sub UnloadAllForms()

    On Error Resume Next

    #If SeguridadAlkon Then
        Call UnprotectForm
    #End If

    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next

End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean

    '*****************************************************************
    'Only allow characters that are Win 95 filename compatible
    '*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function

    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function

    End If
    
    If KeyAscii > 126 Then
        Exit Function

    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function

    End If
    
    'else everything is cool
    LegalCharacter = True

End Function

Sub SetConnected()
    '*****************************************************************
    'Sets the client to "Connect" mode
    '*****************************************************************
    'Set Connected
    Connected = True
    
    #If SeguridadAlkon Then
        'Unprotect character creation form
        Call UnprotectForm
    #End If
    
    'Unload the connect form
    Unload frmCrearPersonaje
    
    'frmMain.label8.Caption = UserName
    'Load main form
    frmMain.Visible = True
    
    Conectar = False
    
    Audio.StopMp3
    
    ZonaActual = 0
    LastZona = ""
    CheckZona
        
    'frmMain.SetRender (True)
    
    #If SeguridadAlkon Then
        'Protect the main form
        Call ProtectForm(frmMain)
    #End If

End Sub

Sub ChangeDirection(ByVal Direccion As E_Heading)
    Call WriteChangeHeading(Direccion)

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)

    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/28/2008
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
    ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
    ' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
    '***************************************************
    Dim LegalOk As Boolean

    If Conectar Or UserEmbarcado Then Exit Sub
    If Cartel Then Cartel = False
    
    Select Case Direccion

        Case E_Heading.north
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)

        Case E_Heading.east
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)

        Case E_Heading.south
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)

        Case E_Heading.west
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)

    End Select
    
    If TiempoRetos = 0 Then
        If LegalOk And Not UserParalizado Then
            If Not UserDescansar And Not UserMeditar Then
                Call WriteWalk(Direccion)
                MoveCharbyHead UserCharIndex, Direccion
                MoveScreen Direccion

            End If

        Else

            If charlist(UserCharIndex).Heading <> Direccion Then
                Call WriteChangeHeading(Direccion)

            End If

        End If

    End If
    
    If frmMain.macrotrabajo.Enabled Then frmMain.DesactivarMacroTrabajo
    
    Call frmMain.RefreshMiniMap
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
    
    CheckZona

End Sub

Sub RandomMove()
    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/03/2006
    ' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
    '***************************************************
    Call MoveTo(RandomNumber(north, west))

End Sub

Private Sub CheckKeys()

    '*****************************************************************
    'Checks keys and respond
    '*****************************************************************
    Static lastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    If WAIT_ACTION = eWAIT_FOR_ACTION.RPU Or Not MainTimer.Check(TimersIndex.PuedeRPUMover) Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.PuedeMover) Then Exit Sub
    
    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    '    If Abs((GetTickCount() And &H7FFFFFFF) - lastMovement) > 36 Then
    '        lastMovement = (GetTickCount() And &H7FFFFFFF)
    '    Else
    '        Exit Sub
    '    End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido And Not Conectar Then

            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call MoveTo(north)
                Exit Sub

            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call MoveTo(east)
                Exit Sub

            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call MoveTo(south)
                Exit Sub

            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call MoveTo(west)
                Exit Sub

            End If
                        
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else

            Dim kp As Boolean

            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)

            End If
            
            frmMain.Coord.Caption = "(" & UserPos.X & "," & UserPos.Y & ")"
            CheckZona

        End If

    End If

End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!

Sub CargarMap(ByVal Map As Integer)

    '**************************************************************
    'Formato de mapas optimizado para reducir el espacio que ocupan.
    'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
    '**************************************************************
    Dim Y       As Long

    Dim X       As Long

    Dim tempint As Integer

    Dim ByFlags As Byte

    Dim Handle  As Integer
    
    'If ArchivoMapa > 0 Then
       
    'End If
    
    If (Map = 1 And Not Map1Loaded) Or (Map > 1 And Map <> Map2Loaded) Then
    
        ArchivoMapa = FreeFile()
                
        Open DirRecursos & "Mapa" & Map & ".AO" For Binary As ArchivoMapa

        If Map = 1 Then
            ReDim DataMap1(LOF(ArchivoMapa))
            Get #ArchivoMapa, , DataMap1
            Map1Loaded = True
        Else
            ReDim DataMap2(LOF(ArchivoMapa))
            Get #ArchivoMapa, , DataMap2
            Map2Loaded = Map

        End If
            
        Close ArchivoMapa
    
    End If
    
    Dim Pos    As Integer

    Dim tmpInt As Integer

    Dim I      As Integer
    
    If Map = 1 Then
        Pos = 0
        tmpInt = (DataMap1(Pos + 1) And &H7F) * &H100 Or DataMap1(Pos) Or -(DataMap1(Pos + 1) > &H7F) * &H8000

        If tmpInt = 23678 Then
            Pos = Pos + 2
            tmpInt = (DataMap1(Pos + 1) And &H7F) * &H100 Or DataMap1(Pos) Or -(DataMap1(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2
            MapInfo.Width = (DataMap1(Pos + 1) And &H7F) * &H100 Or DataMap1(Pos) Or -(DataMap1(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2
            MapInfo.Height = (DataMap1(Pos + 1) And &H7F) * &H100 Or DataMap1(Pos) Or -(DataMap1(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2
            MapInfo.MapVersion = (DataMap1(Pos + 1) And &H7F) * &H100 Or DataMap1(Pos) Or -(DataMap1(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2
            tmpInt = (DataMap1(Pos + 1) And &H7F) * &H100 Or DataMap1(Pos) Or -(DataMap1(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2

            For I = Pos To Pos + tmpInt - 1
                MapInfo.Name = MapInfo.Name & Chr(DataMap1(I))
            Next I

            Pos = Pos + tmpInt
            
            For I = Pos To Pos + 9
                MapInfo.Date = MapInfo.Date & Chr(DataMap1(I))
            Next I
            
            Pos = Pos + 10
                    
            MapInfo.offset = Pos

        End If

    Else
        Pos = 0
        tmpInt = (DataMap2(Pos + 1) And &H7F) * &H100 Or DataMap2(Pos) Or -(DataMap2(Pos + 1) > &H7F) * &H8000

        If tmpInt = 23678 Then
            Pos = Pos + 2
            tmpInt = (DataMap2(Pos + 1) And &H7F) * &H100 Or DataMap2(Pos) Or -(DataMap2(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2
            MapInfo.Width = (DataMap2(Pos + 1) And &H7F) * &H100 Or DataMap2(Pos) Or -(DataMap2(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2
            MapInfo.Height = (DataMap2(Pos + 1) And &H7F) * &H100 Or DataMap2(Pos) Or -(DataMap2(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2
            MapInfo.MapVersion = (DataMap2(Pos + 1) And &H7F) * &H100 Or DataMap2(Pos) Or -(DataMap2(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2
            tmpInt = (DataMap2(Pos + 1) And &H7F) * &H100 Or DataMap2(Pos) Or -(DataMap2(Pos + 1) > &H7F) * &H8000
            Pos = Pos + 2

            For I = Pos To Pos + tmpInt - 1
                MapInfo.Name = MapInfo.Name & Chr(DataMap2(I))
            Next I

            Pos = Pos + tmpInt
            
            For I = Pos To Pos + 9
                MapInfo.Date = MapInfo.Date & Chr(DataMap2(I))
            Next I
            
            Pos = Pos + 10
                    
            MapInfo.offset = Pos

        End If

    End If
    
    For Y = 1 To YMaxMapSize
        For X = 1 To XMaxMapSize

            If MapData(X, Y).CharIndex > 0 Then
                Call EraseChar(MapData(X, Y).CharIndex)

            End If

            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
        Next X
    Next Y
    
End Sub

Sub SwitchMap(ByVal Map As Integer)
    '**************************************************************
    'Formato de mapas optimizado para reducir el espacio que ocupan.
    'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
    '**************************************************************
    UserMap = Map
    CurMap = Map
    
    CargarMap (Map)
  'Call Effect_RedFountain_Begin(Engine_PixelPosX(302), Engine_PixelPosY(859), 1, 1000)
    Call General_Particle_Create(106, 741, 802)
    Call Effect_Snow_Begin(13, 50)

End Sub

Sub AddtoRichPicture(ByVal Text As String, _
                     Optional ByVal red As Integer = -1, _
                     Optional ByVal green As Integer, _
                     Optional ByVal blue As Integer, _
                     Optional ByVal bold As Boolean = False, _
                     Optional ByVal italic As Boolean = False, _
                     Optional ByVal bCrLf As Boolean = False)

'lo pongo aca, para q no tengan q andar cambiando todo
'osea, si tienen consola de arriba, el richtextbox, no agan esto
    Dim TextosConsola As String
    If frmMain.macrotrabajo Then
    If Text = "Yacimiento de Hierro - 1" Or Text = "Yacimiento de Plata - 1" Then
    Exit Sub
    End If
    End If
    TextosConsola = Date & " " & CStr(Time) & ": " & Text


    'Open App.path & "/INIT/" & UserName & ".txt" For Append As #1
   'Print #1, textosconsola
   ' Close #1
   If red = 190 Then
   red = 255
   green = 255
   blue = 255
   End If
   
    
    AddtoRichTextBox FrmMensajes.mensajes, TextosConsola, red, green, blue, True, False, bCrLf
   FrmMensajes.mensajes.SaveFile (App.path & "/INIT/" & UserName & ".rtf")


    If sintextos = False Then

        If Text > " " Then Exit Sub    ' compruebo si el texto es mayo a un espacio no imprime  :D
        GoTo a    ' imprime hasta 6 veces
        Exit Sub

    End If

    #If RenderFull = 0 Then

        If left(Text, 1) = " " Then Exit Sub
a:

        Dim I As Byte

        For I = 2 To MaxLineas
            Con(I - 1).T = Con(I).T
            'Con(i - 1).Color = Con(i).Color
            Con(I - 1).b = Con(I).b
            Con(I - 1).G = Con(I).G
            Con(I - 1).R = Con(I).R
        Next I

        Con(MaxLineas).T = Text
        Con(MaxLineas).b = blue
        Con(MaxLineas).G = green
        Con(MaxLineas).R = red
        OffSetConsola = 16

        UltimaLineavisible = False
    #Else

        Dim nId As Long

        Dim AText As String

        Dim Lineas() As String

        Dim I As Integer

        Dim l As Integer

        Dim LastEsp As Integer

        Lineas = Split(Text, vbCrLf)

        For l = 0 To UBound(Lineas)
            Text = Lineas(l)
            nId = LineasConsola + 1

            If nId = 601 Then

                For I = 0 To 500
                    Consola(I) = Consola(I + 100)
                Next I

                nId = 501

                If OffSetConsola > 101 Then OffSetConsola = OffSetConsola - 100

            End If

            LineasConsola = nId
            frmMain.pConsola.FontBold = bold
            frmMain.pConsola.FontItalic = italic
            Consola(nId).Texto = Text
            Consola(nId).Color = RGB(red, green, blue)
            Consola(nId).bold = bold
            Consola(nId).italic = italic

            If LineasConsola > 6 Then
                OffSetConsola = LineasConsola - 6
                frmMain.BarritaConsola.top = 68

            End If

            If frmMain.pConsola.TextWidth(Text) > frmMain.pConsola.Width Then
                LastEsp = 0

                For I = 1 To Len(Text)

                    If mid(Text, I, 1) = " " Then LastEsp = I
                    If frmMain.pConsola.TextWidth(left$(Text, I)) > frmMain.pConsola.Width Then Exit For
                Next I

                If LastEsp = 0 Then LastEsp = I
                AText = right$(Text, Len(Text) - LastEsp)
                Text = left$(Text, LastEsp)
                Consola(nId).Texto = Text
                Call AddtoRichPicture(AText, red, green, blue, bold, italic)
            Else
                frmMain.ReDrawConsola

            End If

        Next l

    #End If

End Sub

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String

    '*****************************************************************
    'Gets a field from a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    '*****************************************************************
    Dim I          As Long

    Dim LastPos    As Long

    Dim CurrentPos As Long

    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For I = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next I
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

    End If

End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long

    '*****************************************************************
    'Gets the number of fields in a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 07/29/2007
    '*****************************************************************
    Dim count     As Long

    Dim curPos    As Long

    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count

End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")

End Function

Public Function IsIp(ByVal Ip As String) As Boolean

    Dim I As Long
    
    For I = 1 To UBound(ServersLst)

        If ServersLst(I).Ip = Ip Then
            IsIp = True
            Exit Function

        End If

    Next I

End Function

Sub Main()

    #If RenderFull = 0 Then
        Set frmMain = frmMain2
    #Else
        Set frmMain = frmMain1
    #End If
    sintextos = True

    If FileExist(App.path & "\Init\Config.ini", vbNormal) Then
        Call ReadConfig
    Else
        Call mOpciones_Default

    End If

    'Load config file
    If FileExist(PathInit & "\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()

    End If

    WAIT_ACTION = 0

    AlphaSalir = 255
MostrarMenuInventario = True
MmenuBarras = True
    Set curGeneral = New clsAniCursor
    Set curGeneralCrimi = New clsAniCursor
    Set curGeneralCiuda = New clsAniCursor
    Set curProyectil = New clsAniCursor
    Set curProyectilPequena = New clsAniCursor

    Set picMouseIcon = LoadPicture(PathRecursosCliente & "\Recursos\MouseIcons\Baston.ico")
    curGeneral.AniFile = PathRecursosCliente & "\Recursos\MouseIcons\General.ani"
    curGeneralCrimi.AniFile = PathRecursosCliente & "\Recursos\MouseIcons\GeneralCrimi.ani"
    curGeneralCiuda.AniFile = PathRecursosCliente & "\Recursos\MouseIcons\GeneralCiuda.ani"
    curProyectil.AniFile = PathRecursosCliente & "\Recursos\MouseIcons\Mira.ani"
    curProyectilPequena.AniFile = PathRecursosCliente & "\Recursos\MouseIcons\MiraPequena.ani"

    curGeneral.CursorOn frmMain.hwnd
    curGeneral.CursorOn frmMain.pRender.hwnd

    frmMain.picHechiz.MouseIcon = picMouseIcon
    frmMain.LanzarImg.MouseIcon = picMouseIcon
    'Helios Menu desconectar 14/06/21 07:54
    frmCerrar.Opcion(0).MouseIcon = picMouseIcon
    frmCerrar.Opcion(1).MouseIcon = picMouseIcon
    frmCerrar.Opcion(2).MouseIcon = picMouseIcon
    'frmMain.btnHechizos.MouseIcon = picMouseIcon
    'frmMain.btnInventario.MouseIcon = picMouseIcon

    Dim picMousePointIcon As Picture

    Set picMousePointIcon = LoadPicture(PathRecursosCliente & "\Recursos\MouseIcons\Point.ico")
    '    frmMain.Image1(0).MouseIcon = picMousePointIcon    'Opciones
    '    frmMain.Image1(1).MouseIcon = picMousePointIcon    'Stats
    '    frmMain.Image1(2).MouseIcon = picMousePointIcon    'Clanes
    '    frmMain.Image1(3).MouseIcon = picMousePointIcon    'Quests
    '    'frmMain.Image1(4).MouseIcon = picMousePointIcon 'Party
    '    'Set picMousePointIcon = LoadPicture(PathRecursosCliente & "\MouseIcons\Espada.ico")
    '    'frmMain.btnInventario.MouseIcon = picMousePointIcon
    'aura
    CargarAuras
    'aura
    Call InitDebug

    'Load ao.dat config file
    If FileExist(PathInit & "\ao.dat", vbArchive) Then
        Call LoadClientSetup

        If ClientSetup.bDinamic Then
            Set SurfaceDB = New clsSurfaceManDyn
        Else
            Set SurfaceDB = New clsSurfaceManStatic

        End If

    Else
        'Use dynamic by default
        Set SurfaceDB = New clsSurfaceManDyn

    End If

    'If FindPreviousInstance Then
    '    Call MsgBox("AoYind ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    'End
    'End If

    Call LeerLineaComandos
    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.

    ReDim SurfaceSize(15000)
    ReDim Consola(600)
    ReDim PCred(600)
    ReDim PCgreen(600)
    ReDim PCblue(600)
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")

    ChDrive App.path
    ChDir App.path

    #If SeguridadAlkon Then

        'Obtener el HushMD5
        Dim fMD5HushYo As String * 32

        fMD5HushYo = MD5.GetMD5File(App.path & "\" & App.EXEName & ".exe")
        Call MD5.MD5Reset
        MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 55)

        Debug.Print fMD5HushYo
    #Else
        MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
    #End If

    tipf = Config_Inicio.tip

    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    Call Resolution.SetResolution(1024, 768)

    'Set picMouseIcon = LoadPicture(DirRecursos & "Hand.ico")

    frmCargando.Show
    frmCargando.Refresh

    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.05
    DoEvents

    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.2
    DoEvents

    'TODO : esto de ServerRecibidos no se podría sacar???
    ServersRecibidos = True

    Call InicializarNombres

    ' Initialize FONTTYPES
    Call Protocol.InitFonts

    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.25
    DoEvents
    #If RenderFull = 0 Then

        If Not InitTileEngine(frmMain.hwnd, frmMain.top, frmMain.pRender.left, 32, 32, 24, 32, 14, 9, 9, 0.018) Then

            Call CloseClient

        End If

    #Else

        If Not InitTileEngine(frmMain.hwnd, 125, 2, 32, 32, 19, 25, 9, 9, 9, 0.018) Then
            Call CloseClient

        End If

    #End If

    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.4
    DoEvents

    UserMap = 0

    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call CargarZonas
    Call CargarPasos
    Call CargarTutorial
    'quest
    Call CargarNpc
    Call CargarQuests
    Call CargarObjetos
    'quest

    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.45
    DoEvents

    'Inicializamos el sonido
    Call Audio.Initialize(dX, frmMain.hwnd, App.path & "\" & Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & "\", App.path & "\" & Config_Inicio.DirMusica & "\")
    'Enable / Disable audio

    'Audio
    Audio.MusicActivated = mOpciones.Music
    Audio.SoundActivated = mOpciones.sound
    Audio.SoundEffectsActivated = mOpciones.SoundEffects
    Audio.MusicVolume = mOpciones.VolMusic
    Audio.SoundVolume = mOpciones.VolSound

    SinMidi = False

    'Guilds
    DialogosClanes.CantidadDialogos = mOpciones.DialogCantMessages

    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv, 0, 0, MAX_INVENTORY_SLOTS, True)
    
     'Inicializa el inventario de hechizos
    Call invSpells.Initialize(frmMain.PicSpells, 0, 0, MAX_INVENTORY_SLOTS, True)

    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.55
    DoEvents

    Call CargarMap(1)
    'Call CargarMap(2)

    frmCargando.BProg.Width = frmCargando.BBProg.Width * 1
    DoEvents

    UserSeguroResu = True
    UserSeguro = True

    #If SeguridadAlkon Then
        CualMI = 0
        Call InitMI
    #End If

    Unload frmCargando

    Call Audio.PlayMIDI(MIdi_Inicio & ".mid")

    Set frmMain.Client = New CSocketMaster

    frmMain.SetRender (True)
    frmMain.Show

    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False

    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.PuedeGolpe, INT_PUEDE_GOLPE)
    Call MainTimer.SetInterval(TimersIndex.HabilitaLanzarHechizo, INT_HABILITA_ICONO_LANZAR_HECHIZO)
    Call MainTimer.SetInterval(TimersIndex.PuedeLanzarHechizo, INT_LANZAR_HECHIZO)
    Call MainTimer.SetInterval(TimersIndex.PuedeGolpeMagia, INT_GOLPE_MAGIA)
    Call MainTimer.SetInterval(TimersIndex.PuedeMagiaGolpe, INT_MAGIA_GOLPE)
    Call MainTimer.SetInterval(TimersIndex.PuedeFlechas, INT_FLECHAS)
    Call MainTimer.SetInterval(TimersIndex.PuedeUsar, INT_PUEDE_USAR)
    Call MainTimer.SetInterval(TimersIndex.PuedeUsarDobleClick, INT_PUEDE_USAR_DOBLECLICK)
    Call MainTimer.SetInterval(TimersIndex.PuedeGolpeUsar, INT_GOLPE_USAR)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.Hide, INT_HIDE)
    Call MainTimer.SetInterval(TimersIndex.Buy, INT_BUY)
    Call MainTimer.SetInterval(TimersIndex.Montar, INT_MONTAR)
    Call MainTimer.SetInterval(TimersIndex.Anclar, INT_ANCLAR)
    Call MainTimer.SetInterval(TimersIndex.Telep, INT_TELEP)
    Call MainTimer.SetInterval(TimersIndex.PuedeMover, INT_PUEDE_MOVER)
    Call MainTimer.SetInterval(TimersIndex.PuedeMoverEquitando, INT_PUEDE_MOVER_EQUITANDO)
    Call MainTimer.SetInterval(TimersIndex.PuedeRPUMover, INT_PUEDE_RPU_MOVER)

    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False

    'Init timers
    Call MainTimer.Start(TimersIndex.PuedeGolpe)
    Call MainTimer.Start(TimersIndex.HabilitaLanzarHechizo)
    Call MainTimer.Start(TimersIndex.PuedeLanzarHechizo)
    Call MainTimer.Start(TimersIndex.PuedeGolpeMagia)
    Call MainTimer.Start(TimersIndex.PuedeMagiaGolpe)
    Call MainTimer.Start(TimersIndex.PuedeFlechas)
    Call MainTimer.Start(TimersIndex.PuedeUsar)
    Call MainTimer.Start(TimersIndex.PuedeUsarDobleClick)
    Call MainTimer.Start(TimersIndex.PuedeGolpeUsar)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.Hide)
    Call MainTimer.Start(TimersIndex.Buy)
    Call MainTimer.Start(TimersIndex.Montar)
    Call MainTimer.Start(TimersIndex.Anclar)
    Call MainTimer.Start(TimersIndex.Telep)
    Call MainTimer.Start(TimersIndex.PuedeMover)
    Call MainTimer.Start(TimersIndex.PuedeMoverEquitando)
    Call MainTimer.Start(TimersIndex.PuedeRPUMover)

    'Set the dialog's font
    Dialogos.font = frmMain.font
    DialogosClanes.font = frmMain.font

    ' Load the form for screenshots
    Call Load(frmScreenshots)

    'Call Audio.PlayingMusic("9.mp3")

    Call InitBarcos

    GTCInicial = (GetTickCount() And &H7FFFFFFF)

    Conectar = True
    EngineRun = True

    Nombres = True

    If mOpciones.Recordar = True Then
        frmMain.tUser.Text = mOpciones.RecordarUsuario
        frmMain.tPass.Text = mOpciones.RecordarPassword

    End If

    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada

        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call CalcularBarcos
            Call ShowNextFrame(frmMain.top, frmMain.left, frmMain.MouseX, frmMain.MouseY)

            'Play ambient sounds
            Call RenderSounds

            Call CheckKeys
        Else
            Call CalcularBarcos

        End If

        'FPS Counter - mostramos las FPS
        If Abs((GetTickCount() And &H7FFFFFFF) - lFrameTimer) >= 1000 Then
            lFrameTimer = (GetTickCount() And &H7FFFFFFF)

        End If

        #If SeguridadAlkon Then
            Call CheckSecurity
        #End If

        ' If there is anything to be sent, we send it
        Call FlushBuffer

        DoEvents
    Loop

    Call CloseClient

End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)

End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

    On Error GoTo errHnd

    Dim lPos As Long

    Dim lX   As Long

    Dim iAsc As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")

    If (lPos <> 0) Then

        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1

            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))

                If Not CMSValidateChar_(iAsc) Then Exit Function

            End If

        Next lX
        
        'Finale
        CheckMailString = True

    End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()

    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus

    End If

End Sub

Public Sub ShowSendCMSGTxt()

    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus

    End If

End Sub

''
' Checks the command line parameters, if you are running Ao with /nores command and checks the AoUpdate parameters
'
'

Public Sub LeerLineaComandos()

    '*************************************************
    'Author: Unknown
    'Last modified: 25/11/2008 (BrianPr)
    '
    '*************************************************
    Dim T()      As String

    Dim I        As Long
    
    Dim UpToDate As Boolean

    Dim Patch    As String
    
    'Parseo los comandos
    T = Split(Command, " ")

    For I = LBound(T) To UBound(T)

        Select Case UCase$(T(I))

            Case "/NORES" 'no cambiar la resolucion
                NoRes = True

            Case "/UPTODATE"
                UpToDate = True

        End Select

    Next I

    NoRes = True
    UpToDate = True
    Call AoUpdate(UpToDate, NoRes)

End Sub

''
' Runs AoUpdate if we haven't updated yet, patches aoupdate and runs Client normally if we are updated.
'
' @param UpToDate Specifies if we have checked for updates or not
' @param NoREs Specifies if we have to set nores arg when running the client once again (if the AoUpdate is executed).

Private Sub AoUpdate(ByVal UpToDate_ As Boolean, ByVal NoRes_ As Boolean)

    '*************************************************
    'Author: BrianPr
    'Created: 25/11/2008
    'Last modified: 25/11/2008
    '
    '*************************************************
    Dim extraArgs  As String

    Dim Reintentos As Integer

    If Not UpToDate_ Then

        'No recibe update, ejecutar AU
        'Ejecuto el AoUpdate, sino me voy
        If Dir(App.path & "\AoUpdate.exe", vbArchive) = vbNullString Then
            MsgBox "No se encuentra el archivo de actualización AoUpdate.exe por favor descarguelo y vuelva a intentar", vbCritical
            End
        Else
Reintentar:

            On Error GoTo Error

            'FileCopy App.path & "\AoUpdate.exe", App.path & "\AoUpdateTMP.exe"
            If NoRes_ Then
                extraArgs = " /nores"

            End If
            
            Call ShellExecute(0, "Open", App.path & "\AoUpdate.exe", App.EXEName & ".exe", App.path, SW_SHOWNORMAL)
            'Call Shell(App.path & "\AoUpdateTMP.exe", App.EXEName & ".exe")
            End
            Exit Sub

        End If

    Else

        If FileExist(App.path & "\AoUpdateTMP.exe", vbArchive) Then Kill App.path & "\AoUpdateTMP.exe"

    End If

    Exit Sub

Error:

    If err.Number = 75 Then 'Si el archivo AoUpdateTMP.exe está en uso, entonces esperamos 5 ms y volvemos a intentarlo hasta que nos deje.
        Reintentos = Reintentos + 1

        If Reintentos = 3 Then
            Call MsgBox("El proceso AoUpdateTMP.exe se encuentra abierto o protegido y no es posible reemplazarlo. Cierre el proceso y vuelva a ejecutar el juego.", vbError)
            End
        Else
            Sleep 500
            GoTo Reintentar:

        End If
        
    Else
        MsgBox err.Description & vbCrLf, vbInformation, "[ " & err.Number & " ]" & " Error "
        End

    End If

End Sub

Private Sub LoadClientSetup()

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 24/06/2006
    '
    '**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open PathInit & "\ao.dat" For Binary Access Read Lock Write As fHandle
    Get fHandle, , ClientSetup
    Close fHandle
    
    NoRes = ClientSetup.bNoRes
    
    ClientSetup.WinSock = True
    
    GraphicsFile = "Graficos.ind"

End Sub

Private Sub SaveClientSetup()
    '**************************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 03/11/10
    '
    '**************************************************************
    'Dim fHandle As Integer
    
    'fHandle = FreeFile
    
    'ClientSetup.bNoMusic = Not Audio.MusicActivated
    'ClientSetup.bNoSound = Not Audio.SoundActivated
    'ClientSetup.bNoSoundEffects = Not Audio.SoundEffectsActivated
    'ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
    'ClientSetup.bGldMsgConsole = Not DialogosClanes.Activo
    'ClientSetup.bCantMsgs = DialogosClanes.CantidadDialogos
    
    'Open PathInit & "\AO.dat" For Binary As fHandle
    'Put fHandle, , ClientSetup
    'Close fHandle
End Sub

Private Sub InicializarNombres()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
    '**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cArkhein) = "Arkhein"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    Ciudades(eCiudad.cLindos) = "Lindos"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"

End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Removes all text from the console and dialogs
    '**************************************************************
    'Clean console and dialogs
    LineasConsola = 0
    OffSetConsola = 0
    ZonaActual = 0
    CambioZona = 0
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs

End Sub

Public Sub CloseSock()

    If Not ClientSetup.WinSock Then
        frmMain.Client.CloseSck
    Else
        frmMain.WSock.Close

    End If

End Sub

Public Sub CloseClient()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    CloseSock
    
    EngineRun = False
    frmCargando.Show

    'Call SaveClientSetup
    
    'Stop tile engine
    Call LiberarObjetosDX
    
    'Destruimos los cursores
    Call DestCursor
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    Set ParticlesORE = Nothing
    Set Barco(0) = Nothing
    Set Barco(1) = Nothing
    Set curGeneral = Nothing
    Set curGeneralCrimi = Nothing
    Set curGeneralCiuda = Nothing
    Set curProyectil = Nothing
    Set curProyectilPequena = Nothing
    
    UserEmbarcado = False
    
    #If SeguridadAlkon Then
        Set MD5 = Nothing
    #End If

    Call Audio.StopMidi
    
    Call UnloadAllForms
    
    'Si se cambio la resolucion, la reseteamos.
    If ResolucionCambiada Then Resolution.ResetResolution
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
    
    End

End Sub

Public Function BuscarZona(ByVal X As Integer, ByVal Y As Integer) As Integer

    Dim I        As Integer

    Dim Encontro As Boolean

    Dim NewMidi  As Integer

    Encontro = False

    For I = 1 To NumZonas

        If UserMap = Zonas(I).Mapa And X >= Zonas(I).x1 And X <= Zonas(I).x2 And Y >= Zonas(I).y1 And Y <= Zonas(I).y2 Then
            BuscarZona = I
            Encontro = True

            If Zonas(I).Acoplar = 0 Then Exit For

        End If

    Next I

    If Not Encontro And UserMap > 0 Then
        I = IIf(HayAgua(X, Y), 24, 23)
        BuscarZona = I

    End If

End Function

Public Sub CheckZona()

    Dim I        As Integer

    Dim Encontro As Boolean

    Dim NewMidi  As Integer

    Encontro = False

    For I = 1 To NumZonas

        If UserMap = Zonas(I).Mapa And UserPos.X >= Zonas(I).x1 And UserPos.X <= Zonas(I).x2 And UserPos.Y >= Zonas(I).y1 And UserPos.Y <= Zonas(I).y2 Then
            If ZonaActual <> I Then
                If ZonaActual > 0 Then
                    If Zonas(ZonaActual).Segura <> Zonas(I).Segura Then
                        CambioSegura = True
                    Else
                        CambioSegura = False

                    End If

                Else
                    CambioSegura = True

                End If

                ZonaActual = I
            
            End If

            Encontro = True

            If Zonas(I).Acoplar = 0 Then Exit For

        End If

    Next I

    If Not Encontro And UserMap > 0 Then
        I = IIf(HayAgua(UserPos.X, UserPos.Y), 24, 23)

        If ZonaActual <> I Then
            ZonaActual = I

        End If

    End If

    If ZonaActual > 0 Then
        If LastZona <> Zonas(ZonaActual).nombre Then
            CambioZona = 500

            ' ver ReyarB
            If ZonaActual = 11 Then 'Bosque Dork
                nAlpha = 0

            End If
        
            If Zonas(ZonaActual).CantMusica > 0 Then
                NewMidi = Zonas(ZonaActual).Musica(RandomNumber(1, Zonas(ZonaActual).CantMusica))

                If NewMidi <> MidiCambio Then
                    'MidiCambio = NewMidi
                    Audio.MusicVolume = mOpciones.VolMusic
                    Call Audio.PlayBackgroundMusic(NewMidi, MusicTypes.Midi)

                End If

                'SinMidi = False
            Else
                'SinMidi = True
                Call Audio.StopMidi
                Audio.MusicVolume = 0

            End If

            LastZona = Zonas(ZonaActual).nombre

        End If

    End If

End Sub

Sub ClosePj()

'Stop audio
    Dim I As Integer

    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone

    Dim X As Integer

    Dim Y As Integer

    For X = 1 To XMaxMapSize
        For Y = 1 To YMaxMapSize
            MapData(X, Y).CharIndex = 0

            If MapData(X, Y).ObjGrh.GrhIndex = GrhFogata Then
                MapData(X, Y).Graphic(3).GrhIndex = 0
                Call Light_Destroy_ToMap(X, Y)

            End If

            MapData(X, Y).ObjGrh.GrhIndex = 0
        Next Y
    Next X

    On Local Error Resume Next
    frmMain.SendTxt.Visible = False
    frmMain.SendCMSTXT.Visible = False

    FrameUseMotionBlur = False
    TiempoHome = 0
    GoingHome = 0
    AngMareoMuerto = 0
    RadioMareoMuerto = 0
    BlurIntensity = 255
    ZoomLevel = 0
    'D3DDevice.SetRenderTarget pBackbuffer, DeviceStencil, 0

    For I = 0 To Forms.count - 1

        If Forms(I).Name <> frmMain.Name And Forms(I).Name <> frmCrearPersonaje.Name And Forms(I).Name <> frmMensaje.Name Then
            Unload Forms(I)

        End If

    Next I

    'Show connection form
    If Not frmCrearPersonaje.Visible And Not Conectar Then
        ShowConnect

    End If

    'Reset global vars
    UserDescansar = False
    UserParalizado = False
    pausa = False
    UserCiego = False
    UserMeditar = False
    UserNavegando = False
    UserEmbarcado = False
    Set Barco(0) = Nothing
    Set Barco(1) = Nothing
    bRain = False

    bFogata = False
    SkillPoints = 0
    TiempoRetos = 0

    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    frmMain.macrotrabajo.Enabled = False

    'Delete all kind of dialogs
    Call CleanDialogs
    If Dir(App.path & "\INIT\" & UserName & ".rtf", vbArchive) <> "" Then
      
        Kill (App.path & "\INIT\" & UserName & ".rtf")
    
    End If
    'Reset some char variables...
    For I = 1 To LastChar
        charlist(I).invisible = False
    Next I

    'Unload all forms except frmMain
    Dim Frm As Form

    For Each Frm In Forms

        If Frm.Name <> frmMain.Name Then
            Unload Frm

        End If

    Next

    DoConectar

End Sub

Public Function General_Distance_Get(ByVal x1 As Integer, _
                                     ByVal y1 As Integer, _
                                     x2 As Integer, _
                                     y2 As Integer) As Integer

    Dim Dist As Long

    Dist = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    General_Distance_Get = Dist

End Function

Public Sub mOpciones_Default()
    mOpciones.Music = True
    mOpciones.sound = True
    mOpciones.VolMusic = 80
    mOpciones.VolSound = 100
    mOpciones.DialogConsole = False
    mOpciones.DialogCantMessages = 5
    mOpciones.GuildNews = False
    mOpciones.ScreenShooterNivelSuperior = False
    mOpciones.ScreenShooterNivelSuperiorIndex = 15
    mOpciones.ScreenShooterAlMorir = False
    mOpciones.SoundEffects = True
    mOpciones.TransparencyTree = True
    mOpciones.Shadows = True
    mOpciones.BlurEffects = True
    mOpciones.Niebla = True
    mOpciones.CursorFaccionario = True
    
    'Actualizamos el Cursor
    Select Case UserFaccion

        Case 2
            curGeneralCiuda.CursorOn frmMain.hwnd
            curGeneralCiuda.CursorOn frmMain.pRender.hwnd

        Case 1
            curGeneralCrimi.CursorOn frmMain.hwnd
            curGeneralCrimi.CursorOn frmMain.pRender.hwnd

        Case Else
            curGeneral.CursorOn frmMain.hwnd
            curGeneral.CursorOn frmMain.pRender.hwnd

    End Select
        
End Sub

Public Sub checkText(ByVal Text As String)

    Dim Nivel As Integer

    'If Right(Text, Len(MENSAJE_FRAGSHOOTER_TE_HA_MATADO)) = MENSAJE_FRAGSHOOTER_TE_HA_MATADO Then
    '    ScreenShooterCapturePending = True
    '    Exit Sub
    'End If
    If left(Text, Len(MENSAJE_FRAGSHOOTER_HAS_MATADO)) = MENSAJE_FRAGSHOOTER_HAS_MATADO Then
        ScreenShooterCapturePending = True
        FragShooterEsperandoLevel = True
        Exit Sub

    End If

    If FragShooterEsperandoLevel Then
        If right(Text, Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA)) = MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA Then
            If CInt(mid(Text, Len(MENSAJE_FRAGSHOOTER_HAS_GANADO), (Len(Text) - (Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA) + Len(MENSAJE_FRAGSHOOTER_HAS_GANADO))))) / 2 > mOpciones.ScreenShooterNivelSuperiorIndex Then
                ScreenShooterCapturePending = True

            End If

        End If

    End If

    FragShooterEsperandoLevel = False

End Sub

Public Sub DestCursor()

    '    curGeneral.CursorOff frmMain.hwnd
    '    curGeneral.CursorOff frmMain.pRender.hwnd
    '    curGeneralCiuda.CursorOff frmMain.hwnd
    '    curGeneralCiuda.CursorOff frmMain.pRender.hwnd
    '    curGeneralCrimi.CursorOff frmMain.hwnd
    '    curGeneralCrimi.CursorOff frmMain.pRender.hwnd
    '    curGeneral.CursorOff frmMain.hwnd
    '    curGeneral.CursorOff frmMain.pRender.hwnd
    '    curProyectil.CursorOff frmMain.pRender.hwnd
    '    curProyectilPequena.CursorOff frmMain.pRender.hwnd
End Sub

Public Sub SetCursor(ByVal tCursor As eCursor)

    Select Case tCursor

        Case eCursor.General

            If mOpciones.CursorFaccionario = False Then
                curGeneral.CursorOn frmMain.hwnd
                curGeneral.CursorOn frmMain.pRender.hwnd
            Else

                Select Case UserFaccion

                    Case 2
                        curGeneralCiuda.CursorOn frmMain.hwnd
                        curGeneralCiuda.CursorOn frmMain.pRender.hwnd

                    Case 1
                        curGeneralCrimi.CursorOn frmMain.hwnd
                        curGeneralCrimi.CursorOn frmMain.pRender.hwnd

                    Case Else
                        curGeneral.CursorOn frmMain.hwnd
                        curGeneral.CursorOn frmMain.pRender.hwnd

                End Select

            End If

        Case eCursor.proyectil
            curProyectil.CursorOn frmMain.pRender.hwnd

        Case eCursor.ProyectilPequena
            curProyectilPequena.CursorOn frmMain.pRender.hwnd

    End Select

End Sub

Public Function getTagPosition(ByVal Nick As String) As Integer
    
    Dim buf As Integer

    buf = InStr(Nick, "<")

    If buf > 0 Then
        getTagPosition = buf
        Exit Function

    End If
    
    buf = InStr(Nick, "[")

    If buf > 0 Then
        getTagPosition = buf
        Exit Function

    End If
    
    getTagPosition = Len(Nick) + 2
    
End Function

'Particulas *****************
'*****************************
Function Engine_ElapsedTime() As Long
 
    '**************************************************************
    'Gets the time that past since the last call
    '**************************************************************
 
    Dim Start_Time As Long
 
    'Get current time
 
    Start_Time = timeGetTime
 
    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - End_Time
 
    'Get next end time
    End_Time = Start_Time
 
End Function
 
Public Function Engine_GetAngle(ByVal CenterX As Integer, _
                                ByVal CenterY As Integer, _
                                ByVal TargetX As Integer, _
                                ByVal TargetY As Integer) As Single
 
    '************************************************************
    'Gets the angle between two points in a 2d plane
    '************************************************************
 
    On Error GoTo ErrOut

    Dim SideA As Single

    Dim SideC As Single
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270

        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180

        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
    Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
    Exit Function
 
End Function
 
Public Sub Engine_Init_RenderStates()
 
    '************************************************************
    'Set the render states of the Direct3D Device
    'This is in a seperate sub since if using Fullscreen and device is lost
    'this is eventually called to restore settings.
    '************************************************************
    'Set the shader to be used
 
    D3DDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
 
    'Set the render states
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
 
    'Particle engine settings
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
 
    'Set the texture stage stats (filters)
    '//D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    '//D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
 
End Sub
 
Public Sub Engine_Init_ParticleEngine()
 
    '*****************************************************************
    'Loads all particles into memory - unlike normal textures, these stay in memory. This isn't
    'done for any reason in particular, they just use so little memory since they are so small
    '*****************************************************************
 
    Dim I As Byte
 
    'Set the particles texture
 
    NumEffects = 20
    ReDim Effect(1 To NumEffects)
 
    For I = 1 To UBound(ParticleTexture())
        Set ParticleTexture(I) = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\Recursos\" & "p" & I & ".png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
    Next I
 
End Sub
 
Function Engine_PixelPosX(ByVal X As Long) As Long
    '*****************************************************************
    'Converts a tile position to a screen position
    'More info: [url=http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosX]http://www.vbgore.com/GameClient.TileEn ... _PixelPosX[/url]
    '*****************************************************************
 
    Engine_PixelPosX = (X - 1) * TilePixelWidth
 
End Function
 
Function Engine_PixelPosY(ByVal Y As Long) As Long
    '*****************************************************************
    'Converts a tile position to a screen position
    'More info: [url=http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosY]http://www.vbgore.com/GameClient.TileEn ... _PixelPosY[/url]
    '*****************************************************************
 
    Engine_PixelPosY = (Y - 1) * TilePixelHeight
 
End Function
 
Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
 
    '************************************************************
    'Tile Position to Screen Position
    'Takes the tile position and returns the pixel location on the screen
    '************************************************************
    
    'This acts just as a dummy in this project
 
End Function
 
Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
 
    '************************************************************
    'Tile Position to Screen Position
    'Takes the tile position and returns the pixel location on the screen
    '************************************************************
 
    'This acts just as a dummy in this project
    
End Function

'Particulas
'Particulas *********************

Sub CargarAlas()
 
' \ Author : maTih.-
' \ Note   : Llena el array de alas y los grhIndex
 
alaPath = App.path & "\INIT\Alas.ini"
 
Dim cantidadAlas As Byte
Dim loopX        As Long
 
cantidadAlas = Val(GetVar(alaPath, "INIT", "Cantidad"))
 
If cantidadAlas = 0 Then Exit Sub
 
ReDim alaArray(1 To cantidadAlas) As Alas
 
For loopX = 1 To cantidadAlas
    
    CargarAlaIndex loopX
    
Next loopX
 
End Sub

Sub CargarAlaIndex(ByVal alaIndex As Byte)



    Dim loopX As Long

    With alaArray(alaIndex)

        For loopX = E_Heading.north To E_Heading.west
            InitGrh .GrhIndex(loopX), Val(GetVar(alaPath, "ALA" & alaIndex, "Direccion" & loopX)), 0
        Next loopX

    End With

End Sub

Sub AddtoRichTextBox(RichTextBox As RichTextBox, Text As String, Optional red As Integer = -1, Optional green As Integer, Optional blue As Integer, Optional bold As Boolean, Optional italic As Boolean, Optional bCrLf As Boolean)

    
    With RichTextBox
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0

        .SelBold = IIf(bold, True, False)
        .SelItalic = IIf(italic, True, False)

        If Not red = -1 Then .SelColor = RGB(red, green, blue)

        .SelText = IIf(bCrLf, Text, Text & vbCrLf)

        RichTextBox.Refresh
    End With

End Sub
