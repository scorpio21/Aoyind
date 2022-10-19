Attribute VB_Name = "ModBots"
Option Explicit
 
'Defensa del bot jeje
Private Const IA_MINDEF  As Integer = 10
Private Const IA_MAXDEF  As Integer = 12
 
 Public Const MAX_BOTS   As Byte = 25
 
'Charindex reservado.
Private Const IA_CHAR    As Integer = (MAXCHARS - MAX_BOTS)
 
'Datos del char
 
'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
 
'ATENCION : Ac� van los n�meros de objetos!!!
 
Private Const IA_HEAD    As Integer = 4
Private Const IA_BODY    As Integer = 512
 
 
'ATENCION : Ac� van los n�meros de objetos!!!
 
'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
 
'Cantidad de hechizos que lanza
 
Private Const IA_M_SPELL As Byte = 3
Private Const IA_NUMCHAT As Byte = 5
 
'Constantes de intervalos.
 
Private Const IA_SINT   As Integer = 800    'Intervalo entre hechizo-hechizo.
Private Const IA_SREMO  As Integer = 500    'Intervalo remo.
Private Const IA_MOVINT As Integer = 500    'Intervalo caminta.
Private Const IA_USEOBJ As Integer = 200    'Intervalo usar potas.
Private Const IA_HITINT As Integer = 200    'Intervalo para golpe
Private Const IA_PROINT As Integer = 700    'Intervalo de flecha
Private Const IA_TALKIN As Integer = 4000   'Intervalo de hablAR :P
 
'Probabilidades de que te pegue
 
Private Const IA_CASTEO As Byte = 77
 
Private Const IA_PROBEV As Byte = 160
Private Const IA_PROBEX As Byte = 220
 
Private Const IA_SLOTS  As Byte = 20
 
Type ia_Interval
     SpellCount         As Byte         'Intervalo para tirar hechizos.
     UseItemCount       As Byte         'Intervalo para usar pociones.
     MoveCharCount      As Byte         'Intervalo para mover el char.
     ParalizisCount     As Byte         'Intervalo para removerse.
     HitCount           As Byte         'Intervalo para pegar golpesito.
     ArrowCount         As Byte         'Intervalo para flechas
     ChatCount          As Byte         'INtervalo para hablar XD
End Type
 
Type ia_Spells
     DamageMin          As Byte         'Minimo da�o que hace.
     DamageMax          As Byte         'Maximo da�o que hace.
     spellIndex         As Byte         'Lo usamos para el fx.
End Type
 
Enum eIASupportActions
     SRemover = 1                       'Remueve.
     SCurar = 2                         'Cura.
End Enum
 
Enum eIAClase
     Clerigo = 1                        'Bot Clero
     Mago = 2                           'Bot Mago
     Cazador = 3                        'Bot kza
End Enum
 
Enum eIAactions
     ePegar = 1                          'accion pegar.
     eMagia = 2                          'atacar con hechizo
End Enum
 
Enum eIAMoviments
     SeguirVictima = 1                   'Si seguia la victima
     MoverRandom = 2                     'Random moviment :P
End Enum
 
Type Bot
     GrupoID            As Integer
     esCriminal         As Boolean
     Pos                As WorldPos     'Posicion en el mundo.
     maxVida            As Integer      'Maxima vida.
     Vida               As Integer      'Vida del bot.
     clase              As eIAClase     'Clases de bot.
     Tag                As String       'Tag del bot.
     Mana               As Integer      'Mana del bot.
     maxMana            As Integer      'Maxima mana
     Char               As Char         'Apariencia.
     Invocado           As Boolean      'Si existe.
     Paralizado         As Boolean      'Si est� inmo.
     Intervalos         As ia_Interval  'Intervalos de acciones.
     Viajante           As Boolean      'Bot Viajante :P
     ViajanteUser       As Integer      'Usuario que atac� al viajante.
     UltimaAccion       As eIAactions   'ULTIMA ACCION/ATAQUE.
     UltimoMovimiento   As eIAMoviments 'ULTIMO MOVIMIENTO
     Navegando          As Boolean      'Navegando?
     ViajanteAntes      As WorldPos     'Pos cuando un viajante ataca un usuario.
     Inv(1 To IA_SLOTS) As Obj          'Inventario del bot.
     UltimaIdaObjeto    As Boolean      'Ultimo movimiento fue buscar objs?
End Type
 
Public ia_Bot(1 To MAX_BOTS)           As Bot
Public ia_spell(1 To IA_M_SPELL)       As ia_Spells
Public ia_Chats(1 To IA_NUMCHAT)      As String
 
'Cantidad de bots invocados.
Public NumInvocados                    As Byte
 
Function ia_CascoByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Devuelve el casco/gorro seg�n la clase del bot
 
Select Case ia_Bot(BotIndex).clase
 
       Case eIAClase.Clerigo        'Bot clero
            ia_CascoByClase = 131   'Completo : P
        
       Case eIAClase.Mago           'Bot mago.
            ia_CascoByClase = 662   'Vara
            
       Case eIAClase.Cazador        'Bot kza
            ia_CascoByClase = 405   'de plata
        
End Select
 
End Function
 
Function ia_EscudoByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Devuelve el escudo seg�n la clase del bot
 
Select Case ia_Bot(BotIndex).clase
 
       Case eIAClase.Clerigo        'Bot clero
            ia_EscudoByClase = 130  'De plata.
        
       Case eIAClase.Mago           'Bot mago.
            ia_EscudoByClase = -1   'Nada
            
       Case eIAClase.Cazador        'bot kaza
            ia_EscudoByClase = 404  'escudo d tortu
        
End Select
 
End Function
 
Function ia_ArmaByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Devuelve el arma seg�n la clase del bot
 
Select Case ia_Bot(BotIndex).clase
 
       Case eIAClase.Clerigo        'Bot clero
            ia_ArmaByClase = 129    'Dos filos : P
        
       Case eIAClase.Mago           'Bot mago.
            ia_ArmaByClase = 400    'Vara
            
       Case eIAClase.Cazador        'bot cazador
            ia_ArmaByClase = 665    'arko de kza
        
End Select
 
End Function
 
Function ia_VidaByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Devuelve la vida seg�n la clase.
 
Select Case ia_Bot(BotIndex).clase
       Case eIAClase.Clerigo        '<Clerigo.
            'Vida random. (de clerigo 41)
            ia_VidaByClase = 21 + (RandomNumber(8, 10) * 41)
        
       Case eIAClase.Mago           '<Mago
            'Vida random (de mago 39)
            ia_VidaByClase = 21 + (RandomNumber(7, 9) * 39)
            
       Case eIAClase.Cazador        '<Kza
            'Vida random de cazador humano 42
            ia_VidaByClase = 21 + (RandomNumber(8, 11) * 42)
            
End Select
 
End Function
 
Function ia_ManaByClase(ByVal BotIndex As Byte) As Integer
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Devuelve man� seg�n la clase.
 
Select Case ia_Bot(BotIndex).clase
       Case eIAClase.Clerigo        '<Clerigo.
            'Mana de clero 41 : P
            ia_ManaByClase = 1480
        
       Case eIAClase.Mago           '<Mago
            'Mana de mago 39 : P
            ia_ManaByClase = 1954
            
       Case eIAClase.Cazador        'caza sin mana
            ia_ManaByClase = 0
            
End Select
 
End Function
 
Function ia_CalcularGolpe(ByVal VictimIndex As Integer) As Integer
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Calcula el golpe (da�o) q hace el bot al user.
 
Dim ParteCuerpo     As Integer
Dim Da�oAbsorvido   As Integer
 
ParteCuerpo = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
 
'Si pega en la cabeza.
If ParteCuerpo = PartesCuerpo.bCabeza Then
   'Si tiene casco baja el golpe
       If UserList(VictimIndex).Invent.CascoEqpObjIndex <> 0 Then
          Da�oAbsorvido = RandomNumber(ObjData(UserList(VictimIndex).Invent.CascoEqpObjIndex).MinDef, ObjData(UserList(VictimIndex).Invent.CascoEqpObjIndex).MaxDef)
       End If
Else
    'Se fija por la armadura.
       If UserList(VictimIndex).Invent.ArmourEqpObjIndex <> 0 Then
          Da�oAbsorvido = RandomNumber(ObjData(UserList(VictimIndex).Invent.ArmourEqpObjIndex).MinDef, ObjData(UserList(VictimIndex).Invent.ArmourEqpObjIndex).MaxDef)
       End If
End If
       
'DEVUELVE.
ia_CalcularGolpe = (RandomNumber(150, 180) - Da�oAbsorvido)
       
End Function
 
Function ia_AciertaGolpe(ByVal VictimIndex As Integer) As Boolean
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Evasi�n del usuario, esto le faltan unos retoques.
 
Dim tempEvasion     As Long
Dim tempEvasionEsc  As Long
Dim tempResultado   As Long
 
'Evasi�n del usuario.
tempEvasion = PoderEvasion(VictimIndex)
 
'Evasi�n del usuario con escudos.
'Tiene escudo?
If UserList(VictimIndex).Invent.EscudoEqpObjIndex <> 0 Then
    tempEvasionEsc = PoderEvasionEscudo(VictimIndex)
    tempEvasionEsc = tempEvasion + tempEvasionEsc
Else
    tempEvasionEsc = 0
End If
 
'Acierta?
tempResultado = MaximoInt(10, MinimoInt(90, 50 + (IA_PROBEX - tempEvasion) * 0.4))
 
'Random.
ia_AciertaGolpe = (RandomNumber(1, 100) <= tempResultado)
 
End Function
 
Function ia_PuedeMeele(ByRef PosBot As WorldPos, ByRef PosVictim As WorldPos, ByRef NewHeading As eHeading) As Boolean
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Se fija si est� al lado, y guarda el heading.
 
With PosVictim
    
    'Mirando hacia la derecha lo tiene ?
    If PosBot.X + 1 = .X Then
       ia_PuedeMeele = (.Y = PosBot.Y)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.EAST
       End If
       
       Exit Function
    End If
    
    'mirando hacia izq?
    If PosBot.X - 1 = .X Then
       ia_PuedeMeele = (.Y = PosBot.Y)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.WEST
       End If
       
       Exit Function
    End If
    
    'mirando arriba
    If PosBot.Y - 1 = .Y Then
       ia_PuedeMeele = (.X = PosBot.X)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.NORTH
       End If
       
       Exit Function
    End If
    
    'Abajo.
    If PosBot.Y + 1 = .Y Then
       ia_PuedeMeele = (PosBot.X = .X)
       
       If ia_PuedeMeele Then
          NewHeading = eHeading.SOUTH
       End If
       
       Exit Function
    End If
    
End With
 
End Function
 
Sub ia_CreateChar(ByVal ProximoBot As Byte)
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Crea el char.
 
Dim PackageToSend   As String
 
With ia_Bot(ProximoBot).Char
 
    .Body = ObjData(IA_BODY).Ropaje
    .Head = IA_HEAD
    
    'Siempre tienen arma.
    .WeaponAnim = ObjData(ia_ArmaByClase(ProximoBot)).WeaponAnim
    
    'Escudo no, me fijo si tienen..
    If ia_EscudoByClase(ProximoBot) <> -1 Then
        .ShieldAnim = ObjData(ia_EscudoByClase(ProximoBot)).ShieldAnim
    End If
    
    'Casco si..
    .CascoAnim = ObjData(ia_CascoByClase(ProximoBot)).CascoAnim
    
    'Precalculado : P
    .CharIndex = IA_CHAR + ProximoBot
    
    'Preparo el paquete de datos.
    
            Dim tmp_Color   As Byte
            
            If ia_Bot(ProximoBot).esCriminal Then
               tmp_Color = 1
            Else
               tmp_Color = 0
            End If
    
    PackageToSend = PrepareMessageCharacterCreate(.Body, .Head, eHeading.SOUTH, .CharIndex, ia_Bot(ProximoBot).Pos.X, ia_Bot(ProximoBot).Pos.Y, .WeaponAnim, .ShieldAnim, 0, 0, .CascoAnim, ia_Bot(ProximoBot).Tag, tmp_Color, 0, 0, 0)
    
    'Actualizo el area.
    ia_SendToBotArea ProximoBot, PackageToSend
    
End With
 
End Sub
 
Public Function ia_Spawn(ByVal iaClase As eIAClase, ByRef PosToSpawn As WorldPos, ByRef BotTag As String, ByVal Viajante As Boolean, ByVal esPk As Boolean, ByVal g_ID As Integer) As Integer
 
' @designer     : Walyto132
' @date         : 2021/06/05
 
Dim ProximoBot  As Byte
Dim PackageSend As String
 
ProximoBot = IA_GetNextSlot
 
If Not ProximoBot <> 0 Then Exit Function
 
With ia_Bot(ProximoBot)
    
    .Invocado = True
    
    .clase = iaClase
    
    .GrupoID = g_ID
    
    .Mana = ia_ManaByClase(ProximoBot)
    .Vida = ia_VidaByClase(ProximoBot)
    .maxMana = .Mana
    .maxVida = .Vida
    
    .esCriminal = esPk
    
    'Si es "viajante"..
    .Viajante = Viajante
    
    .Tag = BotTag
    
    .Paralizado = False
    
    'Seteo la posici�n.
    .Pos = PosToSpawn
    
    'Creo el char.
    ia_CreateChar ProximoBot
   
    'Primer action ! : D
    ia_Action ProximoBot
   
    'PackageSend = PrepareMessageChatOverHead("VeNGan PutOs xD!", .Char.CharIndex, vbCyan)
   
    'ia_SendToBotArea ProximoBot, PackageSend
   
    .Intervalos.SpellCount = 100
   
    NumInvocados = NumInvocados + 1
    
    MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = ProximoBot
    
    'devuelvo el id del bot
    ia_Spawn = ProximoBot
   
End With
 
End Function
 
Public Sub ia_Spells()
 
' @designer     : Walyto132
' @date         : 2021/06/05
 
'Un poco hardcodeado pero bueno :D
 
'Hechizo 1 : descarga.
ia_spell(1).DamageMax = 120
ia_spell(1).DamageMax = 177
ia_spell(1).spellIndex = 23
 
'Hechizo 2 : apoca
 
ia_spell(2).DamageMin = 190
ia_spell(2).DamageMax = 220
ia_spell(2).spellIndex = 25
 
'Paralizar.
ia_spell(3).DamageMax = 0
ia_spell(3).DamageMin = 0
ia_spell(3).spellIndex = 9
 
ia_Chats(1) = "LKA J�KKLJA K�L JA�LK�LKAJ LK AJ"
ia_Chats(2) = "NEGRO"
ia_Chats(3) = "CHAU CHE"
ia_Chats(4) = "NANANANA TE ISE PAPIYAA"
ia_Chats(5) = "HIJO DE PUTA"
 
End Sub
 
Sub ia_RandomMoveChar(ByVal BotIndex As Byte, ByVal siguiendoIndex As Integer, ByRef HError As Boolean)
 
' @designer     : Walyto132
' @date         : 2021/06/05
 
With ia_Bot(BotIndex)
 
    Dim nRandom     As Byte
   
    '25% De probabilidades de moverse a
    'cualquiera de las cuatro direcciones.
   
    nRandom = RandomNumber(1, 4)
   
    Select Case nRandom
   
           Case 1
           'Dim verificar As Boolean
         '   verificar = ia_LegalPos(.Pos.X + 1, .Pos.Y, BotIndex, siguiendoIndex)
               If ia_LegalPos(.Pos.X + 1, .Pos.Y, BotIndex, siguiendoIndex) = True Then HError = True: Exit Sub
                
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = 0
                
                .Pos.X = .Pos.X + 1
           
           Case 2
           
                If ia_LegalPos(.Pos.X - 1, .Pos.Y, BotIndex, siguiendoIndex) = True Then HError = True: Exit Sub
           
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = 0
                
                .Pos.X = .Pos.X - 1
           
           Case 3
           
                If ia_LegalPos(.Pos.X, .Pos.Y + 1, BotIndex, siguiendoIndex) = False Then HError = True: Exit Sub
           
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = 0
               
                .Pos.Y = .Pos.Y + 1
           
           Case 4
           
                If ia_LegalPos(.Pos.X, .Pos.Y - 1, BotIndex, siguiendoIndex) = True Then HError = True: Exit Sub
                
                'Borro el BotIndex del tile anterior.
                MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = 0
               
                .Pos.Y = .Pos.Y - 1
   
    End Select
 
End With
 
End Sub
 
Sub ia_CargarRutas(ByRef MAPFILE As String, ByVal MapIndex As Integer)
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @modificated  :  Carga las rutas de un mapa : D
 
Dim loopX   As Long
Dim loopY   As Long
Dim tmpVal  As eHeading
 
For loopX = 1 To 100
    For loopY = 1 To 100
        
        tmpVal = val(GetVar(MAPFILE, CStr(loopX) & "," & CStr(loopY), "Direccion"))
        
        If tmpVal <> 0 Then
          ' MapData(MapIndex, loopX, loopY).Rutas(1) = tmpVal
        End If
        
    Next loopY
Next loopX
 
End Sub
 
Function ia_LegalPos(ByVal X As Long, ByVal Y As Long, ByVal BotIndex As Byte, Optional ByVal siguiendoUser As Integer) As Boolean
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @modificated  :  Esta funci�n ya no trabaja con la pos del npc si no que ahora usa los par�metros.
 
ia_LegalPos = False
 
With MapData(ia_Bot(BotIndex).Pos.map).Tile(X, Y)
 
     '�Es un mapa valido?
    If (ia_Bot(BotIndex).Pos.map <= 0 Or ia_Bot(BotIndex).Pos.map > NumMaps) Or ia_Bot(BotIndex).Pos.map < 1 Or X > MapInfo(ia_Bot(BotIndex).Pos.map).Width Or Y < 1 Or Y > MapInfo(ia_Bot(BotIndex).Pos.map).Height Then Exit Function
 
     'Tile bloqueado?
     If .Blocked <> 0 Then Exit Function
   
     'Hay un usuario?
     If .UserIndex > 0 Then
        'Si no es un adminInvisible entonces nos vamos.
        If UserList(.UserIndex).flags.AdminInvisible <> 1 Then Exit Function
        If EsAdmin(UserList(.UserIndex).Name) = True Then Exit Function
    End If
    
  '  If EsAdmin(UserList(.UserIndex).Name) = True Then
   ' siguiendoUser = 1
   ' Exit Function
   ' End If
 
    'Hay un NPC?
    If .NpcIndex <> 0 Then Exit Function
     
    'Hay un bot?
    If .BotIndex <> 0 Then Exit Function
    
    'Siguiendo Index?
    If siguiendoUser <> 0 Then
        'V�lido para evitar el rango Y pero no su eje X.
        If Abs(Y - UserList(siguiendoUser).Pos.Y) > RANGO_VISION_Y Then Exit Function
   
        If Abs(X - UserList(siguiendoUser).Pos.X) > RANGO_VISION_X Then Exit Function
    End If
    
     ia_LegalPos = True
   
End With
 
End Function
 
Sub ia_SearchPath(ByVal BotIndex As Byte, ByRef tPos As WorldPos, ByRef findHeading As eHeading)
 
' @designer     : Walyto132
' @date         :  2012/03/13
' @                Busc� una ruta y devuelve un puntero con el heading.
 
findHeading = FindDirection(ia_Bot(BotIndex).Pos, tPos)
 
End Sub
 
Sub ia_MoveToHeading(ByVal BotIndex As Byte, ByVal toHeading As eHeading, ByRef FoundErr As Boolean)
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Mueve el char del npc hacia una posici�n.
 
'FoundErr = True
 
Select Case toHeading
 
       Case eHeading.NORTH  '<Move norte.
            'No legal pos.
            Dim legal As Boolean
            legal = ia_LegalPos(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y - 1, BotIndex)
            If Not ia_LegalPos(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y - 1, BotIndex) Then Exit Sub
            
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(BotIndex).Pos.map).Tile(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = 0
            'Set la nueva posici�n
            ia_Bot(BotIndex).Pos.Y = ia_Bot(BotIndex).Pos.Y - 1
            
       Case eHeading.EAST   '<Move este.

            'Si hay posici�n inv�lida no se peude mover.
            If Not ia_LegalPos(ia_Bot(BotIndex).Pos.X + 1, ia_Bot(BotIndex).Pos.Y, BotIndex) Then Exit Sub
            
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(BotIndex).Pos.map).Tile(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = 0
            
            'Set la nueva posici�n
            ia_Bot(BotIndex).Pos.X = ia_Bot(BotIndex).Pos.X + 1
            
       Case eHeading.SOUTH  '<Move sur.
            'Si hay posici�n inv�lida no se peude mover.
            If Not ia_LegalPos(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y + 1, BotIndex) Then Exit Sub
            
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(BotIndex).Pos.map).Tile(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = 0
            
            'Set la nueva posici�n
            ia_Bot(BotIndex).Pos.Y = ia_Bot(BotIndex).Pos.Y + 1
            
       Case eHeading.WEST   '<Move oeste.
            'Si hay posici�n inv�lida no se peude mover.
            If Not ia_LegalPos(ia_Bot(BotIndex).Pos.X - 1, ia_Bot(BotIndex).Pos.Y, BotIndex) Then Exit Sub
            
            'Se mueve, borro el anterior botIndex.
            MapData(ia_Bot(BotIndex).Pos.map).Tile(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = 0
            
            'Set la nueva posici�n
            ia_Bot(BotIndex).Pos.X = ia_Bot(BotIndex).Pos.X - 1
            
End Select
 
FoundErr = False
 
End Sub
 
 
Sub ia_MoveViajante(ByVal BotIndex As Byte, ByVal Direccion As eHeading)
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Move el viajante hacia una posici�n
 
Dim HabiaAgua As Boolean
 
With ia_Bot(BotIndex)
 
     'Hacia donde se mueve..
     Select Case Direccion
            
            Case eHeading.NORTH     'Norte.
                 MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = 0
                 .Pos.Y = .Pos.Y - 1
                 MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = BotIndex
                 
            Case eHeading.EAST      'Este.
                 MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = 0
                 .Pos.X = .Pos.X + 1
                 MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = BotIndex
            
            Case eHeading.SOUTH     'Sur.
                MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = 0
                 .Pos.Y = .Pos.Y + 1
                 MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = BotIndex
                 
            Case eHeading.WEST      'Oeste.
                 MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = 0
                 .Pos.X = .Pos.X - 1
                 MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = BotIndex
     End Select
     
     HabiaAgua = HayAgua(.Pos.map, .Pos.X, .Pos.Y)
     
     If HabiaAgua Then
        'Si hay agua cambio el cuerpo.
        ia_SendToBotArea BotIndex, PrepareMessageCharacterChange(395, 0, Direccion, .Char.CharIndex, 0, 0, 0, 0, 0, 0)
        .Navegando = True
     Else
        'No habia agua, y... estaba navegando?
        If .Navegando Then
           'cambio el body y demas.
           ia_SendToBotArea BotIndex, PrepareMessageCharacterChange(.Char.Body, .Char.Head, Direccion, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, .Char.alaIndex)
           .Navegando = False
        End If
    End If
    
     'Actualizamso
     ia_SendToBotArea BotIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y)
        
End With
 
End Sub
 
Function ia_HeadingToMolestNpc(ByVal NpcIndex As Integer) As eHeading
 
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Devuelve un heading para un npc que est� molestando el paso.
 
Dim nPos    As WorldPos
 
nPos = Npclist(NpcIndex).Pos
 
With MapData(Npclist(NpcIndex).Pos.map).Tile(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
 
     'Pos legal hacia arriba?
     If LegalPosNPC(nPos.map, nPos.X, nPos.Y - 1, 0) Then
        'Mientras no halla bot.
        If Not .BotIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.NORTH
        End If
     End If
     
     'Pos legal hacia abajo?
     If LegalPosNPC(nPos.map, nPos.X, nPos.Y + 1, 0) Then
        'Mientras no halla bot.
        If Not .BotIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.SOUTH
        End If
     End If
     
     'Pos legal hacia la izquierda?
     If LegalPosNPC(nPos.map, nPos.X - 1, nPos.Y, 0) Then
        'Mientras no halla bot.
        If Not .BotIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.WEST
        End If
     End If
     
     'Pos legal hacia la derecha?
     If LegalPosNPC(nPos.map, nPos.X + 1, nPos.Y, 0) Then
        'Mientras no halla bot.
        If Not .BotIndex <> 0 Then
           ia_HeadingToMolestNpc = eHeading.EAST
        End If
     End If
     
End With
 
End Function
 
Function ia_Objetos(ByVal BotIndex As Byte) As WorldPos
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Busca objetos valiosos en el area.
 
Dim loopX   As Long
Dim loopY   As Long
Dim BotPos  As WorldPos
 
BotPos = ia_Bot(BotIndex).Pos
 
'********************************
 
'borro esto ya que no libero esta parte : p
 
'********************************
 
ia_Objetos.map = 0
 
End Function
 
Function ia_SlotInventario(ByVal BotIndex As Byte) As Byte
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Busca un slot libre.
 
Dim loopX   As Long
 
For loopX = 1 To IA_SLOTS
    With ia_Bot(BotIndex).Inv(loopX)
         'No hay objeto.
         If Not .ObjIndex <> 0 Then
            ia_SlotInventario = CByte(loopX)
            Exit Function
         End If
    End With
Next loopX
 
ia_SlotInventario = 0
 
End Function
 
Sub ia_ActionViajante(ByVal BotIndex As Byte)
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Acciones de los bots que viajan hacia mapas.
 
Dim RutaDir     As eHeading
Dim molestNpc   As Integer
Dim ObjetoPos   As WorldPos
 
With ia_Bot(BotIndex)
 
     'Est� paralizado?
     If .Paralizado Then
        'Puede tirar hechizos.
        If .Intervalos.SpellCount = 0 Then
           'se remueve
           ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("AN HOAX VORP", .Char.CharIndex, vbCyan)
           .Paralizado = False
           .Intervalos.SpellCount = (IA_SINT / 30)
        End If
     End If
        
     'Se puede mover?
     If Not .Intervalos.MoveCharCount = 0 Then Exit Sub
        
     .Intervalos.MoveCharCount = (IA_MOVINT / 50)
     
     'Tiene una ruta?
     RutaDir = ia_HayRuta(.Pos)
    
     'Ve un objeto valioso?
     ObjetoPos = ia_Objetos(BotIndex)
     
     If ObjetoPos.map <> 0 Then
        'Lo va a buscar, pero antes , setea su vieja pos.
        If Not .UltimaIdaObjeto Then
            .ViajanteAntes = .Pos
        End If
        
        ia_SearchPath BotIndex, ObjetoPos, RutaDir
        .UltimaIdaObjeto = True
     End If
     
     'No hay ruta?
     If Not RutaDir <> 0 Then
        'habia atacado un usuario? si es as� volvemos a la pos.
        ia_SearchPath BotIndex, .ViajanteAntes, RutaDir
     End If
     
     If RutaDir <> 0 Then
        
        'Hacia donde mueve?
        Select Case RutaDir
                
               Case eHeading.NORTH      '<Mueve norte.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.map).Tile(.Pos.X, .Pos.Y - 1).NpcIndex
                    
                    #If Barcos <> 0 Then
                        If molestNpc <> 0 Then
                            ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("�Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                            Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                        End If
                    #End If
                    
               Case eHeading.SOUTH      '<Mueve sur.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.map).Tile(.Pos.X, .Pos.Y + 1).NpcIndex
                    
                    If molestNpc <> 0 Then
                       ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("�Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                       'muevo el npc
                       Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                    End If
                       
               Case eHeading.EAST       '<Mueve este.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.map).Tile(.Pos.X + 1, .Pos.Y).NpcIndex
                    
                    If molestNpc <> 0 Then
                       ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("�Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                       'muevo el npc
                       Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                    End If
                    
               Case eHeading.WEST       '<Mueve oeste.
                    'Hay npc en su camino?
                    molestNpc = MapData(.Pos.map).Tile(.Pos.X - 1, .Pos.Y).NpcIndex
                    
                    If molestNpc <> 0 Then
                       ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("�Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                       'Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                    End If
        End Select
        
        'Move:p
        ia_MoveViajante BotIndex, RutaDir
        'Set el heading.
        .Char.heading = RutaDir
     End If
     
 
     
     'Encontramos una salida? - translados.
     If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).TileExit.map <> 0 Then
        'Mapa v�lido?
        If MapaValido(MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).TileExit.map) Then
            'Asignamos nuevas posiciones, borramos el char anterior.
            ia_SendToBotArea BotIndex, PrepareMessageCharacterRemove(.Char.CharIndex)
            'Pos del npc.
            .Pos.map = MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).TileExit.map
            
            'Por si no tiene heading.
            If Not .Char.heading <> 0 Then .Char.heading = eHeading.SOUTH
            
            'Nueva X?
            If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).TileExit.X <> 0 Then
                .Pos.X = MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).TileExit.X
            End If
            
            'Nueva Y?
            If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).TileExit.Y <> 0 Then
                .Pos.Y = MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).TileExit.Y
            End If
            
             MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = BotIndex
            'Creamos.
            
            Dim tmp_Color   As eNickColor
            
            'preparo el color del nick
            If .esCriminal Then
               tmp_Color = eNickColor.ieCriminal
            Else
               tmp_Color = eNickColor.ieCiudadano
            End If
            
            ia_SendToBotArea BotIndex, PrepareMessageCharacterCreate(.Char.Body, .Char.Head, .Char.heading, .Char.CharIndex, .Pos.X, .Pos.Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, .Tag, tmp_Color, 0, 0, 0)
        End If
     End If
     
End With
 
End Sub
 
Function ia_HayRuta(ByRef InPos As WorldPos) As eHeading
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Devuelve la dircci�n de la ruta en una pos.
 
With MapData(InPos.map).Tile(InPos.X, InPos.Y)
     
     'ia_HayRuta = .Rutas(1)
     
End With
 
End Function
 
Sub ia_SupportOthers(ByVal BotIndex As Byte, ByRef Supported As Boolean)
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Un bot supportea otro.
 
Dim botIndexToSupport   As Byte
Dim supportAction       As eIASupportActions
 
'Si no tiene intervalo..
If ia_Bot(BotIndex).Intervalos.SpellCount <> 0 Then Exit Sub
 
'Busca un bot a ayudar.
botIndexToSupport = ia_GetSupportBot(BotIndex, supportAction)
 
'No encontr�, no supportea..
If Not botIndexToSupport <> 0 Then Supported = False: Exit Sub
 
'Que acci�n debe realizar?
Select Case supportAction
 
       Case eIASupportActions.SCurar        '<Cura un compa�ero
            'Lanza graves.
            'Crea fx.
            ia_SendToBotArea botIndexToSupport, Protocol.PrepareMessageCreateFX(ia_Bot(botIndexToSupport).Char.CharIndex, Hechizos(5).FXgrh, Hechizos(5).Loops)
            
            'Cartel.
            ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("EN CORP SANCTIS", ia_Bot(BotIndex).Char.CharIndex, vbCyan)
            
            'Suma un random de vida.
            ia_Bot(botIndexToSupport).Vida = ia_Bot(botIndexToSupport).maxVida + RandomNumber(55, 77)
            
            'PARA QUE NO PASE LA VIDA MAXIMA
            If ia_Bot(botIndexToSupport).Vida > ia_Bot(botIndexToSupport).maxVida Then ia_Bot(botIndexToSupport).Vida = ia_Bot(botIndexToSupport).maxVida
       
            Supported = True
       
      Case eIASupportActions.SRemover       '<Remueve paralizis.
            'Crea el fx, el remo no tiene fx.
            'ia_sendtobotarea botindextosupport
            
            'Paralizis count.
            If ia_Bot(botIndexToSupport).Intervalos.ParalizisCount > 6 Then Exit Sub
            
            'Cartel
            ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("AN HOAX VORP", ia_Bot(BotIndex).Char.CharIndex, vbCyan)
            
            'Saca el flag
            ia_Bot(botIndexToSupport).Paralizado = False
            
            Supported = True
            
End Select
 
End Sub
 
Function ia_BotEnArea(ByVal BotIndex As Byte, ByVal otherBotIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
 
Dim BotIndexPos As WorldPos
 
BotIndexPos = ia_Bot(BotIndex).Pos
 
Dim loopX   As Long
Dim loopY   As Long
 
For loopY = BotIndexPos.Y - MargenY To BotIndexPos.Y + MargenY
        For loopX = BotIndexPos.X - MargenX To BotIndexPos.X + MargenX   'hay un bot
            If MapData(BotIndexPos.map).Tile(loopX, loopY).BotIndex = otherBotIndex Then
                ia_BotEnArea = True
                Exit Function
            End If
        
        Next loopX
Next loopY
 
ia_BotEnArea = False
 
End Function
 
Function ia_GetSupportBot(ByVal BotIndex As Byte, ByRef SAction As eIASupportActions) As Byte
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Busca un bot a ayudar.
 
Dim loopX   As Long
 
For loopX = 1 To MAX_BOTS
    
    'Si no es mi BotIndex
    If loopX <> BotIndex Then
        
       'Est� invocado?
       If ia_Bot(loopX).Invocado Then
          'Est� en el area?
          If ia_BotEnArea(BotIndex, loopX) Then
             'Est� paralizado/tiene poca vida?
             If ia_Bot(loopX).Vida <> ia_Bot(loopX).maxVida Or ia_Bot(loopX).Paralizado Then
                'Encontrado.
                ia_GetSupportBot = CByte(loopX)
                'Devuelve la acci�n.
                SAction = IIf(ia_Bot(loopX).Vida <> ia_Bot(loopX).maxVida, eIASupportActions.SCurar, eIASupportActions.SRemover)
                Exit Function
             End If
          End If
       End If
       
    End If
    
Next loopX
 
ia_GetSupportBot = 0
End Function
 
Sub ia_Action(ByVal BotIndex As Byte)
 
'On Error GoTo Errhandler        '< maTih XD
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Acciones de los bots.
 
Dim pIndex      As Integer
Dim sRandom     As Integer
Dim rMan        As Integer
Dim FoundErr    As Boolean
Dim moveHeading As eHeading
Dim AyudoBot    As Boolean
 
If EnPausa Then Exit Sub
 
With ia_Bot(BotIndex)
 
    'Es un bot viajante?
    If .Viajante Then
          'Mientras no est� contra ning�n pibe
          If Not .ViajanteUser <> 0 Then
             ia_CheckInts BotIndex
             ia_ActionViajante BotIndex
             Exit Sub
          End If
    End If
    
    'si no lo ataco nadie  busca un target
    If (.ViajanteUser = 0) Then
        pIndex = ia_FindTarget(.Pos, .esCriminal)
    Else
        pIndex = .ViajanteUser
    End If
    
    'No hay usuario.
    If pIndex <= 0 Then Exit Sub
 
    'Contadores de intervalo.
    ia_CheckInts BotIndex
   
    'EL bot boquea XD
    If Not .Intervalos.ChatCount <> 0 Then
       .Intervalos.ChatCount = (IA_TALKIN / 40)
        
       'Envia msj random
       ia_SendToBotArea BotIndex, PrepareMessageChatOverHead(ia_Chats(RandomNumber(1, 5)), .Char.CharIndex, vbRed)
       .Intervalos.SpellCount = (IA_SINT / 100)
    End If
    
    'Si se puede mover AND no est� inmo se mueve al azar.
    If .Intervalos.MoveCharCount = 0 And .Paralizado = False Then
        
        'Tiene target?
        If pIndex <> 0 Then
           'busco un path.
           ia_SearchPath BotIndex, UserList(pIndex).Pos, moveHeading
        End If
        
        'Es clero?
        If Not .clase <> eIAClase.Clerigo Then
           'Si tiene la vida llena lo persigue.
           If .Vida = .maxVida Then
              ia_MoveToHeading BotIndex, moveHeading, FoundErr
           Else
            'Si no , se mueve al azar.
              ia_RandomMoveChar BotIndex, pIndex, FoundErr
           End If
         End If
                   
         'Es mago?
        If .clase = eIAClase.Mago Or .clase = eIAClase.Cazador Then
           'Si no tiene la vida llena se mueve al azar.
           If Not .Vida = .maxVida Then
              ia_RandomMoveChar BotIndex, pIndex, FoundErr
           Else
              'Tiene la vida llena, que fue el ultimo movimiento?
              'Siguio la victima?
              If .UltimoMovimiento = eIAMoviments.SeguirVictima Then
                 'Mueve random.
                 ia_RandomMoveChar BotIndex, pIndex, FoundErr
                 'Seteo.
                 .UltimoMovimiento = eIAMoviments.MoverRandom
              Else
                 'Se movi� al azar, sigue su victima.
                 ia_MoveToHeading BotIndex, moveHeading, FoundErr
                 'Seteo el nuevo flag.
                 .UltimoMovimiento = eIAMoviments.SeguirVictima
             End If
        End If
       End If
       
       'se movio.
        If Not FoundErr Then
          'Se movi�, guardo el BotIndex.
          MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = BotIndex
          
          'NEW--------
          'Checkeo si es una posici�n v�lida.
    
          'Actualizamos.
          ia_SendToBotArea BotIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y)
          
          .Intervalos.MoveCharCount = (IA_MOVINT / 40)
        End If
        
    End If
   
   
    'STATS..
   
        'Prioriza la vida ante todo
       
        If .Vida < .maxVida Then
           
            'Checkeo el intervalo.
            If .Intervalos.UseItemCount > 0 Then Exit Sub
           
            'Recupera 20 cada 200 ms.
            .Vida = .Vida + 20
           
            If .Vida > .maxVida Then .Vida = .maxVida
           
            'Uso la poci�n, seteo el interval
            .Intervalos.UseItemCount = (IA_USEOBJ / 40)
           
            Exit Sub
        End If
       
        'Si tenia la vida llena usa azules.
       
        If .Mana < .maxMana Then
       
            'Checkeo el intervalo.
           
            If .Intervalos.UseItemCount = 0 Then
            
                Dim recuperoMana    As Long
                
                'Recupera un % de la mana.
                If .clase <> eIAClase.Mago Then
                    recuperoMana = Porcentaje(.maxMana, 5)
                Else
                    recuperoMana = Porcentaje(.maxMana, 3)
                End If
                
                'aumento el mana
                .Mana = .Mana + recuperoMana
           
                'controlo el limite
                If .Mana > .maxMana Then .Mana = .maxMana
           
            'seteo el int
            .Intervalos.UseItemCount = (IA_USEOBJ / 40)
 
            End If
           
            'Hacer una constante despu�s, con esto hacemos un random
            'Para que azulee y combee a la ves.
            If RandomNumber(1, 4) < 4 Then Exit Sub
        End If
   
    'Bueno si est� ac� es por que tenia la vida y mana llenas.
     
    'Es cazador??
    If .clase = eIAClase.Cazador Then
       'Intervalo permite?
       If Not .Intervalos.ArrowCount = 0 Then Exit Sub
       'Kza manqea XD - 25% de prob fallar
       If RandomNumber(1, 100) > 65 Then Exit Sub
       'Probabilidad de evadir.
       If Not RandomNumber(1, 100) <= MaximoInt(10, MinimoInt(90, 50 + ((220 - PoderEvasion(pIndex)) * 0.4))) Then
          'Atac� y fall�!!
          Call WriteConsoleMsg(pIndex, .Tag & " Te lanz� un flechazo pero fall�!", FontTypeNames.FONTTYPE_FIGHT)
          'setea intervalo
          .Intervalos.ArrowCount = (IA_PROINT / 25)
          Exit Sub
       End If
       
       Dim ArrowDamage  As Integer  '<Da�oBase.
       Dim ArmourIndex  As Integer  '<ArmaduraObjIndex
       Dim HelmetIndex  As Integer  '<CascoObjIndex
       
       ArrowDamage = RandomNumber(185, 225)
       
       'Restamos si tiene armadura.
       ArmourIndex = UserList(pIndex).Invent.ArmourEqpObjIndex
       HelmetIndex = UserList(pIndex).Invent.CascoEqpObjIndex
       
       'Pega en cabeza?
       If RandomNumber(1, 6) = 6 Then
          'Absorve.
          If HelmetIndex <> 0 Then
             ArrowDamage = ArrowDamage - RandomNumber(ObjData(HelmetIndex).MinDef, ObjData(HelmetIndex).MaxDef)
          End If
       Else
          'Armadura absorce.
          If ArmourIndex <> 0 Then
             ArrowDamage = ArrowDamage - RandomNumber(ObjData(ArmourIndex).MinDef, ObjData(ArmourIndex).MaxDef)
          End If
       End If
       
       'crea fx.
       'SendData SendTarget.ToPCArea, pIndex, mod_DunkanProtocol.Send_CreateArrow(.Char.CharIndex, UserList(pIndex).Char.CharIndex, ObjData(553).GrhIndex)
       
       'crea da�o
       'Call mod_DunkanGeneral.Enviar_Da�oAUsuario(pIndex, ArrowDamage)
       
       'Sacude un flechazo.
       UserList(pIndex).Stats.MinHP = UserList(pIndex).Stats.MinHP - ArrowDamage
       
       Call WriteConsoleMsg(pIndex, .Tag & " Te ha pegado un flechazo por " & ArrowDamage, FontTypeNames.FONTTYPE_FIGHT)
       
       'Muere?
       If UserList(pIndex).Stats.MinHP <= 0 Then
          UserDie pIndex
          Call WriteConsoleMsg(pIndex, .Tag & " Te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
       End If
        
       'Intervalo
       .Intervalos.ArrowCount = (IA_PROINT / 20)
        
       'client update
       WriteUpdateHP pIndex
       Exit Sub
    End If
    
    'Puede castear?
    'Si el usuario no tiene la vida llena ataca
    Dim tmpHP   As Long
    
    tmpHP = (UserList(pIndex).Stats.MinHP)
    
    'obtengo el % de vida del user
    tmpHP = (tmpHP * 100) / (UserList(pIndex).Stats.MaxHP)
   
    If .Intervalos.SpellCount = 0 Then
    
    'Es cl�rigo y puede pegar??
    If (.clase = eIAClase.Clerigo) And .Intervalos.HitCount = 0 And Not .UltimaAccion = eIAactions.ePegar Then
       'Est� al alcance de la v�ctima para un gole meele?
       Dim newBotHeading   As eHeading
       
       If ia_PuedeMeele(.Pos, UserList(pIndex).Pos, newBotHeading) Then
            'Acierta el golpe?
            If ia_AciertaGolpe(pIndex) Then
               'Antes que nada cambiamos el heading, si es v�lido.
               If newBotHeading <> 0 And newBotHeading <> .Char.heading Then
                  'ia_SendToBotArea botIndex, mod_DunkanProtocol.Send_ChangeHeadingChar(.Char.CharIndex, newBotHeading)
               End If
               
               'Calcula el golpe
               Dim GolpeVal     As Integer
               GolpeVal = ia_CalcularGolpe(pIndex)
               
               'Resta hp.
               UserList(pIndex).Stats.MinHP = UserList(pIndex).Stats.MinHP - GolpeVal
               
               'crea el fx de la sangre.
               SendData SendTarget.ToPCArea, pIndex, PrepareMessageCreateFX(UserList(pIndex).Char.CharIndex, FXSANGRE, 5)
               
               'Avisa.
               Call WriteConsoleMsg(pIndex, .Tag & " Te ha pegado por " & CStr(GolpeVal) & ".", FontTypeNames.FONTTYPE_FIGHT)
               
               'Setea flag.
               .UltimaAccion = eIAactions.ePegar
               
               'Muere?
               If UserList(pIndex).Stats.MinHP <= 0 Then
                  Call UserDie(pIndex)
               End If
               
               'update hp.
               WriteUpdateHP pIndex
               
               'Intervalo de golpe.
               .Intervalos.HitCount = (IA_HITINT / 40)
               'Intervalo de hechizo.
               .Intervalos.SpellCount = (IA_SINT / 40)
               'Intervalo de golpe+pociones.
               .Intervalos.UseItemCount = (IA_USEOBJ / 60)
               Exit Sub
            End If
        End If
    End If
    
       'Feo, aunque digamos que solo hace apoca desc remo
       'As� que va a andar bien.
       
       'Si la mana es < a 300 [gasto del remo] no hacemos nada.
       
       If .Mana < 300 Then Exit Sub
       
       'Si est� paralizado AND el usuario no tiene poka vida prioriza removerse.
       
        If .Paralizado And tmpHP > 60 Then
            
            'Intervalo de remo :@
            If .Intervalos.ParalizisCount <> 0 Then Exit Sub
            
            'Palabras m�gicas.
            
            ia_SendToBotArea BotIndex, PrepareMessageChatOverHead(Hechizos(10).PalabrasMagicas, .Char.CharIndex, vbCyan)
           
            .Paralizado = False
           
            'Agrego esto por que si no tirarle inmo era al pedo
            'Seguia caminando practicamente :PP
           
            .Intervalos.ParalizisCount = (IA_SREMO / 10)
           
            'Se removi� entonces salimos del sub y seteamos el intervalo
           
            .Intervalos.SpellCount = (IA_SINT / 40)
           
            Exit Sub
           
        End If
       
        'No est� paralizado entonces castea un hechizo random.
       
        'chances de pegar
        If RandomNumber(1, 100) > IA_CASTEO Then Exit Sub
       
        sRandom = RandomNumber(1, IA_M_SPELL)
       
        'Ayuda otros bots si es que hay
        If NumInvocados <> 1 Then
           ia_SupportOthers BotIndex, AyudoBot
           
           'ayudo ya al bot?
           If AyudoBot Then
              'SETEA INTERVALO
              .Intervalos.SpellCount = (IA_SINT / 40)
              Exit Sub
           End If
           
        End If
           
        'Si el usuario ya estaba paralizado AND el random es paralizar, entonces buscamos de nuevo
        If UserList(pIndex).flags.Paralizado = 1 And sRandom = 3 Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
        
        'Si soy mago y el usuario es mago tambi�n no paraliza.
        If UserList(pIndex).clase = eClass.Mage And .clase = eIAClase.Mago Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
        
        'Si el usuario tiene menos del 75% de vida juega al ataque.
        
        If tmpHP < 75 Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
        
        'Si no llega con la mana del hechizo AND la del otro
        'tampoco entonces no hacemos nada
       
        If sRandom = 1 Then
           
            'Si no llega a la mana del spell 1 (descarga)
            'No hacemos nada ya que tampoco llega
            'al apocalipsis.
           
            rMan = Hechizos(ia_spell(1).spellIndex).ManaRequerido
           
            If rMan > .Mana Then Exit Sub
           
        ElseIf sRandom = 2 Then
       
            rMan = Hechizos(ia_spell(2).spellIndex).ManaRequerido
               
            'Pero si es spell 2 (apoca) AND llegamos
            'con la mana para descarga, entonces
            'Seteamos sRandom como 1 y casteamos
            'descarga.
           
            If rMan > .Mana Then
               
                'Modifico la formula y hago un random
                'Dado a que una ves que queda en -1000 de mana
                'Nunca m�s tira apoca y castea puras descargas.
               
                If .Mana > 460 And RandomNumber(1, 100) < 30 Then
                    sRandom = 1
                Else
                    Exit Sub
                End If
            End If
       End If
       
        rMan = Hechizos(ia_spell(sRandom).spellIndex).ManaRequerido
       
        'Descontamos la man� y seteamos el intervalo.
        .Mana = .Mana - rMan
       
        'Set �ltima action.
        .UltimaAccion = eIAactions.eMagia
        
        .Intervalos.SpellCount = (IA_SINT / 20) 'Se chekea cada 40 ms.
       
        'Creamos el fx y le descontamos la vida al usuario.
       
        ia_SendToBotArea BotIndex, Protocol.PrepareMessageCreateFX(UserList(pIndex).Char.CharIndex, Hechizos(ia_spell(sRandom).spellIndex).FXgrh, Hechizos(ia_spell(sRandom).spellIndex).Loops)
       
        ia_SendToBotArea BotIndex, PrepareMessageChatOverHead(Hechizos(ia_spell(sRandom).spellIndex).PalabrasMagicas, .Char.CharIndex, vbCyan)
       
        'Paralizar?
        If sRandom = 3 Then
           'Paralizado : P
           UserList(pIndex).flags.Paralizado = 1
           UserList(pIndex).Counters.Paralisis = IntervaloParalizado
           WriteParalizeOK pIndex
           WriteConsoleMsg pIndex, .Tag & " Te ha paralizado", FontTypeNames.FONTTYPE_FIGHT
        End If
       
        'Random damage :D
       
        sRandom = RandomNumber(ia_spell(sRandom).DamageMin, ia_spell(sRandom).DamageMax)
       
        'Al da�o le restamos , si el usuario tiene, defensa m�gica.
        If UserList(pIndex).Invent.AnilloEqpObjIndex <> 0 Then
           sRandom = sRandom - RandomNumber(ObjData(UserList(pIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(pIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        End If
        
        'NO numeros negativos.
        If sRandom < 0 Then sRandom = 0
       
        'Quitamos da�o.
        UserList(pIndex).Stats.MinHP = UserList(pIndex).Stats.MinHP - sRandom
            
        If sRandom <> 0 Then
            'AVISO AL USUARIO DE ESTO
            Call WriteConsoleMsg(pIndex, .Tag & " Te ha quitado " & sRandom & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        'Check si muere.
        If UserList(pIndex).Stats.MinHP <= 0 Then
             UserDie pIndex
             
            'Era viajante y mat� el usuario?, resteo el ui
             If Not pIndex <> .ViajanteUser Then
                .ViajanteUser = 0
             End If
             
             'aviso que murio.
             Call WriteConsoleMsg(pIndex, .Tag & " Te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
             
        End If
       
        'Actualizamos el cliente.
       
        WriteUpdateHP pIndex
       
    End If
End With
 
Exit Sub
 
'Errhandler:
 
End Sub
 
Sub ia_EnviarChar(ByVal UserIndex As Integer, ByVal BotIndex As Byte)
 
' @designer     : Walyto132
' @date         :  2012/03/13
' @                Envia el char del bot a un usuario (sistema de areas!!)
 
    With ia_Bot(BotIndex).Char
            Dim tmp_Color   As eNickColor
            
            If ia_Bot(BotIndex).esCriminal Then
               tmp_Color = eNickColor.ieCriminal
            Else
               tmp_Color = eNickColor.ieCiudadano
            End If
            
            Call Protocol.WriteCharacterCreate(UserIndex, .Body, .Head, eHeading.SOUTH, .CharIndex, ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y, .WeaponAnim, .ShieldAnim, 0, 0, .CascoAnim, ia_Bot(BotIndex).Tag, tmp_Color, 0, 0, 0)
    End With
 
End Sub
 
Sub ia_UserDamage(ByVal spell As Byte, ByVal BotIndex As Byte, ByVal UserIndex As Integer, Optional ByVal is_RuneArea As Boolean = False)
 
' @designer     : Walyto132
' @date         : 2021/06/05
 
Dim rMan     As Integer
Dim Damage   As Integer
Dim usedFont As FontTypeNames
 
usedFont = FontTypeNames.FONTTYPE_FIGHT
 
'Checkeo que el hechizo no sea 0.
If Not spell <> 0 Then Exit Sub
 
With UserList(UserIndex)
 
    rMan = Hechizos(spell).ManaRequerido
   
    'Llega con la mana?
    If rMan > .Stats.MinMAN Then
        WriteConsoleMsg UserIndex, "No tienes suficiente mana!", usedFont
        Exit Sub
    End If
    
    'Skills?
    If Hechizos(spell).MinSkill > .Stats.UserSkills(eSkill.Magia) Then
       WriteConsoleMsg UserIndex, "No tienes suficientes skills en magia", usedFont
       Exit Sub
    End If
   
    'Soy ciudadano y el target es un bot viajante?
    If Not Criminal(UserIndex) And ia_Bot(BotIndex).Viajante And .flags.Seguro Then
        WriteConsoleMsg UserIndex, "Para atacar bots viajantes debes desactivar el seguro", usedFont
        Exit Sub
    End If
    
    If Hechizos(spell).Inmoviliza Or Hechizos(spell).Paraliza Then
       
        'Le pongo el flag en verdadero.
        ia_Bot(BotIndex).Paralizado = True
       
        'Mensaje informando.
        WriteConsoleMsg UserIndex, "Has paralizado ah " & ia_Bot(BotIndex).Tag, usedFont
        
        'Creo la animacion sobre el char.
        ia_SendToBotArea BotIndex, Protocol.PrepareMessageCreateFX(ia_Bot(BotIndex).Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).Loops)
        
        'SpellWorlds.
        DecirPalabrasMagicas Hechizos(spell).PalabrasMagicas, UserIndex
       
        'Quito mana y energia
        .Stats.MinMAN = .Stats.MinMAN - rMan
       
        'le doy intervalo
       
        ia_Bot(BotIndex).Intervalos.ParalizisCount = (IA_SREMO / 10)
       
        WriteUpdateMana UserIndex
       
        Exit Sub
    End If
   
    'Era un Viajante
   
    Damage = RandomNumber(Hechizos(spell).MinHP, Hechizos(spell).MaxHP)
    Damage = Damage + Porcentaje(Damage, 3 * .Stats.ELV)
   
   If Not Damage <> 0 Then Exit Sub
   
   If ia_Bot(BotIndex).Viajante Then
        Dim eraPK   As Boolean
        
        If Not ia_Bot(BotIndex).ViajanteAntes.map Then
            ia_Bot(BotIndex).ViajanteAntes = ia_Bot(BotIndex).Pos
        End If
        
        'No era criminal.
        eraPK = Criminal(UserIndex)
    
        'No era criminal y atac� un viajante, es criminal.
        If Not eraPK Then VolverCriminal UserIndex
    
        'Ahora el bot se enoj� viejo..
        ia_Bot(BotIndex).ViajanteUser = UserIndex
    
        'UserList(UserIndex).AtacoViajante = BotIndex
    
        WriteConsoleMsg UserIndex, "Has atacado un viajante!! ahora eres un criminal, y adem�s el viajante te atacar�!", usedFont
   End If
   
    'Quitamos vida
    If Hechizos(spell).StaffAffected Then
       If UserList(UserIndex).clase = eClass.Mage Then
          If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
             Damage = (Damage * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
          Else
             Damage = Damage * 0.7 'Baja damage a 70% del original
          End If
        End If
     End If
        
     If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        Damage = Damage * 1.04  'laud magico de los bardos
     End If
    
    ia_Bot(BotIndex).Vida = ia_Bot(BotIndex).Vida - Damage
    
    'No est� paralizado.
    If Not ia_Bot(BotIndex).Paralizado Then
        'Le pegaron, se cag� todo y se mueve random.
        Dim keepMoving  As Boolean
    
        ia_RandomMoveChar BotIndex, UserIndex, keepMoving
    
        'No hubo error, por ende se movi�.
        If Not keepMoving Then
           'Guardo la nueva pos.
           MapData(ia_Bot(BotIndex).Pos.map).Tile(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y).BotIndex = BotIndex
       
           'Actualizo el area del bot.
           ia_SendToBotArea BotIndex, PrepareMessageCharacterMove(ia_Bot(BotIndex).Char.CharIndex, ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y)
       
           'Intervalo de caminata.
           ia_Bot(BotIndex).Intervalos.MoveCharCount = (IA_MOVINT / 40)
        End If
        
    End If
    
    'Aviso al usuario.
    WriteConsoleMsg UserIndex, "Le has quitado " & CStr(Damage) & " puntos de vida a " & ia_Bot(BotIndex).Tag, usedFont
   
    'Tiro las spell worlds
    DecirPalabrasMagicas Hechizos(spell).PalabrasMagicas, UserIndex
   
    'Creo el fx.
    ia_SendToBotArea BotIndex, Protocol.PrepareMessageCreateFX(ia_Bot(BotIndex).Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).Loops)
   
    'saco mana y energia y actualizo el cliente
    .Stats.MinMAN = .Stats.MinMAN - rMan
       
    WriteUpdateMana UserIndex
   
    If ia_Bot(BotIndex).Vida <= 0 Then
        'Muri�?
        ia_EraseChar BotIndex, True
        WriteConsoleMsg UserIndex, "Has matado ah " & ia_Bot(BotIndex).Tag & ".", usedFont
    End If
   
End With
 
End Sub

Sub ia_DamageHit(ByVal BotIndex As Byte, ByVal UserIndex As Integer)

' @designer     : Walyto132
' @date         : 2021/06/05

    Dim nDamage As Integer

    'Calculo el da�o.
    nDamage = CalcularDa�o(UserIndex)

    'Resto la defensa del bot.
    nDamage = nDamage - (RandomNumber(IA_MINDEF, IA_MAXDEF))

    'Aviso al usuario.
    WriteConsoleMsg UserIndex, "Le has pegado ah " & ia_Bot(BotIndex).Tag & " por " & nDamage, FontTypeNames.FONTTYPE_FIGHT

    'Creo da�o :)
    'ia_SendToBotArea BotIndex, mod_DunkanProtocol.Send_CreateDamage(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y, nDamage)

    'Resto vida.
    ia_Bot(BotIndex).Vida = ia_Bot(BotIndex).Vida - nDamage

    'seteo el flag.
    'UserList(UserIndex).AtacoViajante = BotIndex

    'Murio?
    If ia_Bot(BotIndex).Vida <= 0 Then

        'Era viajante?
        If ia_Bot(BotIndex).Viajante Then
            'Reset el flag.
            'UserList(UserIndex).AtacoViajante = 0
        End If

        ia_EraseChar BotIndex, True

    End If

End Sub

Sub ia_SendToBotArea(ByVal BotIndex As Byte, ByVal PackData As String)

' @designer     : Walyto132
' @date         :  2012/03/13
' @                Envia paquetes al area de un bot.

'Nueva versi�n del sub, m�s simple y dir�a que m�s pr�ctica : P

    With ia_Bot(BotIndex)
        'd3 ao, borro esto : p

        'con esto tenemos algo simple, cuando mandamos el send
        'tobotarea, nos devuelve un array con los ui y el ping de cada
        'uno, y flush_ping tiene el promedio :), despues solo nos
        'queda comprobar si el usuario puede flushbuffear los datos
        'y enviamos, sacrificamos memoria pero ganamos MUCHA conexi�n.

        'Dim flush_Ping      As Integer
        'Dim arr_PingUsers() As Integer

        'Call modSendData.SendToAreaByPos(.Pos.map, .Pos.X, .Pos.Y, PackData, .GrupoID, flush_Ping)

        'Do While flush_Ping <> 0
        '    If can_Update_Ping(arr_PingUsers(flush_Ping)) Then
        '       Call flusH_buffer_to_base_Ping(arr_PingUsers(flush_Ping), flush_Ping, .GrupoID)
        '    End If

        '    flush_Ping = flush_Ping - 1

        'Loop

        Call modSendData.SendToAreaByPos(.Pos.map, .Pos.X, .Pos.Y, PackData)
    End With

End Sub
 
Sub ia_TirarInventario(ByVal BotIndex As Byte)
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Pincha el inventario de un bot.
 
Dim loopX   As Long
Dim iObjs() As Integer
Dim iObj    As Obj
Dim tmpPos  As WorldPos
 
'Arma array de objetos
ia_ArrayObjetos iObjs, BotIndex
 
For loopX = 1 To UBound(iObjs())
 
    'Crea el objeto.
    iObj.ObjIndex = iObjs(loopX)
 
    'Si el objIndex es >= 36 and <=30  , son pociones
    If iObjs(loopX) >= 36 And iObjs(loopX) <= 39 Then
       iObj.Amount = RandomNumber(1000, 1200)
    Else
       'No eran pociones, son flechas?
       If Not iObjs(loopX) <> 553 Then
          iObj.Amount = RandomNumber(500, 900)
       Else
          iObj.Amount = 1
       End If
    End If
    
    'Si eran pociones azules y el bot era caza..
    If iObj.Amount = 37 And ia_Bot(BotIndex).clase = eIAClase.Cazador Then iObj.Amount = 0
    
    'si hay objIndex.
    If iObj.ObjIndex Then
        'Busca un tile libre.
        Call Tilelibre(ia_Bot(BotIndex).Pos, tmpPos, iObj, True, True)
    
        'Si encontr� (raro que no encuentre)
        If tmpPos.X <> 0 And tmpPos.Y <> 0 Then
           'Crea el objeto
           MakeObj iObj, tmpPos.map, tmpPos.X, tmpPos.Y
        End If
    End If
    
Next loopX
 
'Ya tiro los objetos de su equipo, ahora , si era viajante, tira los que lukeo, si es que tiene.
If ia_Bot(BotIndex).Viajante Then
   For loopX = 1 To IA_SLOTS
       With ia_Bot(BotIndex).Inv(loopX)
            
            iObj.ObjIndex = .ObjIndex
            iObj.Amount = .Amount
            
            Call Tilelibre(ia_Bot(BotIndex).Pos, tmpPos, iObj, True, True)
            
            'Si encontr� posici�n.
            If tmpPos.X <> 0 And tmpPos.Y <> 0 Then
               MakeObj iObj, tmpPos.map, tmpPos.X, tmpPos.Y
            End If
       End With
   Next loopX
End If
 
End Sub
 
Sub ia_ArrayObjetos(ByRef arrayObjs() As Integer, ByVal BotIndex As Byte)
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Arma un array de objetos.
 
'Set primeras dimensiones. (potas,arma y casco)
 
ReDim arrayObjs(1 To 4) As Integer
 
'Pociones.
arrayObjs(1) = 38
arrayObjs(2) = 37
 
'Arma
arrayObjs(3) = ia_ArmaByClase(BotIndex)
 
'Casco
arrayObjs(4) = ia_CascoByClase(BotIndex)
 
'Si no es mago, tiene escudo y dopas.
If ia_Bot(BotIndex).clase <> eIAClase.Mago Then
   'redim
   ReDim Preserve arrayObjs(1 To 7) As Integer
   arrayObjs(5) = ia_EscudoByClase(BotIndex)
   arrayObjs(6) = 36
   arrayObjs(7) = 39
End If
 
'Si es caza, tira flechas.
'No sabemos el ultimo elemento que tenemos!! no jugarsela a tirar 5.
 
If ia_Bot(BotIndex).clase = eIAClase.Cazador Then
   ReDim Preserve arrayObjs(1 To UBound(arrayObjs()) + 1) As Integer
   arrayObjs(UBound(arrayObjs())) = 553
End If
 
End Sub
 
Sub ia_EraseChar(ByVal BotIndex As Byte, Optional ByVal killedbyUSER As Boolean = False)
 
' @designer     : Walyto132
' @date         : 2021/06/05
' @note         :  Borra el char y los datos del bot.
 
With ia_Bot(BotIndex)
    'Borro el char.
    ia_SendToBotArea BotIndex, PrepareMessageCharacterRemove(.Char.CharIndex)
    
    'Borro el botIndex
    MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).BotIndex = 0
    
    Dim dummyPos    As WorldPos
    
    .ViajanteAntes = dummyPos
    
    'Mat� un usuario? pincha inventario
    If killedbyUSER Then
       ia_TirarInventario BotIndex
    End If
    
    'Reset char,
    With .Char
         .Body = 0
         .CascoAnim = 0
         .FX = 0
         .Loops = 0
         .Head = 0
         .heading = 0
         .ShieldAnim = 0
         .WeaponAnim = 0
    End With
    
    'Reset STATS
    .Vida = 0
    .Mana = 0
    
    'Reset pos.
    With .Pos
         .map = 0
         .X = 0
         .Y = 0
    End With
    
    'Reset flags.
    .Invocado = False
    .Paralizado = False
   
    'Reset intervalos.
    With .Intervalos
         .MoveCharCount = 0
         .SpellCount = 0
         .UseItemCount = 0
         .ParalizisCount = 0
    End With
    
    'Reset viajante flag.
    .Viajante = False
    .ViajanteUser = 0
    
    'Resta el contador
    NumInvocados = NumInvocados - 1
    
End With
 
End Sub
 
Sub ia_CheckInts(ByVal BotIndex As Byte)
 
' @designer     : Walyto132
' @date         : 2021/06/05
 
With ia_Bot(BotIndex).Intervalos
     
    If .ArrowCount > 0 Then .ArrowCount = .ArrowCount - 1
    If .MoveCharCount > 0 Then .MoveCharCount = .MoveCharCount - 1
    If .SpellCount > 0 Then .SpellCount = .SpellCount - 1
    If .UseItemCount > 0 Then .UseItemCount = .UseItemCount - 1
    If .ParalizisCount > 0 Then .ParalizisCount = .ParalizisCount - 1
    If .HitCount > 0 Then .HitCount = .HitCount - 1
    If .ChatCount > 0 Then .ChatCount = .ChatCount - 1
    
End With
 
End Sub
 
Function ia_FindTarget(Pos As WorldPos, Optional ByVal esPk As Boolean = False) As Integer
 
' @designer     : Walyto132
' @date         :  2012/03/13
' @note         :  Busca alguien a quien pegar
 
Dim loopX       As Long         '< Bucle del tileX.
Dim loopY       As Long         '< Bucle del tileY.
Dim tmpIndex    As Integer
 
For loopY = Pos.Y - MargenY To Pos.Y + MargenY
        For loopX = Pos.X - MargenX To Pos.X + MargenX
            'Hay usuario?
            If MapData(Pos.map).Tile(loopX, loopY).UserIndex > 0 Then
               'No est� muerto
               If UserList(MapData(Pos.map).Tile(loopX, loopY).UserIndex).flags.Muerto = 0 Then
                  
                  If UserList(MapData(Pos.map).Tile(loopX, loopY).UserIndex).flags.AdminInvisible = 1 Then Exit Function
                  If EsAdmin(UserList(MapData(Pos.map).Tile(loopX, loopY).UserIndex).Name) = True Then Exit Function
                  'Es ciuda el bot y el usuario?
                  If Not esPk Then
                     'el bot no es pk.
                     ia_FindTarget = MapData(Pos.map).Tile(loopX, loopY).UserIndex
                  Else
                     tmpIndex = MapData(Pos.map).Tile(loopX, loopY).UserIndex
                     If Not esPk And Criminal(tmpIndex) Then
                         ia_FindTarget = tmpIndex
                     Else
                        If esPk And Not Criminal(tmpIndex) Then
                           ia_FindTarget = tmpIndex
                        End If
                    End If
                 End If
                  Exit Function
               End If
            End If
        Next loopX
Next loopY
 
ia_FindTarget = 0
End Function
 
Function IA_GetNextSlot() As Byte
 
' @ Devuelve un slot para bots.
 
Dim loopX   As Long
 
For loopX = 1 To MAX_BOTS
    If Not ia_Bot(loopX).Invocado Then
       IA_GetNextSlot = CByte(loopX)
       Exit Function
    End If
Next loopX
 
IA_GetNextSlot = 0
 
End Function


