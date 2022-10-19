Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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
Public CodVerificacion As String
Public CorreoVal As String
Public Resolucion As Boolean
Public SeguroResu As Long
Public SeguroConIma As Long
Public Consolacom As Long
Public TT2 As New CBalloonToolTip
Public contarr As Integer
Public Type tRanking
    value(0 To 9) As Long
    nombre(0 To 9) As String
End Type


Public Ranking As tRanking

Public Enum eRanking
    TopFrags = 1
    TopTorneos = 2
    TopLevel = 3
    TopOro = 4
    TopRetos = 5
    TopClanes = 6
    TopMuertesP = 7
End Enum
Public RankingOro As String
Public RecuadroX As Single
Public RecuadroY As Single
Public RecuadroON As Boolean
Public RecuadroSON As Boolean
Public RecuadroInv As Boolean
Public Vidarender As Boolean
Public Manarender As Boolean
Public alaPath As String
Public ParticlesORE As New clsOreParticles
Public Colorsangre As Byte
Public invSpells As New clsGrapchicalInventory

Public NieveOn As Boolean

Public SimboloY As Integer

Public SimboloOn As Boolean

Public COLOR_AZUL As Long

Public Movement_Speed As Single

'aura
Public OpcionesPath As String

Public ActivarAuras As String

Public RotarActivado As String

Public Rotacion As Single

'aura
Public Enum eMessages

    mUserMuerto
    mUserParalizado
    mUserComerciando
    mUserEquitando
    mUserEmbarcado
    mUserCongelado
    mUserSaliendo
    UserEstaLejos
    mUserChiquito
    DomarAnimalParaMontarlo
    MonturaNoTeAceptaComoSuAmo
    UserHaMontado
    NoHayLugarParaDesmontar
    UserHaDesmontado
    UserHaDomado
    NoHasLogradoDomarlo
    UserNoPuedeUsarObjetoAqui
    UserNoTieneConocimientosNecesarios
    UserTieneEmbarcacionAnclada
    UserSeVuelveVisible
    UserRecuperaSuAparienciaNormal
    TeHasEscondidoEntreLasSombras
    NoHasLogradoEsconderte
    OponenteNoTienEquipadoItems
    NoHasLogradoDesarmarATuOponente
    NoTienesSuficienteEnergia
    YaDomasteALaCriatura
    NoMascotasZonaSegura
    LaCriaturaTieneAmo
    NoPodesDomarDosCriaturasIguales
    NoPodesControlarMasCriaturas
    HasPescado
    HasPescadoAlgunosPeces
    NoHasPescado
    LaRedSeUsaEnMarAbierto
    DebesQuitarSeguroParaRobarCiudadano
    RealNoRobaCiudadanos
    CaosNoRobaCaos
    EstasMuyCansadoParaRobar
    EstasMuyCansadaParaRobar
    LeHasRobadoMonedasDeOro
    NoTieneOro
    NoHasLogradoRobarNada
    HanIntentadoRobarte
    UserNoTieneObjetos
    HasRobadoObjetos
    HasHurtadoObjetos
    HasConseguidoLena
    NoHasConseguidoLena
    HasExtraidoMinerales
    NoHasExtraidoMinerales
    TerminasDeMeditar
    HasRecuperadoMana
    UserMonturaSiendoAtacada
    CriaturaAgonizando
    EstasMuyCansadoParaLuchar
    EstasMuyCansadaParaLuchar
    NoPodesAtacarEsteNPC
    LejosParaDisparar
    VictimaMuerto
    NoPodesPelearAqui
    NoPodesAtacarteAVosMismo
    LejosParaAtacar

    LanzaHechizoA
    TeLanzanHechizo

    HasDesequipadoEsudoOponente
    TeHanDesequipadoEscudo
    HasDesarmadoAlOponente
    TeHanDesarmado
    HasDesequipadoCascoOponente
    TeHanDesequipadoCasco
    HasApunaladoA
    TeHanApunalado
    HasApunaladoCriatura
    NoHasApunalado
    HasGolpeadoCriticamente
    TeHanGolpeadoCriticamente
    HasGolpeadoCriticamenteCriatura

    NpcInmune
    Hechizo_PropioMSG
    Hechizo_TargetMSG

End Enum

Public Enum eEffects

    Inmovilizar = 1
    Flecha1 = 2
    Flecha2 = 3
    Flecha3 = 4
    Flecha4 = 5
    Bala = 6

End Enum

Public sintextos As Boolean            'helios consola

Public MostrarMenuInventario As Boolean            'helios 2/06/2021

Public MmenuBarras As Boolean            'helios 2/06/2021

Public ScreenShooterCapturePending As Boolean

Public FragShooterEsperandoLevel As Boolean

Public AngMareoMuerto As Single

Public RadioMareoMuerto As Single

Public TiempoHome As Integer

Public GoingHome As Byte

'Objetos públicos
Public DialogosClanes As New clsGuildDlg

Public Dialogos As New clsDialogs

Public Audio As New clsAudio

Public vPasos As New clsPasos

Public Inventario As New clsGrapchicalInventory

Public InvBanco(1) As New clsGrapchicalInventory

'Inventarios de comercio con usuario
Public InvComUsu As New clsGrapchicalInventory                               ' Inventario del usuario visible en el comercio

Public InvOroComUsu(2) As New clsGrapchicalInventory    ' Inventarios de oro (ambos usuarios)

Public InvOfferComUsu(1) As New clsGrapchicalInventory    ' Inventarios de ofertas (ambos usuarios)

Public InvComNpc As New clsGrapchicalInventory

'Inventarios de herreria
Public Const MAX_LIST_ITEMS As Byte = 4

Public InvLingosHerreria(1 To MAX_LIST_ITEMS) As New clsGrapchicalInventory

Public InvMaderasCarpinteria(1 To MAX_LIST_ITEMS) As New clsGrapchicalInventory

Public SurfaceDB As clsSurfaceManager                        'No va new porque es una interfaz, el new se pone al decidir que clase de objeto es

Public CustomKeys As New clsCustomKeys

Public CustomMessages As New clsCustomMessages

Public incomingData As New clsByteQueue

Public outgoingData As New clsByteQueue

Public DummyCode() As Byte

Public iServer As Integer

Public iCliente As Integer

Public ContarClip As Long         'Helios 28/06/2021

Public PulsarEsconder As Long         'Helios 07/06/2021

Public CTextos As Long         'helios 2/06/2021

''
'The main timer of the game.
Public MainTimer As New clsTimer

#If SeguridadAlkon Then

    Public MD5 As New clsMD5
#End If

'Sonidos
Public Const SND_CLICKNEW As Integer = 210

Public Const SND_CLICKOFF As Integer = 212

Public Const SND_MOUSEOVER As Integer = 211

Public Const SND_CLICK As Integer = 221

Public Const SND_PASOS1 As Integer = 23

Public Const SND_PASOS2 As Integer = 24

Public Const SND_NAVEGANDO As Integer = 50

Public Const SND_OVER As Integer = 222

Public Const SND_CADENAS As Integer = 220

Public Const SND_LLUVIAINEND As Integer = 224

Public Const SND_LLUVIAOUTEND As Integer = 225

Public Const SND_LLUVIAIN As Integer = 226

Public Const SND_LLUVIAOUT As Integer = 227

Public Const SND_LLUVIAINSTART As Integer = 228

Public Const SND_LLUVIAOUTSTART As Integer = 229

Public Const SND_FUEGO As Integer = 230

Public Const SND_SNAPSHOT As Integer = 231

Public Const SND_APUÑALAR As Integer = 173

' Head index of the casper. Used to know if a char is killed

Public WAIT_ACTION As Integer

Public Enum eWAIT_FOR_ACTION

    None = 0
    RPU = 1

End Enum

' Constantes de intervalo
'INTERVALOS DE COMBATE
Public Const INT_PUEDE_RPU_MOVER As Integer = 25    'Puede moverse despues de pedir RPU o Colisionar

Public Const INT_PUEDE_MOVER As Integer = 38                 'Puede caminar

Public Const INT_PUEDE_MOVER_EQUITANDO As Integer = 36    'Puede equitar

Public Const INT_RPU_MOVE As Integer = 125                  'Puede moverse luego de RPU

Public Const INT_PUEDE_GOLPE As Integer = 500                  'Puede golpear con arma o puño

Public Const INT_HABILITA_ICONO_LANZAR_HECHIZO As Integer = 450    ' Te habilita el icono para lanzar hechizo

Public Const INT_LANZAR_HECHIZO As Integer = 1000    'Puede lanzar un hechizo

Public Const INT_GOLPE_MAGIA As Integer = 800                  'Puede lanzar un hechizo luego de golpear con arma o puño

Public Const INT_MAGIA_GOLPE As Integer = 600                  'Puede golpear luego de haber tirado un hechizo

Public Const INT_FLECHAS As Integer = 1180                   'Puede lanzar proyectil

Public Const INT_PUEDE_USAR As Integer = 450                  'Puede usar item

Public Const INT_PUEDE_USAR_DOBLECLICK As Integer = 125    'Puede usar item haciendo doble click constante

Public Const INT_GOLPE_USAR As Integer = 700                  'Puede usar un item luego de haber golpeado con arma o puño

'INTERVALOS
Public Const INT_WORK As Integer = 700                  'Puede trabajar

Public Const INT_SENTRPU As Integer = 800                  'Puede reniciar su posición "L"

Public Const INT_HIDE As Integer = 800                  'Puede ocultarse

Public Const INT_BUY As Integer = 150                  'Puede comprar un item en el comercio de NPCS

Public Const INTERVALO_INVI As Byte = 120               'Intervalo de invisibilidad

Public Const INT_MONTAR As Integer = 600                  'Puede montar/desmontar una criatura

Public Const INT_ANCLAR As Integer = 600                  'Puede anclar/desanclar embarcación (nadar)

Public Const INT_MACRO_HECHIS As Integer = 2788    'INT macro de hechizos

Public Const INT_MACRO_TRABAJO As Integer = 900    'INT macro de trabajo

Public Const INT_TELEP As Integer = 350                  'Puede teletransportarse con shift + click

Public MacroBltIndex As Integer

Public Const CASPER_HEAD As Integer = 500

Public Const CASPER_HEAD_CRIMI As Integer = 501

Public Const FRAGATA_FANTASMAL As Integer = 87

Public Const NUMATRIBUTES As Byte = 5

'Musica
Public Const MIdi_Inicio As Byte = 6

Public RawServersList As String

Public Type tColor

    r As Byte
    G As Byte
    b As Byte

End Type

Public ColoresPJ(0 To 50) As tColor

Public Type tServerInfo

    Ip As String
    puerto As Integer
    Desc As String
    PassRecPort As Integer

End Type

Public ServersLst() As tServerInfo

Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean

Public ClanName As String

Public Site As String

'USER FLAGS
Public UserCiego As Boolean

Public UserEstupido As Boolean

Public UserEquitando As Boolean

Public UserCongelado As Boolean

Public UserChiquito As Boolean

Public UserNadando As Boolean

Public NoRes As Boolean            'no cambiar la resolucion

Public GraphicsFile As String    'Que graficos.ind usamos

Public RainBufferIndex As Long

Public FogataBufferIndex As Long

Public TerrenoBufferIndex As Long

Public Const bCabeza = 1

Public Const bPiernaIzquierda = 2

Public Const bPiernaDerecha = 3

Public Const bBrazoDerecho = 4

Public Const bBrazoIzquierdo = 5

Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000

Public Const tUs = 600

Public Const PrimerBodyBarco = 84

Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero() As tItemsConstruibles

Public ArmadurasHerrero() As tItemsConstruibles

Public ObjCarpintero() As tItemsConstruibles

Public CarpinteroMejorar() As tItemsConstruibles

Public HerreroMejorar() As tItemsConstruibles

Public UsaMacro As Boolean

Public CnTd As Byte

'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40

Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]

Public Tips() As String * 255

Public Const LoopAdEternum As Integer = 999

'Direcciones
Enum E_Heading

    north = 1
    east = 2
    south = 3
    west = 4

End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000

Public Const MAX_INVENTORY_SLOTS As Byte = 20

Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 20

Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50

Public Const MAXHECHI As Byte = 35

Public Const MAXSKILLPOINTS As Byte = 100

Public Const INV_OFFER_SLOTS As Byte = 20

Public Const INV_GOLD_SLOTS As Byte = 1

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1

Public Const GOLD_OFFER_SLOT As Integer = INV_OFFER_SLOTS + 1

Public Const Fogata As Integer = 1521

Public Enum eClass

    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Paladín
    Hunter      'Cazador
    Worker
    Pirat       'Pirata

End Enum

Public Enum eCiudad

    cUllathorpe = 1
    cNix
    cBanderbill
    cArkhein
    cArghal
    cLindos

End Enum

Enum eRaza

    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano

End Enum

Public Enum eSkill

    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Talar = 9
    Comerciar = 10
    Defensa = 11
    Pesca = 12
    Mineria = 13
    Carpinteria = 14
    Herreria = 15
    Liderazgo = 16
    Domar = 17
    Proyectiles = 18
    Wrestling = 19
    Navegacion = 20

End Enum

Public Enum eAtributos

    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5

End Enum

Enum eGenero

    Hombre = 1
    Mujer

End Enum

Public Enum PlayerType

    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80

End Enum

Public Enum eObjType

    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otescudo = 16
    otcasco = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otAlas = 40
    otCualquiera = 1000

End Enum

Public HayTrueno As Boolean

Public Enum ePicRenderType

    BloodDie
    Blood
    Ceguera
    TextKills

End Enum

'Alpha Efectos
Public CUENTA As Integer

Public AlphaCuenta As Integer

Public AlphaBlood As Integer

Public AlphaBloodUserDie As Integer

Public AlphaCeguera As Single

Public AlphaSalir As Integer

Public OrigHora As Byte

Public AlphaRelampago As Integer

Public HayRelampago As Boolean

Public TextKillsType As Byte

Public AlphaTextKills As Integer

'''''''''''''''''''' CURSOR '''''''''''''''''''''''
Public Enum eCursor

    General
    Hechiz
    proyectil
    ProyectilPequena

End Enum

Public curGeneral As New clsAniCursor

Public curGeneralCrimi As New clsAniCursor

Public curGeneralCiuda As New clsAniCursor

Public curProyectil As New clsAniCursor

Public curProyectilPequena As New clsAniCursor

Public Const FundirMetal As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Enum eNPCType

    Comun = 0
    Revividor = 1
    Guardia = 2
    Entrenador = 3
    Banquero = 4
    Noble = 5
    DRAGON = 6
    Timbero = 7
    ResucitadorNewbie = 9
    Mercader = 10
    Fortaleza = 11
    Marinero = 12
    Gobernador = 13
    Montura = 14

End Enum

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = " fallo el golpe!!!"

Public Const MENSAJE_CRIATURA_MATADO As String = " te ha matado!!!"

Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "¡Has rechazado el ataque con el escudo!"

Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO As String = "¡El usuario rechazo el ataque con su escudo!"

Public Const MENSAJE_FALLADO_GOLPE As String = "¡Has fallado el golpe!"

Public Const MENSAJE_SEGURO_ACTIVADO As String = ">>SEGURO DE COMBATE ACTIVADO<<"

Public Const MENSAJE_SEGURO_DESACTIVADO As String = ">>SEGURO DE COMBATE DESACTIVADO<<"

Public Const MENSAJE_PIERDE_NOBLEZA As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."

Public Const MENSAJE_USAR_MEDITANDO As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_SEGURO_RESU_ON As String = ">>SEGURO DE RESURRECCION ACTIVADO<<"

Public Const MENSAJE_SEGURO_RESU_OFF As String = ">>SEGURO DE RESURRECCION DESACTIVADO<<"

Public Const MENSAJE_GOLPE_CABEZA As String = " te ha pegado en la cabeza por "

Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "

Public Const MENSAJE_GOLPE_BRAZO_DER As String = " te ha pegado el brazo derecho por "

Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "

Public Const MENSAJE_GOLPE_PIERNA_DER As String = " te ha pegado la pierna derecha por "

Public Const MENSAJE_GOLPE_TORSO As String = " te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"

Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "Le has pegado "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "

Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado en el brazo izquierdo por "

Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado en el brazo derecho por "

Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado en la pierna izquierda por "

Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado en la pierna derecha por "

Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_HECHIZO As String = "Le has causado "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "Le has pegado "

Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "

Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "

Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "

Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "

Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "

Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."

Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."

Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la víctima..."

Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."

Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."

Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."

Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "

Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

Public Const MENSAJE_FRAGSHOOTER_TE_HA_MATADO As String = "te ha matado!"

Public Const MENSAJE_FRAGSHOOTER_HAS_MATADO As String = "Has matado a"

Public Const MENSAJE_FRAGSHOOTER_HAS_GANADO As String = "Has ganado "

Public Const MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA As String = "puntos de experiencia."

'Inventario
Type Inventory

    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Single
    ObjType As Integer
    MinDef As Integer
    MaxDef As Integer
    MaxHit As Integer
    MinHit As Integer
    PuedeUsarItem As Byte

End Type

Type NpCinV

    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Single
    ObjType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
    PuedeUsarItem As Byte

End Type

Type tReputacion    'Fama del usuario

    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long

    Promedio As Long

End Type

Type tEstadisticasUsu

    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long

End Type

Type tItemsConstruibles

    Name As String
    OBJIndex As Integer
    GrhIndex As Integer
    LinH As Integer
    LinP As Integer
    LinO As Integer
    Madera As Integer
    MaderaElfica As Integer
    Upgrade As Integer
    UpgradeName As String
    UpgradeGrhIndex As Integer

End Type

Private Type tZona

    nombre As String
    Mapa As Byte
    x1 As Integer
    y1 As Integer
    x2 As Integer
    y2 As Integer
    Segura As Byte
    Acoplar As Byte
    Terreno As Byte
    Niebla As Byte
    NieblaR As Byte
    NieblaG As Byte
    NieblaB As Byte
    Musica(1 To 5) As Byte
    CantMusica As Byte
    CantSonidos As Byte

    TieneNpcInvocacion As Byte
    NpcInvocadoIndex As Integer
    Sonido(1 To 5) As Byte

End Type

Public Zonas() As tZona

Public NumZonas As Integer

Public ZonaActual As Integer

Public CambioZona As Single

Public CambioSegura As Boolean

Public zTick As Long

Public zTick2 As Long

Public PergaminoTick As Long

Public PergaminoDireccion As Byte

Public TiempoAbierto As Date

Public LastZona As String

Public Nombres As Boolean

Public UserSeguroResu As Boolean

Public UserSeguro As Boolean

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV

Public UserMeditar As Boolean

Public UserAccount As String

Public UserName As String

Public UserPassword As String

Public UserMaxHP As Integer

Public UserMinHP As Integer

Public UserMaxMAN As Integer

Public UserMinMAN As Integer

Public UserMaxSTA As Integer

Public UserMinSTA As Integer

Public UserMaxAGU As Byte

Public UserMinAGU As Byte

Public UserMaxHAM As Byte

Public UserMinHAM As Byte

Public UserGLD As Long

Public UserLvl As Integer

Public UserPort As Integer

Public UserServerIP As String

Public UserEstado As Byte         '0 = Vivo & 1 = Muerto

Public UserPasarNivel As Long

Public UserExp As Long

Public UserReputacion As tReputacion

Public UserEstadisticas As tEstadisticasUsu

Public UserDescansar As Boolean

Public UserEmbarcado As Boolean

Public BarcoOffSetX As Single

Public BarcoOffSetY As Single

Public tipf As String

Public PrimeraVez As Boolean

Public FPSFLAG As Boolean

Public pausa As Boolean

Public UserParalizado As Boolean

Public UserNavegando As Boolean

Public UserHogar As eCiudad

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As eClass

Public UserSexo As eGenero

Public UserRaza As eRaza

Public UserEmail As String

Public UserFaccion As Byte

Public Const NUMCIUDADES As Byte = 6

Public Const NUMSKILLS As Byte = 20

Public Const NUMATRIBUTOS As Byte = 5

Public Const NUMCLASES As Byte = 12

Public Const NUMRAZAS As Byte = 5

Public UserSkills(1 To NUMSKILLS) As Byte

Public UserSkillsMod(1 To NUMSKILLS) As Byte

Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte

Public AtributosNames(1 To NUMATRIBUTOS) As String

Public FuerzaBk As Byte

Public AgilidadBk As Byte

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String

Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer

Public SPLibres As Integer

Public Oscuridad As Integer

Public logged As Boolean

Public UsingSkill As Integer

Public UsingSecondSkill As Integer

Public MD5HushYo As String * 16

Public pingTime As Long

Public Enum E_MODO

    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    Cuentas = 4
    CrearCuenta = 5
    BorrarPersonaje = 6

End Enum

Public EstadoLogin As E_MODO

Public Enum FxMeditar

    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    XXGRANDE = 34

End Enum

Public Enum eClanType

    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal

End Enum

Public Enum eEditOptions

    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold

End Enum

Public Enum eTerreno

    Default = 0
    Bosque = 1
    Ciudad = 2
    Desierto = 3
    Nieve = 4
    Dungeon = 5
    Mar = 6

End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger

    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
    AUTORESU = 7
    MIRARUP = 8
    MIRARLEFT = 9
    MIRARRIGHT = 10

End Enum

Public Const ORO_INDEX As Integer = 12

Public Const ORO_GRH As Integer = 511

'Server stuff
Public RequestPosTimer As Integer    'Used in main loop

Public stxtbuffer As String    'Holds temp raw data from server

Public stxtbuffercmsg As String    'Holds temp raw data from server

Public SendNewChar As Boolean    'Used during login

Public Connected As Boolean    'True when connected to server

Public DownloadingMap As Boolean    'Currently downloading a map from server

Public UserMap As Integer

'Control
Public prgRun As Boolean    'When true the program ends

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long
''RichTextBox Transparente''
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TRANSPARENT = &H20&
''[END]''

'para escribir y leer variables
Public Declare Function writeprivateprofilestring _
                         Lib "kernel32" _
                             Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                                 ByVal lpKeyname As Any, _
                                                                 ByVal lpString As String, _
                                                                 ByVal lpFileName As String) As Long

Public Declare Function getprivateprofilestring _
                         Lib "kernel32" _
                             Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                               ByVal lpKeyname As Any, _
                                                               ByVal lpdefault As String, _
                                                               ByVal lpreturnedstring As String, _
                                                               ByVal nsize As Long, _
                                                               ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1

Public Declare Function ShellExecute _
                         Lib "shell32.dll" _
                             Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                    ByVal lpOperation As String, _
                                                    ByVal lpFile As String, _
                                                    ByVal lpParameters As String, _
                                                    ByVal lpDirectory As String, _
                                                    ByVal nShowCmd As Long) As Long

'Lista de cabezas
Public Type tIndiceCabeza

    Head(1 To 4) As Integer

End Type

Public Type tIndiceArma

    Arma(1 To 4) As Integer

End Type

Public Type tIndiceCuerpo

    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

Public Type tIndiceFx

    Animacion As Integer
    OffSetX As Integer
    OffSetY As Integer

End Type

Public mOpciones As tOpciones

Public Type tOpciones

    'AUDIO
    Music As Boolean
    sound As Boolean
    SoundEffects As Boolean
    VolMusic As Byte
    VolSound As Byte

    'SCREEN SHOOTER
    ScreenShooterNivelSuperior As Boolean
    ScreenShooterNivelSuperiorIndex As Integer
    ScreenShooterAlMorir As Boolean

    'GUILDS
    DialogConsole As Boolean
    DialogCantMessages As Byte
    GuildNews As Boolean

    'ACCOUNT
    Recordar As Boolean
    RecordarUsuario As String
    RecordarPassword As String

    'VIDEO
    TransparencyTree As Boolean
    Shadows As Boolean
    BlurEffects As Boolean
    Niebla As Boolean
    MostrarAyuda As Boolean
    'Otros
    CursorFaccionario As Boolean

End Type

Public Const GRH_HALF_STAR As Integer = 5357

Public Const GRH_FULL_STAR As Integer = 5358

Public Const GRH_GLOW_STAR As Integer = 5359

Public Const LH_GRH As Integer = 724

Public Const LP_GRH As Integer = 725

Public Const LO_GRH As Integer = 723

Public Const MADERA_GRH As Integer = 550

Public Const MADERA_ELFICA_GRH As Integer = 1999

Public GuildNames() As String

Public GuildMembers() As String

Public picMouseIcon As Picture

Public TradingUserName As String

Public EsPartyLeader As Boolean

Public MirandoParty As Boolean

Public TiempoRetos As Long

Public DragAndDrop As Boolean

Public Const ColorTransparente As Long = vbMagenta

Public Const ColorTransparenteDX As Long = vbMagenta

'Public VolumenCambio As Integer
Public MidiCambio As Byte

Public SinMidi As Boolean

Public iTickMidi As Long

'Particulas
'*************
Public LastTexture As Long

'Public Const ScreenWidth As Long = 800
'Public Const ScreenHeight As Long = 600
Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Public ParticleOffsetX As Long

Public ParticleOffsetY As Long

Public LastOffsetX As Long

Public LastOffsetY As Long

Public End_Time As Long

Public ElapsedTime As Single
'Particulas
'****************************

Public Function EsNPC(ByVal CharIndex As Integer) As Boolean

    'If CharIndex > 0 Then
    If Not CharIndex < 0 Then
        If charlist(CharIndex).iHead = 0 Then
            EsNPC = True
            Exit Function

        End If
    
        EsNPC = False

    End If

End Function

