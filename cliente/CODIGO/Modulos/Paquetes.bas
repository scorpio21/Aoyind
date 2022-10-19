Attribute VB_Name = "Paquetes"
Option Explicit

Public Enum ClientPacketIDGuild

    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    ShowGuildNews

End Enum

Public Enum ClientPacketIDGM

    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    onlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember    '/ACEPTCONSE
    AcceptChaosCouncilMember    '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle    '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    ToggleCentinelActivated    '/CENTINELAACTIVADO
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    AddGM
    CuentaRegresiva         '/CUENTAREGRESIVA

End Enum

Public Enum ServerPacketID

    OpenAccount
    logged                  ' LOGGED
    ChangeHour
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CancelOfferItem
    RecibirRanking
    SendAura                ' Auras
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCSwing                ' N1
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    UserSwing               ' U1
    SafeModeOn              ' SEGON
    SafeModeOff             ' SEGOFF
    ResuscitationSafeOn
    ResuscitationSafeOff
    NobilityLost            ' PN
    CantUseWhileMeditating  ' M!
    Ataca                   ' USER ATACA
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateBankGold
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    ShowMessageScroll
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    CreateEfecto
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    SetInvisible            ' NOVER
    SetOculto               ' NOVER OCULT
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    GuildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildMemberInfo
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    Pong
    UpdateTagAndStatus
    UsersOnline
    CommerceChat
    ShowPartyForm
    StopWorking
    RetosAbre
    RetosRespuesta
    QuestDetails
    QuestListSend
    NpcQuestListSend
    UpdateNPCSimbolo
    SetEquitando            'MONTAR
    SetCongelado            'Congelado
    SetChiquito             'Chiquito
    CreateAreaFX            'Area FX
    PalabrasMagicas         'Palabras magicas
    UserSpellNPC            'US2
    SetNadando
    MultiMessage            'Messages in client
    FirstInfo               'Primera informacion
    CuentaRegresiva         'Cuenta
    PicInRender             'PicInRender
    Quit                    'Quit
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    ShowBarco
    AgregarPasajero
    QuitarPasajero
    QuitarBarco
    GoHome
    GotHome
    Tooltip
    UserInEvent
    EventPacketSv

End Enum

'eventos
Public Enum SvEventPacketID

    SendListEvent = 1
    SendDataEvent = 2

End Enum

'eventos

Public Enum ClientPacketID

    BorrarPJ                'Borrar
    OpenAccount             'ALOGIN
    LoginExistingChar       'OLOGIN
    LoginNewChar            'NLOGIN
    LoginNewAccount         'CALOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    MoveSpell               'DESPHE
    MoveBank
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    CommerceChat
    guild
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    ComandosVarios                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    PartyLeave              '/SALIRPARTY
    PartyCreate             '/CREARPARTY
    PartyJoin               '/PARTY
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    PartyOnline             '/ONLINEPARTY
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    ChangePassword          '/CONTRASEÑA
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARCLAN
    PartyKick               '/ECHARPARTY
    PartySetLeader          '/PARTYLIDER
    PartyAcceptMember       '/ACCEPTPARTY
    Ping                    '/PING
    RequestPartyForm
    ItemUpgrade
    InitCrafting
    RetosAbrir              '/Retos
    RetosCrear
    RetosDecide
    IntercambiarInv
    Quest                   '/QUEST
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    gm
    WarpMeToTarget          '/TELEPLOC
    Home
    Equitar                 'Montar
    DejarMontura            'Dejar Montura
    CreateEfectoClient           'Creamos efecto desde el cliente
    CreateEfectoClientAction     'Luego hacemos la accion (daño, etc)
    AnclarEmbarcacion            'Ancla la embarcacion
    EventPacket
    SolicitaRranking
End Enum

'eventos
Public Enum EventPacketID

    NewEvent = 1
    CloseEvent = 2
    RequiredEvents = 3
    RequiredDataEvent = 4
    ParticipeEvent = 5
    AbandonateEvent = 6

End Enum

'eventos

Public Enum FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_RETOS
    FONTTYPE_EXP

End Enum

Public PaqueteName(0 To 255) As String

Public Sub InitDebug()
    PaqueteName(0) = "OpenAccount"
    PaqueteName(1) = "logged"
    PaqueteName(2) = "ChangeHour"
    PaqueteName(3) = "RemoveDialogs"
    PaqueteName(4) = "RemoveCharDialog"
    PaqueteName(5) = "NavigateToggle"
    PaqueteName(6) = "Disconnect"
    PaqueteName(7) = "CommerceEnd"
    PaqueteName(8) = "BankEnd"
    PaqueteName(9) = "CommerceInit"
    PaqueteName(10) = "BankInit"
    PaqueteName(11) = "UserCommerceInit"
    PaqueteName(12) = "UserCommerceEnd"
    PaqueteName(13) = "UserOfferConfirm"
    PaqueteName(14) = "CancelOfferItem"
    PaqueteName(15) = "SendAura"
    PaqueteName(16) = "ShowBlacksmithForm"
    PaqueteName(17) = "ShowCarpenterForm"
    PaqueteName(18) = "NPCSwing"
    PaqueteName(19) = "NPCKillUser"
    PaqueteName(20) = "BlockedWithShieldUser"
    PaqueteName(21) = "BlockedWithShieldOther"
    PaqueteName(22) = "UserSwing"
    PaqueteName(23) = "SafeModeOn"
    PaqueteName(24) = "SafeModeOff"
    PaqueteName(25) = "ResuscitationSafeOn"
    PaqueteName(26) = "ResuscitationSafeOff"
    PaqueteName(27) = "NobilityLost"
    PaqueteName(28) = "CantUseWhileMeditating"
    PaqueteName(29) = "Ataca"
    PaqueteName(30) = "UpdateSta"
    PaqueteName(31) = "UpdateMana"
    PaqueteName(32) = "UpdateHP"
    PaqueteName(33) = "UpdateGold"
    PaqueteName(34) = "UpdateBankGold"
    PaqueteName(35) = "UpdateExp"
    PaqueteName(36) = "ChangeMap"
    PaqueteName(37) = "PosUpdate"
    PaqueteName(38) = "NPCHitUser"
    PaqueteName(39) = "UserHitNPC"
    PaqueteName(40) = "UserAttackedSwing"
    PaqueteName(41) = "UserHittedByUser"
    PaqueteName(42) = "UserHittedUser"
    PaqueteName(43) = "ChatOverHead"
    PaqueteName(44) = "ConsoleMsg"
    PaqueteName(45) = "GuildChat"
    PaqueteName(46) = "ShowMessageBox"
    PaqueteName(47) = "ShowMessageScroll"
    PaqueteName(48) = "UserIndexInServer"
    PaqueteName(49) = "UserCharIndexInServer"
    PaqueteName(50) = "CharacterCreate"
    PaqueteName(51) = "CharacterRemove"
    PaqueteName(52) = "CharacterMove"
    PaqueteName(53) = "ForceCharMove"
    PaqueteName(54) = "CharacterChange"
    PaqueteName(55) = "ObjectCreate"
    PaqueteName(56) = "ObjectDelete"
    PaqueteName(57) = "BlockPosition"
    PaqueteName(58) = "PlayWave"
    PaqueteName(59) = "guildList"
    PaqueteName(60) = "AreaChanged"
    PaqueteName(61) = "PauseToggle"
    PaqueteName(62) = "RainToggle"
    PaqueteName(63) = "CreateFX"
    PaqueteName(64) = "CreateEfecto"
    PaqueteName(65) = "UpdateUserStats"
    PaqueteName(66) = "WorkRequestTarget"
    PaqueteName(67) = "ChangeInventorySlot"
    PaqueteName(68) = "ChangeBankSlot"
    PaqueteName(69) = "ChangeSpellSlot"
    PaqueteName(70) = "atributes"
    PaqueteName(71) = "BlacksmithWeapons"
    PaqueteName(72) = "BlacksmithArmors"
    PaqueteName(73) = "CarpenterObjects"
    PaqueteName(74) = "RestOK"
    PaqueteName(75) = "ErrorMsg"
    PaqueteName(76) = "Blind"
    PaqueteName(77) = "Dumb"
    PaqueteName(78) = "ShowSignal"
    PaqueteName(79) = "ChangeNPCInventorySlot"
    PaqueteName(80) = "UpdateHungerAndThirst"
    PaqueteName(81) = "Fame"
    PaqueteName(82) = "MiniStats"
    PaqueteName(83) = "LevelUp"
    PaqueteName(84) = "SetInvisible"
    PaqueteName(85) = "MeditateToggle"
    PaqueteName(86) = "BlindNoMore"
    PaqueteName(87) = "DumbNoMore"
    PaqueteName(88) = "SendSkills"
    PaqueteName(89) = "TrainerCreatureList"
    PaqueteName(90) = "guildNews"
    PaqueteName(91) = "OfferDetails"
    PaqueteName(92) = "AlianceProposalsList"
    PaqueteName(93) = "PeaceProposalsList"
    PaqueteName(94) = "CharacterInfo"
    PaqueteName(95) = "GuildLeaderInfo"
    PaqueteName(96) = "GuildMemberInfo"
    PaqueteName(97) = "GuildDetails"
    PaqueteName(98) = "ShowGuildFundationForm"
    PaqueteName(99) = "ParalizeOK"
    PaqueteName(100) = "ShowUserRequest"
    PaqueteName(101) = "TradeOK"
    PaqueteName(102) = "BankOK"
    PaqueteName(103) = "ChangeUserTradeSlot"
    PaqueteName(104) = "Pong"
    PaqueteName(105) = "UpdateTagAndStatus"
    PaqueteName(106) = "UsersOnline"
    PaqueteName(107) = "CommerceChat"
    PaqueteName(108) = "ShowPartyForm"
    PaqueteName(109) = "StopWorking"
    PaqueteName(110) = "RetosAbre"
    PaqueteName(111) = "RetosRespuesta"
    PaqueteName(112) = "SpawnList"
    PaqueteName(113) = "ShowSOSForm"
    PaqueteName(114) = "ShowMOTDEditionForm"
    PaqueteName(115) = "ShowGMPanelForm"
    PaqueteName(116) = "UserNameList"
    PaqueteName(117) = "ShowBarco"
    PaqueteName(118) = "AgregarPasajero"
    PaqueteName(119) = "QuitarPasajero"
    PaqueteName(120) = "QuitarBarco"
    PaqueteName(121) = "GoHome"
    PaqueteName(122) = "GotHome"
    PaqueteName(123) = "Tooltip"
    
End Sub
