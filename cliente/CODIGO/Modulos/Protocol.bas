Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

Private Const MAX_LENGTH As Byte = 120

Public UsersOn           As Integer

Public NotEnoughData     As Boolean
''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR  As String * 1 = vbNullChar

Private Type tFont

    red As Byte
    green As Byte
    blue As Byte
    bold As Boolean
    italic As Boolean

End Type

Public FontTypes(22) As tFont

''
' Initializes the fonts array

Public Sub InitFonts()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .red = 255
        .green = 255
        .blue = 255

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .red = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .red = 32
        .green = 51
        .blue = 223
        .bold = 1
        .italic = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .red = 65
        .green = 190
        .blue = 156

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .red = 65
        .green = 190
        .blue = 156
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .red = 130
        .green = 130
        .blue = 130
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .red = 255
        .green = 180
        .blue = 250

    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .red = 255
        .green = 255
        .blue = 255
        .bold = 1

    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).green = 185
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .red = 228
        .green = 199
        .blue = 27

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .red = 130
        .green = 130
        .blue = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .red = 255
        .green = 60
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .green = 200
        .blue = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .red = 255
        .green = 50
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .green = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .red = 255
        .green = 255
        .blue = 255
        .italic = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .red = 30
        .green = 255
        .blue = 30
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .blue = 200
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .red = 30
        .green = 150
        .blue = 30
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .red = 250
        .green = 250
        .blue = 150
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_RETOS)
        .red = 220
        .green = 220
        .blue = 220
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EXP)
        .red = 220
        .green = 70
        .blue = 0
        .bold = 1

    End With

End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    On Error Resume Next

    Debug.Print incomingData.PeekByte() & " - " & PaqueteName(incomingData.PeekByte()) & " >> " & incomingData.Length

    Select Case incomingData.PeekByte()

    Case ServerPacketID.OpenAccount             ' LOGGED
        Call HandleOpenAccount

    Case ServerPacketID.logged                  ' LOGGED
        Call HandleLogged

    Case ServerPacketID.ChangeHour
        Call HandleChangeHour

    Case ServerPacketID.RemoveDialogs           ' QTDL
        Call HandleRemoveDialogs

    Case ServerPacketID.RemoveCharDialog        ' QDL
        Call HandleRemoveCharDialog

    Case ServerPacketID.NavigateToggle          ' NAVEG
        Call HandleNavigateToggle

    Case ServerPacketID.Disconnect              ' FINOK
        Call HandleDisconnect

    Case ServerPacketID.CommerceEnd             ' FINCOMOK
        Call HandleCommerceEnd

    Case ServerPacketID.CommerceChat
        Call HandleCommerceChat

    Case ServerPacketID.BankEnd                 ' FINBANOK
        Call HandleBankEnd

    Case ServerPacketID.CommerceInit            ' INITCOM
        Call HandleCommerceInit

    Case ServerPacketID.BankInit                ' INITBANCO
        Call HandleBankInit

    Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
        Call HandleUserCommerceInit

    Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
        Call HandleUserCommerceEnd

    Case ServerPacketID.UserOfferConfirm
        Call HandleUserOfferConfirm

    Case ServerPacketID.CancelOfferItem
        Call HandleCancelOfferItem
    Case ServerPacketID.RecibirRanking
        Call HandleRecibirRanking
        'aura
    Case ServerPacketID.SendAura
        Call HandleSendAura
        'aura

    Case ServerPacketID.ShowBlacksmithForm      ' SFH
        Call HandleShowBlacksmithForm

    Case ServerPacketID.ShowCarpenterForm       ' SFC
        Call HandleShowCarpenterForm

    Case ServerPacketID.NPCSwing                ' N1
        Call HandleNPCSwing

    Case ServerPacketID.NPCKillUser             ' 6
        Call HandleNPCKillUser

    Case ServerPacketID.BlockedWithShieldUser   ' 7
        Call HandleBlockedWithShieldUser

    Case ServerPacketID.BlockedWithShieldOther  ' 8
        Call HandleBlockedWithShieldOther

    Case ServerPacketID.UserSwing               ' U1
        Call HandleUserSwing

    Case ServerPacketID.SafeModeOn              ' SEGON
        Call HandleSafeModeOn

    Case ServerPacketID.SafeModeOff             ' SEGOFF
        Call HandleSafeModeOff

    Case ServerPacketID.ResuscitationSafeOff
        Call HandleResuscitationSafeOff

    Case ServerPacketID.ResuscitationSafeOn
        Call HandleResuscitationSafeOn

    Case ServerPacketID.NobilityLost            ' PN
        Call HandleNobilityLost

    Case ServerPacketID.CantUseWhileMeditating  ' M!
        Call HandleCantUseWhileMeditating

    Case ServerPacketID.UpdateSta               ' ASS
        Call HandleUpdateSta

    Case ServerPacketID.UpdateMana              ' ASM
        Call HandleUpdateMana

    Case ServerPacketID.UpdateHP                ' ASH
        Call HandleUpdateHP

    Case ServerPacketID.UpdateGold              ' ASG
        Call HandleUpdateGold

    Case ServerPacketID.UpdateBankGold              ' ASG
        Call HandleUpdateBankGold

    Case ServerPacketID.UpdateExp               ' ASE
        Call HandleUpdateExp

    Case ServerPacketID.ChangeMap               ' CM
        Call HandleChangeMap

    Case ServerPacketID.PosUpdate               ' PU
        Call HandlePosUpdate

    Case ServerPacketID.NPCHitUser              ' N2
        Call HandleNPCHitUser

    Case ServerPacketID.UserHitNPC              ' U2
        Call HandleUserHitNPC

    Case ServerPacketID.UserAttackedSwing       ' U3
        Call HandleUserAttackedSwing

    Case ServerPacketID.UserHittedByUser        ' N4
        Call HandleUserHittedByUser

    Case ServerPacketID.UserHittedUser          ' N5
        Call HandleUserHittedUser

    Case ServerPacketID.ChatOverHead            ' ||
        Call HandleChatOverHead

    Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
        Call HandleConsoleMessage

    Case ServerPacketID.GuildChat               ' |+
        Call HandleGuildChat

    Case ServerPacketID.ShowMessageBox          ' !!
        Call HandleShowMessageBox

    Case ServerPacketID.ShowMessageScroll
        Call HandleShowMessageScroll

    Case ServerPacketID.UserIndexInServer       ' IU
        Call HandleUserIndexInServer

    Case ServerPacketID.UserCharIndexInServer   ' IP
        Call HandleUserCharIndexInServer

    Case ServerPacketID.CharacterCreate         ' CC
        Call HandleCharacterCreate

    Case ServerPacketID.CharacterRemove         ' BP
        Call HandleCharacterRemove

    Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
        Call HandleCharacterMove

    Case ServerPacketID.ForceCharMove
        Call HandleForceCharMove

    Case ServerPacketID.CharacterChange         ' CP
        Call HandleCharacterChange

    Case ServerPacketID.ObjectCreate            ' HO
        Call HandleObjectCreate

    Case ServerPacketID.ObjectDelete            ' BO
        Call HandleObjectDelete

    Case ServerPacketID.BlockPosition           ' BQ
        Call HandleBlockPosition

    Case ServerPacketID.PlayWave                ' TW
        Call HandlePlayWave

    Case ServerPacketID.guildList               ' GL
        Call HandleGuildList

    Case ServerPacketID.AreaChanged             ' CA
        Call HandleAreaChanged

    Case ServerPacketID.PauseToggle             ' BKW
        Call HandlePauseToggle

    Case ServerPacketID.RainToggle              ' LLU
        Call HandleRainToggle

    Case ServerPacketID.CreateFX                ' CFX
        Call HandleCreateFX

    Case ServerPacketID.CreateEfecto                ' CFX
        Call HandleCreateEfecto

    Case ServerPacketID.UpdateUserStats         ' EST
        Call HandleUpdateUserStats

    Case ServerPacketID.WorkRequestTarget       ' T01
        Call HandleWorkRequestTarget

    Case ServerPacketID.ChangeInventorySlot     ' CSI
        Call HandleChangeInventorySlot

    Case ServerPacketID.ChangeBankSlot          ' SBO
        Call HandleChangeBankSlot

    Case ServerPacketID.ChangeSpellSlot         ' SHS
        Call HandleChangeSpellSlot

    Case ServerPacketID.atributes               ' ATR
        Call HandleAtributes

    Case ServerPacketID.BlacksmithWeapons       ' LAH
        Call HandleBlacksmithWeapons

    Case ServerPacketID.BlacksmithArmors        ' LAR
        Call HandleBlacksmithArmors

    Case ServerPacketID.CarpenterObjects        ' OBR
        Call HandleCarpenterObjects

    Case ServerPacketID.RestOK                  ' DOK
        Call HandleRestOK

    Case ServerPacketID.ErrorMsg                ' ERR
        Call HandleErrorMessage

    Case ServerPacketID.Blind                   ' CEGU
        Call HandleBlind

    Case ServerPacketID.Dumb                    ' DUMB
        Call HandleDumb

    Case ServerPacketID.ShowSignal              ' MCAR
        Call HandleShowSignal

    Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
        Call HandleChangeNPCInventorySlot

    Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
        Call HandleUpdateHungerAndThirst

    Case ServerPacketID.Fame                    ' FAMA
        Call HandleFame

    Case ServerPacketID.MiniStats               ' MEST
        Call HandleMiniStats

    Case ServerPacketID.LevelUp                 ' SUNI
        Call HandleLevelUp

    Case ServerPacketID.SetInvisible            ' NOVER
        Call HandleSetInvisible

    Case ServerPacketID.SetOculto               ' NOVER OCULTAR
        Call HandleSetOculto

    Case ServerPacketID.MeditateToggle          ' MEDOK
        Call HandleMeditateToggle

    Case ServerPacketID.BlindNoMore             ' NSEGUE
        Call HandleBlindNoMore

    Case ServerPacketID.Ataca
        Call HandleAtaca

    Case ServerPacketID.DumbNoMore              ' NESTUP
        Call HandleDumbNoMore

    Case ServerPacketID.SendSkills              ' SKILLS
        Call HandleSendSkills

    Case ServerPacketID.TrainerCreatureList     ' LSTCRI
        Call HandleTrainerCreatureList

    Case ServerPacketID.GuildMemberInfo
        Call HandleGuildMemberInfo

    Case ServerPacketID.GuildNews               ' GUILDNE
        Call HandleGuildNews

    Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
        Call HandleOfferDetails

    Case ServerPacketID.AlianceProposalsList    ' ALLIEPR
        Call HandleAlianceProposalsList

    Case ServerPacketID.PeaceProposalsList      ' PEACEPR
        Call HandlePeaceProposalsList

    Case ServerPacketID.CharacterInfo           ' CHRINFO
        Call HandleCharacterInfo

    Case ServerPacketID.GuildLeaderInfo         ' LEADERI
        Call HandleGuildLeaderInfo

    Case ServerPacketID.GuildDetails            ' CLANDET
        Call HandleGuildDetails

    Case ServerPacketID.ShowGuildFundationForm  ' SHOWFUN
        Call HandleShowGuildFundationForm

    Case ServerPacketID.ParalizeOK              ' PARADOK
        Call HandleParalizeOK

    Case ServerPacketID.ShowUserRequest         ' PETICIO
        Call HandleShowUserRequest

    Case ServerPacketID.TradeOK                 ' TRANSOK
        Call HandleTradeOK

    Case ServerPacketID.BankOK                  ' BANCOOK
        Call HandleBankOK

    Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
        Call HandleChangeUserTradeSlot

    Case ServerPacketID.Pong
        Call HandlePong

    Case ServerPacketID.UpdateTagAndStatus
        Call HandleUpdateTagAndStatus

    Case ServerPacketID.UsersOnline
        Call HandleUsersOnline

    Case ServerPacketID.ShowPartyForm
        Call HandleShowPartyForm

    Case ServerPacketID.StopWorking
        Call HandleStopWorking

    Case ServerPacketID.RetosAbre
        Call HandleRetosAbre

    Case ServerPacketID.RetosRespuesta
        Call HandleRetosRespuesta

    Case ServerPacketID.SetEquitando            ' Equitando
        Call HandleSetEquitando

    Case ServerPacketID.SetCongelado            ' Congelado
        Call HandleSetCongelado

    Case ServerPacketID.SetChiquito             ' Chiquito
        Call HandleSetChiquito

    Case ServerPacketID.CreateAreaFX            ' Area CFX
        Call HandleCreateAreaFX

    Case ServerPacketID.PalabrasMagicas
        Call HandlePalabrasMagicas

    Case ServerPacketID.UserSpellNPC            ' U2
        Call HandleUserSpellNPC

    Case ServerPacketID.SetNadando             ' Chiquito
        Call HandleSetNadando

    Case ServerPacketID.MultiMessage            'Messages in client
        Call HandleMultiMessage

    Case ServerPacketID.FirstInfo               'Primera Informacion
        Call HandleFirstInfo

    Case ServerPacketID.CuentaRegresiva         'Cuenta
        Call HandleCuentaRegresiva

    Case ServerPacketID.PicInRender             'PicInRender
        Call HandlePicInRender

    Case ServerPacketID.Quit                    'Quit
        Call HandleQuit

        '*******************
        'GM messages
        '*******************
    Case ServerPacketID.SpawnList               ' SPL
        Call HandleSpawnList

    Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
        Call HandleShowSOSForm

    Case ServerPacketID.ShowMOTDEditionForm     ' ZMOTD
        Call HandleShowMOTDEditionForm

    Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
        Call HandleShowGMPanelForm

    Case ServerPacketID.UserNameList            ' LISTUSU
        Call HandleUserNameList

    Case ServerPacketID.ShowBarco
        Call HandleShowBarco

    Case ServerPacketID.AgregarPasajero
        Call HandleAgregarPasajero

    Case ServerPacketID.QuitarPasajero
        Call HandleQuitarPasajero

    Case ServerPacketID.QuitarBarco
        Call HandleQuitarBarco

    Case ServerPacketID.GoHome
        Call HandleGoHome

    Case ServerPacketID.GotHome
        Call HandleGotHome

    Case ServerPacketID.Tooltip
        Call HandleTooltip

        'quest
    Case ServerPacketID.QuestDetails
        Call HandleQuestDetails

    Case ServerPacketID.QuestListSend
        Call HandleQuestListSend

    Case ServerPacketID.NpcQuestListSend
        Call HandleNpcQuestListSend

    Case ServerPacketID.UpdateNPCSimbolo
        Call HandleUpdateNPCSimbolo
        'quest

        #If SeguridadAlkon Then

        Case Else
            Call HandleIncomingDataEx
        #Else

        Case Else
            'ERROR : Abort!
            incomingData.ReadByte
            Exit Sub
        #End If

    End Select

    'Done with this packet, move on to next one
    If incomingData.Length > 0 And Not NotEnoughData Then
        err.Clear
        Call HandleIncomingData

    End If

End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    ' Variable initialization
    frmMain.pRender.Cls
    EngineRun = True
    GTCPres = (GetTickCount() And &H7FFFFFFF)
    Entrada = 255
    zTick2 = -2147483647
    frmMain.SetRender (False)
    
    Hora = incomingData.ReadByte
    'Call SetDayLight
    'Set connected state
    Call SetConnected

End Sub

Private Sub HandleChangeHour()
    'Remove packet ID
    Call incomingData.ReadByte
    
    Hora = incomingData.ReadByte
    OrigHora = Hora
    Call SetDayLight(incomingData.ReadBoolean)

End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Es medio negro que este aca pero viene bien..
    Call LimpiarArea
    
    Call Dialogos.RemoveAllDialogs

End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check if the packet is complete
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())

End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando

End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim I As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Close connection
    CloseSock
    
    AlphaSalir = 0

    Call ClosePj

End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Reset vars
    Comerciando = False
    
    'Hide form
    Unload frmComerciar

End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    '
    '***************************************************
    If incomingData.Length < 4 Then
        err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat      As String

    Dim FontIndex As Integer

    Dim str       As String

    Dim R         As Byte

    Dim G         As Byte

    Dim b         As Byte
    
    chat = buffer.ReadASCIIString()
    FontIndex = buffer.ReadByte()
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            R = 255
        Else
            R = Val(str)

        End If
            
        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            G = 255
        Else
            G = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            b = 255
        Else
            b = Val(str)

        End If
            
        Call frmComerciarUsu.PrintCommerceMsg(left$(chat, InStr(1, chat, "~") - 1), 1)
        'Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else

        With FontTypes(FontIndex)
            Call frmComerciarUsu.PrintCommerceMsg(chat, 1)

            'frmComerciarUsu.CommerceConsole.Text = frmComerciarUsu.CommerceConsole.Text & chat & vbCrLf
            'Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, chat, .red, .green, .blue, .bold, .italic)
        End With

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    
    Unload frmBancoObj
    Comerciando = False

End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim I As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(frmComerciar.picInvUser, 0, 0, MAX_INVENTORY_SLOTS, False)  '
    Call InvComNpc.Initialize(frmComerciar.picInvNpc, 0, 0, MAX_NPC_INVENTORY_SLOTS, False)    ')

    'Fill user inventory
    For I = 1 To MAX_INVENTORY_SLOTS

        If Inventario.OBJIndex(I) <> 0 Then

            With Inventario
                Call InvComUsu.SetItem(I, .OBJIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), .PuedeUsarItem(I))

            End With

        End If

    Next I
    
    ' Fill Npc inventory
    For I = 1 To 50

        If NPCInventory(I).OBJIndex <> 0 Then

            With NPCInventory(I)
                Call InvComNpc.SetItem(I, .OBJIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .PuedeUsarItem)

            End With

        End If

    Next I
    
    'Set state and show form
    Comerciando = True
    frmComerciar.Show , frmMain
    
    'Reproducimos el saludo del comerciante
    Call Audio.PlayWave(RandomNumber(241, 245))

End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim I        As Long

    Dim BankGold As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    BankGold = incomingData.ReadLong
    Call InvBanco(0).Initialize(frmBancoObj.PicBancoInv, 0, 0, MAX_BANCOINVENTORY_SLOTS, False)
    Call InvBanco(1).Initialize(frmBancoObj.picInv, 0, 0, Inventario.MaxObjs, False)
    
    For I = 1 To Inventario.MaxObjs

        With Inventario
            Call InvBanco(1).SetItem(I, .OBJIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), .PuedeUsarItem(I))

        End With

    Next I
    
    For I = 1 To MAX_BANCOINVENTORY_SLOTS

        With UserBancoInventory(I)
            Call InvBanco(0).SetItem(I, .OBJIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .PuedeUsarItem)

        End With

    Next I
    
    'Set state and show form
    Comerciando = True
    
    frmBancoObj.lblUserGld.Caption = BankGold
    
    frmBancoObj.Show , frmMain
    
    'Reproducimos el saludo del banquero
    Call Audio.PlayWave(RandomNumber(242, 245))

End Sub

' Handles the StopWorking message.
Private Sub HandleStopWorking()
    '***************************************************
    'Author: Budi
    'Last Modification: 12/01/09
    '
    '***************************************************

    Call incomingData.ReadByte
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("¡Has terminado de trabajar!", .red, .green, .blue, .bold, .italic)

    End With
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo

End Sub

' Handles the CancelOfferItem message.

Private Sub HandleCancelOfferItem()

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 05/03/10
    '
    '***************************************************
    Dim slot   As Byte

    Dim Amount As Long
    
    Call incomingData.ReadByte
    
    slot = incomingData.ReadByte
    
    With InvOfferComUsu(0)
        Amount = .Amount(slot)
        
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            ' Actualizo el inventario general
            Call frmComerciarUsu.UpdateInvCom(.OBJIndex(slot), Amount)
            
            ' Borro el item
            Call .SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)

        End If

    End With
    
    ' Si era el único ítem de la oferta, no puede confirmarla
    If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And Not frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) Then Call frmComerciarUsu.HabilitarConfirmar(False)
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call frmComerciarUsu.PrintCommerceMsg("¡No puedes comerciar ese objeto!", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim I As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    TradingUserName = incomingData.ReadASCIIString
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(frmComerciarUsu.picInvComercio, 0, 0, Inventario.MaxObjs, False)
    Call InvOfferComUsu(0).Initialize(frmComerciarUsu.picInvOfertaProp, 0, 0, INV_OFFER_SLOTS, False)
    Call InvOfferComUsu(1).Initialize(frmComerciarUsu.picInvOfertaOtro, 0, 0, INV_OFFER_SLOTS, False)
    Call InvOroComUsu(0).Initialize(frmComerciarUsu.picInvOroProp, 0, 0, INV_GOLD_SLOTS, False, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(1).Initialize(frmComerciarUsu.picInvOroOfertaProp, 0, 0, INV_GOLD_SLOTS, False, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(2).Initialize(frmComerciarUsu.picInvOroOfertaOtro, 0, 0, INV_GOLD_SLOTS, False, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)

    'Fill user inventory
    For I = 1 To MAX_INVENTORY_SLOTS

        If Inventario.OBJIndex(I) <> 0 Then

            With Inventario
                Call InvComUsu.SetItem(I, .OBJIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), .PuedeUsarItem(I))

            End With

        End If

    Next I

    ' Inventarios de oro
    Call InvOroComUsu(0).SetItem(1, ORO_INDEX, UserGLD, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro", 1)
    Call InvOroComUsu(1).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro", 1)
    Call InvOroComUsu(2).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro", 1)

    'Set state and show form
    Comerciando = True
    Call frmComerciarUsu.Show(vbModeless, frmMain)

End Sub

''
' Handles the UserOfferConfirm message.
Private Sub HandleUserOfferConfirm()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    With frmComerciarUsu
        ' Now he can accept the offer or reject it
        .HabilitarAceptarRechazar True
        
        .PrintCommerceMsg TradingUserName & " ha confirmado su oferta!", FontTypeNames.FONTTYPE_CONSE

    End With
    
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvOroComUsu(0) = Nothing
    Set InvOroComUsu(1) = Nothing
    Set InvOroComUsu(2) = Nothing
    Set InvOfferComUsu(0) = Nothing
    Set InvOfferComUsu(1) = Nothing
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False

End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
    Else
        frmHerrero.Show , frmMain

    End If

End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftCarpenter(MacroBltIndex)
    Else
        frmCarp.Show , frmMain

    End If

End Sub

''
' Handles the NPCSwing message.

Private Sub HandleNPCSwing()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 2 Then
        NotEnoughData = True
        Exit Sub

    End If

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim NpcType As Byte

    NpcType = incomingData.ReadByte

    Select Case NpcType

        Case eNPCType.Guardia
            Call AddtoRichPicture("El guardia" & MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)

        Case eNPCType.Montura
            Call AddtoRichPicture("El caballo salvaje" & MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)

        Case Else
            Call AddtoRichPicture("La criatura" & MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)

    End Select
    
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    If incomingData.Length < 2 Then
        NotEnoughData = True
        Exit Sub

    End If

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim NpcType As Byte

    NpcType = incomingData.ReadByte()

    Select Case NpcType

        Case eNPCType.Guardia
            Call AddtoRichPicture("El guardia " & MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)

        Case eNPCType.Montura
            Call AddtoRichPicture("El animal salvaje " & MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)

        Case Else
            Call AddtoRichPicture("La criatura " & MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)

    End Select
    
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichPicture(MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)

End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichPicture(MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)

End Sub

''
' Handles the UserSwing message.

Private Sub HandleUserSwing()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim c As New clsToolTip

    Call c.Init(UserPos.X, UserPos.Y, "Fallas!", 200, 10, 10)
    Tooltips.Add c
    
    Call AddtoRichPicture(MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)

End Sub

''
' Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
   
    UserSeguro = True
  SeguroConIma = 14814
    
    frmMain.PicSeg.Picture = LoadPictureEX("barraSeguroOn.jpg")
    
    Call AddtoRichPicture(MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)
  
End Sub

''
' Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
   
    UserSeguro = False
   SeguroConIma = 14813
    frmMain.PicSeg.Picture = LoadPictureEX("barraSeguro.jpg")
   
    Call AddtoRichPicture(MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, False)

End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandleResuscitationSafeOff()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSeguroResu(False)
    UserSeguroResu = False
    
    frmMain.picResu.Picture = LoadPictureEX("barraResuOff.jpg")
    
    Call AddtoRichPicture(MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, False)

End Sub

''
' Handles the ResuscitationSafeOn message.

Private Sub HandleResuscitationSafeOn()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call frmMain.ControlSeguroResu(True)
    
    UserSeguroResu = True
    
    frmMain.picResu.Picture = LoadPictureEX("barraResuOn.jpg")
    
    Call AddtoRichPicture(MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, False)

End Sub

''
' Handles the NobilityLost message.

Private Sub HandleNobilityLost()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichPicture(MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)

End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichPicture(MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)

End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
    '    frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 100)
    '    frmMain.lblEnergiaN.Caption = UserMinSTA & "/" & UserMaxSTA
    '    frmMain.lblEnergia.Caption = frmMain.lblEnergiaN.Caption

End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinMAN = incomingData.ReadInteger()
    
    If UserMaxMAN > 0 Then
'        frmMain.Bar_Mana(0).max = UserMaxMAN
'        frmMain.Bar_Mana(0).value = UserMinMAN
'        frmMain.Bar_Mana(0).TextAfterCaption = " / " & UserMaxMAN
'        frmMain.Bar_Mana(1).max = UserMaxMAN
'        frmMain.Bar_Mana(1).value = UserMinMAN
'        frmMain.Bar_Mana(1).TextAfterCaption = " / " & UserMaxMAN
    Else
'        frmMain.Bar_Mana(0).max = 1
'        frmMain.Bar_Mana(0).value = 0
'        frmMain.Bar_Mana(0).TextAfterCaption = " / 0"
'        frmMain.Bar_Mana(1).max = 1
'        frmMain.Bar_Mana(1).value = 0
'        frmMain.Bar_Mana(1).TextAfterCaption = " / 0"

    End If

End Sub

Private Sub HandleShowBarco()

    '***************************************************
    'Author: Javier Podavini (El Yind)
    'Last Modification: 15/03/2012
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 43 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X                 As Integer, Y As Integer

    Dim Paso              As Byte

    Dim TiempoPuerto      As Long

    Dim Sentido           As Byte

    Dim Pasajeros(1 To 4) As Integer ' cantidad de user en el barco

    Dim I                 As Integer

    Dim Body              As Integer, Head As Integer, Casco As Integer
    
    Sentido = incomingData.ReadByte()

    Paso = incomingData.ReadByte()
    X = incomingData.ReadInteger()
    Y = incomingData.ReadInteger()
    TiempoPuerto = incomingData.ReadLong()
    
    Debug.Print "Paquete Barco"
    
    For I = 1 To 4 ' cantidad de user en el barco
        Pasajeros(I) = incomingData.ReadInteger()
        Body = incomingData.ReadInteger()
        Head = incomingData.ReadInteger()
        Casco = incomingData.ReadInteger()

        If Pasajeros(I) > 0 Then

            With charlist(Pasajeros(I))
                .Body = BodyData(Body)
                .Head = HeadData(Head)
                .Casco = CascoAnimData(Casco)
                .Alpha = 255

            End With

        End If

    Next I
    
    Call Dialogos.RemoveDialog(9999)
    Call Dialogos.RemoveDialog(10000)
    
    Set Barco(Sentido) = New clsBarco
    Call Barco(Sentido).Init(RutaBarco(Sentido), Paso, X, Y, TiempoPuerto, Sentido, Pasajeros)

End Sub

Private Sub HandleQuitarBarco()

    '***************************************************
    'Author: Javier Podavini (El Yind)
    'Last Modification: 15/03/2012
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 2 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    Dim Paso    As Byte

    Dim Sentido As Byte

    Sentido = incomingData.ReadByte()
    
    Set Barco(Sentido) = Nothing

End Sub

Private Sub HandleAgregarPasajero()

    '***************************************************
    'Author: Javier Podavini (El Yind)
    'Last Modification: 15/03/2012
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 11 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex    As Integer

    Dim Paso         As Byte

    Dim TiempoPuerto As Long

    Dim Sentido      As Byte

    Dim Lugar        As Byte

    Dim Head         As Integer

    Dim Body         As Integer

    Dim Casco        As Integer
    
    Sentido = incomingData.ReadByte()
    CharIndex = incomingData.ReadInteger()
    Lugar = incomingData.ReadByte()
    Head = incomingData.ReadInteger()
    Body = incomingData.ReadInteger()
    Casco = incomingData.ReadInteger()

    If Not Barco(Sentido) Is Nothing Then
        
        charlist(CharIndex).Head = HeadData(Head)
        charlist(CharIndex).Body = BodyData(Body)
        charlist(CharIndex).Casco = CascoAnimData(Casco)
        charlist(CharIndex).Arma = WeaponAnimData(2)
        charlist(CharIndex).Escudo = ShieldAnimData(2)

        Call Barco(Sentido).AgregarPasajero(Lugar, CharIndex)

    End If

End Sub

Private Sub HandleQuitarPasajero()

    '***************************************************
    'Author: Javier Podavini (El Yind)
    'Last Modification: 15/03/2012
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    Dim Paso    As Byte

    Dim Sentido As Byte

    Dim Lugar   As Byte
    
    Sentido = incomingData.ReadByte()
    Lugar = incomingData.ReadByte()

    If Not Barco(Sentido) Is Nothing Then
        Call Barco(Sentido).QuitarPasajero(Lugar)

    End If

End Sub

Private Sub HandleGoHome()

    '***************************************************
    'Author: Javier Podavini (El Yind)
    'Last Modification: 24/05/2012
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    TiempoHome = incomingData.ReadInteger()
    GoingHome = 1
     

    
    Call General_Char_Particle_Create(110, UserCharIndex)
 SetDayLight
   


    AngMareoMuerto = 0
    RadioMareoMuerto = 0

    AddtoRichPicture "Tu espíritu se libera. Comienzas un trance hacia tu hogar que durará " & TiempoHome & " segundos.", 30, 205, 35, True, True, False
    
End Sub

Private Sub HandleGotHome()

    '***************************************************
    'Author: Javier Podavini (El Yind)
    'Last Modification: 24/05/2012
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 2 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    If incomingData.ReadBoolean Then
        AddtoRichPicture "¡Tu espíritu se ha transportado a su hogar!", 30, 205, 35, True, True, False
    Else
        AddtoRichPicture "¡Tu trance se ha interrumpido!", 30, 205, 35, True, True, False

    End If
 charlist(UserCharIndex).particle_count = 0
 Hora = 10
 GoingHome = 2
 SetDayLight True
    
    
End Sub

Private Sub HandleTooltip()

    '***************************************************
    'Author: Javier Podavini (El Yind)
    'Last Modification: 26/05/2012
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 6 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    Dim X       As Integer

    Dim Y       As Integer

    Dim Mensaje As String

    Dim TIPO    As Byte

    Dim R       As Byte, G As Byte, b As Byte
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    X = incomingData.ReadInteger
    Y = incomingData.ReadInteger
    
    TIPO = incomingData.ReadByte
    
    Mensaje = incomingData.ReadASCIIString()
    
    If TIPO = 0 Then
        R = 255
        G = 0
        b = 0
    ElseIf TIPO = 1 Then
        R = 250
        G = 201
        b = 20
    ElseIf TIPO = 2 Then ' PESCA
        R = 153
        G = 217
        b = 234
    ElseIf TIPO = 3 Then ' MINA
        R = 128
        G = 128
        b = 128
    ElseIf TIPO = 4 Then ' TALA
        R = 185
        G = 122
        b = 87
    ElseIf TIPO = 5 Then ' HERRERIA DAGA
        R = 185
        G = 222
        b = 87
    ElseIf TIPO = 6 Then ' SUBE HP
        R = 255
        G = 119
        b = 119
    ElseIf TIPO = 7 Then ' SUBE MANA
        R = 128
        G = 255
        b = 255
    ElseIf TIPO = 8 Then ' SUBE STA
        R = 255
        G = 255
        b = 81
    ElseIf TIPO = 9 Then ' SUBE FUERZA
        R = 17
        G = 153
        b = 17
    ElseIf TIPO = 10 Then ' SUBE AGILIDAD
        R = 255
        G = 255
        b = 81

    End If
    
    Dim c As New clsToolTip

    Call c.Init(X, Y, Mensaje, R, G, b)
    Tooltips.Add c
    
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinHP = incomingData.ReadInteger()
'
'    frmMain.bar_salud(0).max = UserMaxHP
'    frmMain.bar_salud(0).value = UserMinHP
'    frmMain.bar_salud(1).max = UserMaxHP
'    frmMain.bar_salud(1).value = UserMinHP
'
'    frmMain.bar_salud(0).TextAfterCaption = " / " & UserMaxHP
'    frmMain.bar_salud(1).TextAfterCaption = " / " & UserMaxHP
    
    ' frmMain.HpshpV.Value = (((UserMinHP / 100) / (UserMaxHP / 100)) * 100)
    
    '    frmMain.lblVidaN.Caption = UserMinHP & "/" & UserMaxHP
    '    frmMain.lblVida.Caption = frmMain.lblVidaN.Caption
    '
    'Is the user alive??
    If UserMinHP = 0 Then
        UserEstado = 1
        
        'Sacamos un screenshot si está activado el FragShooter:
        If mOpciones.ScreenShooterAlMorir = True Then
            ScreenShooterCapturePending = True

        End If
        
        If frmMain.macrotrabajo Then frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0

    End If

End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    '- 08/14/07: Added GldLbl color variation depending on User Gold and Level
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserGLD = incomingData.ReadLong()

    If UserGLD <> 0 Then
        frmMain.GldLbl.Caption = Format$(UserGLD, "##,##")
    Else
        frmMain.GldLbl.Caption = UserGLD

    End If

End Sub

''
' Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmBancoObj.lblUserGld.Caption = incomingData.ReadLong
    
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserExp = incomingData.ReadLong()

    'frmMain.iBEXP.ToolTipText = "Exp: " & UserExp & "/" & UserPasarNivel
    'frmMain.iBEXP.Width = 122 * Round(CDbl(UserExp) / CDbl(UserPasarNivel), 2) + 4
    ' frmMain.Experiencia.Left = frmMain.iBEXP.Left + frmMain.iBEXP.Width
    ''frmMain.exp.Caption = Round((UserExp / UserPasarNivel) * 100, 2) & "%"
    'frmMain.exp.Caption = Round((UserExp / UserPasarNivel) * 100, 2) & "%"
    'neo sacados
End Sub

''
' Handles the ChangeMap message.

Private Sub HandleChangeMap()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 2 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadByte()
        
    #If SeguridadAlkon Then
        Call InitMI
    #End If
    
    If FileExist(DirRecursos & "Mapa" & UserMap & ".AO", vbNormal) Then
        Call SwitchMap(UserMap)
    Else
        'no encontramos el mapa en el hd
        MessageBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        
        Call CloseClient

    End If

End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
                
    MainTimer.Restart (TimersIndex.PuedeRPUMover)
    MainTimer.Restart (TimersIndex.SendRPU)
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Remove char from old position
    If MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex Then
        MapData(UserPos.X, UserPos.Y).CharIndex = 0

    End If
    
    UserPos.X = incomingData.ReadInteger
    UserPos.Y = incomingData.ReadInteger
    
    'Set char
    MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
    charlist(UserCharIndex).Pos = UserPos
    
    WAIT_ACTION = eWAIT_FOR_ACTION.None
    
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 50 Or MapData(UserPos.X, UserPos.Y).Trigger = 7 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
                
    'Update pos label
    frmMain.Coord.Caption = "(" & UserPos.X & "," & UserPos.Y & ")"

    'CheckZona
End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 5 Then
        NotEnoughData = True
        Exit Sub

    End If

    Dim Golpe   As Integer

    Dim Lugar   As Byte

    Dim NpcType As Byte

    Dim str     As String

    'Remove packet ID
    Call incomingData.ReadByte
    Lugar = incomingData.ReadByte()
    Golpe = incomingData.ReadInteger
    NpcType = incomingData.ReadByte()
    
    Select Case Lugar

        Case bCabeza
            str = MENSAJE_GOLPE_CABEZA

        Case bBrazoIzquierdo
            str = MENSAJE_GOLPE_BRAZO_IZQ

        Case bBrazoDerecho
            str = MENSAJE_GOLPE_BRAZO_DER

        Case bPiernaIzquierda
            str = MENSAJE_GOLPE_PIERNA_IZQ

        Case bPiernaDerecha
            str = MENSAJE_GOLPE_PIERNA_DER

        Case bTorso
            str = MENSAJE_GOLPE_TORSO

    End Select
    
    Select Case NpcType

        Case eNPCType.Guardia
            Call AddtoRichPicture("El guardia" & str & CStr(Golpe & "!!"), 255, 0, 0, True, False, False)

        Case eNPCType.Montura
            Call AddtoRichPicture("El animal salvaje" & str & CStr(Golpe & "!!"), 255, 0, 0, True, False, False)

        Case Else
            Call AddtoRichPicture("La criatura" & str & CStr(Golpe & "!!"), 255, 0, 0, True, False, False)

    End Select
    
    Dim c As New clsToolTip

    Call c.Init(UserPos.X, UserPos.Y, Golpe, 200, 10, 10)
    Tooltips.Add c

End Sub

''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 7 Then
        NotEnoughData = True
        Exit Sub

    End If

    Dim Golpe As Integer

    Dim CharIndex As Integer

    Dim HitArea As Byte

    Dim NpcType As Byte

    'Remove packet ID
    Call incomingData.ReadByte

    Golpe = incomingData.ReadInteger()
    CharIndex = incomingData.ReadInteger()
    HitArea = incomingData.ReadByte()
    NpcType = incomingData.ReadByte()
    Colorsangre = incomingData.ReadByte()
    If HitArea > 0 Then
        Call AddtoRichPicture("La onda expansiva ha alcanzado a la criatura.", 255, 0, 0, False, False, False)

    End If

    Select Case NpcType

    Case eNPCType.Guardia
        Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_1 & "al guardia por " & CStr(Golpe) & MENSAJE_2, 255, 0, 0, True, False, False)

    Case eNPCType.Montura
        Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_1 & "al animal salvaje por " & CStr(Golpe) & MENSAJE_2, 255, 0, 0, True, False, False)

    Case Else
        Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_1 & "a la criatura por " & CStr(Golpe) & MENSAJE_2, 255, 0, 0, True, False, False)

    End Select

    If charlist(CharIndex).Pos.X > 0 Then
        'sangre
        Dim CorSanX As Integer
        Dim CorSanY As Integer
        Select Case charlist(UserCharIndex).Heading
        Case E_Heading.north
            CorSanX = 0
            CorSanY = -30
        Case E_Heading.east
            CorSanX = 30
            CorSanY = 0
        Case E_Heading.south
            CorSanX = 0
            CorSanY = 30
        Case E_Heading.west
            CorSanX = -30
            CorSanY = 0

        End Select

        Effect_BloodSpray_Begin ParticulaX(charlist(CharIndex).Pos.X, charlist(UserCharIndex).Pos.X) + CorSanX, ParticulaY(charlist(CharIndex).Pos.Y, charlist(UserCharIndex).Pos.Y) + CorSanY, 7 + Rnd * 10, 200, 1


        'sangre

        Dim c As New clsToolTip

        Call c.Init(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y, Golpe, 200, 10, 10)
        Tooltips.Add c

    End If

End Sub

Private Sub HandleUserSpellNPC()

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 7 Then
        NotEnoughData = True
        Exit Sub

    End If

    Dim Golpe As Integer

    Dim CharIndex As Integer

    Dim HitArea As Byte

    Dim NpcType As Byte

    'Remove packet ID
    Call incomingData.ReadByte

    Golpe = incomingData.ReadInteger()
    CharIndex = incomingData.ReadInteger()
    HitArea = incomingData.ReadByte()
    NpcType = incomingData.ReadByte()
    Colorsangre = incomingData.ReadByte()

    If HitArea > 0 Then
        Call AddtoRichPicture("La onda expansiva ha alcanzado a la criatura.", 255, 0, 0, False, False, False)

    End If

    Select Case NpcType

    Case eNPCType.Guardia
        Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_HECHIZO & CStr(Golpe) & " puntos de daño al guardia" & "!" & MENSAJE_2, 255, 0, 0, True, False, False)

    Case eNPCType.Montura
        Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_HECHIZO & CStr(Golpe) & " puntos de daño a la criatura" & "!" & MENSAJE_2, 255, 0, 0, True, False, False)

    Case Else
        Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_HECHIZO & CStr(Golpe) & " puntos de daño a la criatura" & "!" & MENSAJE_2, 255, 0, 0, True, False, False)

    End Select

    If charlist(CharIndex).Pos.X > 0 Then
        'sangre
        Effect_BloodSpray_Begin (ParticulaX(charlist(CharIndex).Pos.X, charlist(UserCharIndex).Pos.X)), (ParticulaY(charlist(CharIndex).Pos.Y, charlist(UserCharIndex).Pos.Y)), 7 + Rnd * 10, 200, 1
        'sangre
        Dim c As New clsToolTip

        Call c.Init(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y, Golpe, 200, 10, 10)
        Tooltips.Add c

    End If

End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichPicture(MENSAJE_1 & charlist(incomingData.ReadInteger()).nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
    
End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 7 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim attacker As String

    Dim Golpe    As Integer

    Dim Lugar    As Byte

    Dim HitArea  As Byte

    attacker = charlist(incomingData.ReadInteger()).nombre
    Lugar = incomingData.ReadByte
    Golpe = incomingData.ReadInteger()
    HitArea = incomingData.ReadByte
    
    If HitArea > 0 Then
        Call AddtoRichPicture("La onda expansiva te ha alcanzado!", 255, 0, 0, False, False, False)

    End If
    
    Select Case Lugar

        Case bCabeza
            Call AddtoRichPicture(MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_CABEZA & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichPicture(MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichPicture(MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichPicture(MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichPicture(MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichPicture(MENSAJE_1 & attacker & MENSAJE_RECIVE_IMPACTO_TORSO & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

    End Select
    
    Dim c As New clsToolTip

    Call c.Init(UserPos.X, UserPos.Y - 2, Golpe, 200, 10, 10)
    Tooltips.Add c

End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 9 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim victim    As String

    Dim Golpe     As Integer

    Dim CharIndex As Integer

    Dim Lugar     As Byte

    Dim HitArea   As Byte
    
    CharIndex = incomingData.ReadInteger()
    
    victim = charlist(CharIndex).nombre
    Lugar = incomingData.ReadByte
    Golpe = incomingData.ReadInteger()
    HitArea = incomingData.ReadByte()
    
    If HitArea > 0 Then
        Call AddtoRichPicture("La onda expansiva ha alcanzado a " & victim & "!", 255, 0, 0, False, False, False)

    End If
        
    Select Case Lugar

        Case bCabeza
            Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichPicture(MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_TORSO & CStr(Golpe & MENSAJE_2), 255, 0, 0, True, False, False)

    End Select
    
    Dim c As New clsToolTip

    Call c.Init(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y - 2, Golpe, 200, 10, 10)
    Tooltips.Add c

End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 8 Then
        NotEnoughData = True
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Dim chat As String

    Dim CharIndex As Integer

    Dim R As Byte

    Dim G As Byte

    Dim b As Byte
    Dim FoundIt As Integer
    chat = buffer.ReadASCIIString()
    CharIndex = buffer.ReadInteger()

    R = buffer.ReadByte()
    G = buffer.ReadByte()
    b = buffer.ReadByte()

    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(CharIndex).ACTIVE Or charlist(CharIndex).priv = 10 Then Call Dialogos.CreateDialog(chat, CharIndex, R, G, b)

    Dim Name As String

    Dim Pos As Integer

    Pos = InStr(charlist(CharIndex).nombre, "<")

    If Pos = 0 Then Pos = Len(charlist(CharIndex).nombre) + 2

    Name = left$(charlist(CharIndex).nombre, Pos - 2)


    FoundIt = InStr(1, Name, "!")

    If FoundIt <> 0 Then
        Name = Replace(Name, "!", "")
    Else
        Name = Name
    End If
    If charlist(CharIndex).nombre <> "" And charlist(CharIndex).invisible = False Then Call AddtoRichPicture(Name & "> " & chat, 190, 190, 190, False, True)

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 4 Then
        NotEnoughData = True
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Dim chat      As String

    Dim FontIndex As Integer

    Dim str       As String

    Dim R         As Byte

    Dim G         As Byte

    Dim b         As Byte

    chat = buffer.ReadASCIIString()
    FontIndex = buffer.ReadByte()

    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            R = 255
        Else
            R = Val(str)

        End If

        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            G = 255
        Else
            G = Val(str)

        End If

        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            b = 255
        Else
            b = Val(str)

        End If

        Call AddtoRichPicture(left$(chat, InStr(1, chat, "~") - 1), R, G, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        #If RenderFull = 0 Then

            With FontTypes(FontIndex)

                Dim cantidad As Integer

                Dim TEXTOX() As String

                Dim J        As Integer

                TEXTOX = FormatChat(chat)
                cantidad = UBound(TEXTOX())

                If cantidad = 0 Then
                    'Call RenderTextCentered(.x, .y, Trim$(.textLine(0)), .color)

                    Call AddtoRichPicture(Trim$(TEXTOX(0)), .red, .green, .blue, .bold, .italic)
                Else

                    For J = 0 To cantidad
                        'Call RenderText(.x - 50, .y + offset, Trim$(.textLine(J)), .color)
                        Call AddtoRichPicture(Trim$(TEXTOX(J)), .red, .green, .blue, .bold, .italic)
                        'offset = offset + usedFont.Size + 5
                    Next J

                End If

            End With

        #Else

            With FontTypes(FontIndex)
                Call AddtoRichPicture(chat, .red, .green, .blue, .bold, .italic)

            End With

        #End If

    End If

    If mOpciones.ScreenShooterNivelSuperior = True Then
        Call checkText(chat)

    End If

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    ' If error <> 0 Then _
    '   Err.Raise error
End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/07/08 (NicoNZ)
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String

    Dim str  As String

    Dim R    As Byte

    Dim G    As Byte

    Dim b    As Byte

    Dim tmp  As Integer

    Dim Cont As Integer
    
    chat = buffer.ReadASCIIString()
    
    'If Not DialogosClanes.Activo Then
    If mOpciones.DialogConsole = False Then
        If InStr(1, chat, "~") Then
            str = ReadField(2, chat, 126)

            If Val(str) > 255 Then
                R = 255
            Else
                R = Val(str)

            End If
            
            str = ReadField(3, chat, 126)

            If Val(str) > 255 Then
                G = 255
            Else
                G = Val(str)

            End If
            
            str = ReadField(4, chat, 126)

            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)

            End If
            
            Call AddtoRichPicture(left$(chat, InStr(1, chat, "~") - 1), R, G, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        Else

            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call AddtoRichPicture(chat, .red, .green, .blue, .bold, .italic)

            End With

        End If

    Else
        Call DialogosClanes.PushBackText(ReadField(1, chat, 126))

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    frmMensaje.msg.Caption = buffer.ReadASCIIString()
    frmMensaje.Show
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ShowMessageScroll message.

Private Sub HandleShowMessageScroll()

    If incomingData.Length < 6 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call DrawTextPergamino(buffer.ReadASCIIString(), buffer.ReadInteger, buffer.ReadByte)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()
    
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If

    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCharIndex = incomingData.ReadInteger()
    
    With charlist(UserCharIndex)
        UserPos = .Pos
        .LastPos = .Pos

    End With
    
    CheckZona
    SetDayLight (False)

    'Call CargarTile(UserPos.X, UserPos.Y)
    'Are we under a roof?
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 50 Or MapData(UserPos.X, UserPos.Y).Trigger = 7 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

    If bTecho Then
        bAlpha = 70
        ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, bAlpha)

    End If
    
    frmMain.Coord.Caption = "(" & UserPos.X & "," & UserPos.Y & ")"
    Call frmMain.RefreshMiniMap

End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 24 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim CharIndex As Integer

    Dim Body      As Integer

    Dim Head      As Integer

    Dim Heading   As E_Heading

    Dim X         As Integer

    Dim Y         As Integer

    Dim weapon    As Integer

    Dim shield    As Integer

    Dim helmet    As Integer

    Dim privs     As Integer
    
    CharIndex = buffer.ReadInteger()
    Body = buffer.ReadInteger()
    Head = buffer.ReadInteger()
    Heading = buffer.ReadByte()
    X = buffer.ReadInteger()
    Y = buffer.ReadInteger()
    weapon = buffer.ReadInteger()
    shield = buffer.ReadInteger()
    helmet = buffer.ReadInteger()
    
    With charlist(CharIndex)
        Call SetCharacterFx(CharIndex, buffer.ReadInteger(), buffer.ReadInteger())
        
        .nombre = buffer.ReadASCIIString()
        .Criminal = buffer.ReadByte()
        
        privs = buffer.ReadByte()
        .simbolo = buffer.ReadByte()

        If privs <> 0 Then

            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil

            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil

            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster

            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0

        End If
            'Recibimos alas..
        .alaIndex = buffer.ReadByte()
        
        If .alaIndex <> 0 Then
        .Alas = alaArray(.alaIndex)
        End If

    End With
    
    Call MakeChar(CharIndex, Body, Head, Heading, X, Y, weapon, shield, helmet)
    
    'Call RefreshAllChars
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    If charlist(CharIndex).nombre = "" And charlist(CharIndex).Pos.X > 0 And charlist(CharIndex).Pos.Y > 0 Then

        Dim tN As New clsNPCMuerto

        With charlist(CharIndex)
            Call tN.Init(.iBody, .iHead, 2, 2, 2, .Heading, .Pos, .Alpha)
            Set MapData(.Pos.X, .Pos.Y).Elemento = tN

        End With
        
        NPCMuertos.Add tN

    End If
    
    Call EraseChar(CharIndex)
    Call RefreshAllChars

End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 5 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer

    Dim X         As Integer

    Dim Y         As Integer
    
    CharIndex = incomingData.ReadInteger()
    X = incomingData.ReadInteger()
    Y = incomingData.ReadInteger()
    
    With charlist(CharIndex)

        If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
            .FxIndex = 0

        End If
        
        ' Play steps sounds if the user is not an admin of any kind
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(CharIndex)

        End If

    End With
        
    Call MoveCharbyPos(CharIndex, X, Y)
    
    'Call RefreshAllChars
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    If incomingData.Length < 2 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call MoveCharbyHead(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    'Call RefreshAllChars
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 18 Then
        NotEnoughData = True
        Exit Sub

    End If

    'Remove packet ID
    Call incomingData.ReadByte

    Dim CharIndex As Integer

    Dim tempint As Integer

    Dim headIndex As Integer

    CharIndex = incomingData.ReadInteger()

    With charlist(CharIndex)
        tempint = incomingData.ReadInteger()

        If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
            .Body = BodyData(0)
            .iBody = 0
        Else
            .Body = BodyData(tempint)
            .iBody = tempint

        End If

        headIndex = incomingData.ReadInteger()

        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .iHead = 0
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex

        End If

        .muerto = headIndex = CASPER_HEAD Or headIndex = CASPER_HEAD_CRIMI Or .iBody = FRAGATA_FANTASMAL

        If .muerto Then .Alpha = 80 Else .Alpha = 255
        .Heading = incomingData.ReadByte()

        tempint = incomingData.ReadInteger()

        If tempint <> 0 Then .Arma = WeaponAnimData(tempint)

        tempint = incomingData.ReadInteger()

        If tempint <> 0 Then .Escudo = ShieldAnimData(tempint)

        tempint = incomingData.ReadInteger()

        If tempint <> 0 Then .Casco = CascoAnimData(tempint)

        tempint = incomingData.ReadInteger()

        If tempint <> 0 Then    'Si no mando nada le dejo el que tenia por si solo giro el usuario
            Call SetCharacterFx(CharIndex, tempint, incomingData.ReadInteger())
        Else
            incomingData.ReadInteger

        End If
        .alaIndex = incomingData.ReadInteger()
        If .alaIndex <> 0 Then
            .Alas = alaArray(.alaIndex)
        End If
    End With

    'Call RefreshAllChars
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 5 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Integer

    Dim Y As Integer
    
    X = incomingData.ReadInteger()
    Y = incomingData.ReadInteger()
    
    MapData(X, Y).ObjGrh.GrhIndex = incomingData.ReadInteger()
    
    Call InitGrh(MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex)
    
    If MapData(X, Y).ObjGrh.GrhIndex = GrhFogata Then
        MapData(X, Y).Graphic(3).GrhIndex = -1
        Call Light_Create(X, Y, 255, 220, 220, 6, 0)

    End If

End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Integer

    Dim Y As Integer
    
    X = incomingData.ReadInteger()
    Y = incomingData.ReadInteger()
    
    If MapData(X, Y).ObjGrh.GrhIndex = GrhFogata Then
        MapData(X, Y).Graphic(3).GrhIndex = 0
        Call Light_Destroy_ToMap(X, Y)

    End If
    
    MapData(X, Y).ObjGrh.GrhIndex = 0

End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 4 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Integer

    Dim Y As Integer
    
    X = incomingData.ReadInteger()
    Y = incomingData.ReadInteger()
    
    If MapData(X, Y).Map <> UserMap Then
        If UserMap = 1 Then
            Call CargarTile(CLng(X), CLng(Y), DataMap1)
        Else
            Call CargarTile(CLng(X), CLng(Y), DataMap2)

        End If

    End If
    
    If incomingData.ReadBoolean() Then
        MapData(X, Y).Blocked = 1
    Else
        MapData(X, Y).Blocked = 0

    End If

End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified by: Rapsodius
    'Added support for 3D Sounds.
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim Wave As Byte

    Dim srcX As Integer

    Dim srcY As Integer
    
    Wave = incomingData.ReadByte()
    srcX = incomingData.ReadInteger()
    srcY = incomingData.ReadInteger()
        
    Call Audio.PlayWave(Wave, srcX, srcY)

End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildAdm
        'Clear guild's list
        .guildslist.Clear
        
        GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Dim I As Long

        For I = 0 To UBound(GuildNames())
            Call .guildslist.AddItem(GuildNames(I))
        Next I
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(buffer)
        
        .Show vbModeless, frmMain

    End With
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X    As Integer

    Dim Y    As Integer

    Dim Head As Byte
    
    X = incomingData.ReadInteger()
    Y = incomingData.ReadInteger()
    Head = incomingData.ReadByte()
        
    Call CambioDeArea(X, Y, Head)

End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    pausa = Not pausa

End Sub

''
' Handles the RainToggle message.

Private Sub HandleRainToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

    Dim Tormenta As Byte

    Tormenta = incomingData.ReadByte
    
    If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 7 Or MapData(UserPos.X, UserPos.Y).Trigger = 4)

    If bRain Then
        If Zonas(ZonaActual).Terreno <> eTerreno.Dungeon Then
            'Stop playing the rain sound
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0

            If bTecho Then
                Call Audio.PlayWave(SND_LLUVIAINEND, 0, 0, LoopStyle.Disabled)
            Else
                Call Audio.PlayWave(SND_LLUVIAOUTEND, 0, 0, LoopStyle.Disabled)

            End If

            frmMain.IsPlaying = 0

        End If

    End If
    
    bRain = Not bRain

    If bRain = True Then
        If Tormenta = 1 Then frmMain.tRelampago.Enabled = True ' 50% de Tormenta Electrica
    Else
        frmMain.tRelampago.Enabled = False

    End If

End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 7 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer

    Dim fX        As Integer

    Dim Loops     As Integer
    
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call SetCharacterFx(CharIndex, fX, Loops)

End Sub

''
' Handles the CreateEfecto message.

Private Sub HandleCreateEfecto()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 18 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim tUserCharIndex As Integer

    Dim CharIndex      As Integer

    Dim Efecto         As Byte

    Dim Wav            As Byte

    Dim fX             As Integer

    Dim Loops          As Integer

    Dim X              As Integer

    Dim Y              As Integer

    Dim Xd             As Integer

    Dim Yd             As Integer
    
    tUserCharIndex = incomingData.ReadInteger()
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    Wav = incomingData.ReadByte()
    Efecto = incomingData.ReadByte()
    X = incomingData.ReadInteger()
    Y = incomingData.ReadInteger()
    Xd = incomingData.ReadInteger()
    Yd = incomingData.ReadInteger()
    
    Dim mArroja As New clsArroja

    Call mArroja.Init(tUserCharIndex, CharIndex, fX, Loops, Wav, Efecto, X, Y, Xd, Yd)
    Arrojas.Add mArroja

End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 26 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxHP = incomingData.ReadInteger()
    UserMinHP = incomingData.ReadInteger()
    UserMaxMAN = incomingData.ReadInteger()
    UserMinMAN = incomingData.ReadInteger()
    UserMaxSTA = incomingData.ReadInteger()
    UserMinSTA = incomingData.ReadInteger()
    UserGLD = incomingData.ReadLong()
    UserLvl = incomingData.ReadByte()
    UserPasarNivel = incomingData.ReadLong()
    UserExp = incomingData.ReadLong()
    
    frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
    
    'frmMain.exp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
    '    frmMain.exp.Caption = Round((UserExp / UserPasarNivel) * 100, 2) & "%"
    '    frmMain.Experiencia.Caption = Round((UserExp / UserPasarNivel) * 100, 2) & "%"
    '    frmMain.exp.ToolTipText = "Exp: " & UserExp & "/" & UserPasarNivel
 

    #If RenderFull = 0 Then
       
        'Form1.BarraCir.Value = UserExp
'        frmMain.bar_salud(0).max = UserMaxHP
'        frmMain.bar_salud(0).value = UserMinHP
'        frmMain.bar_salud(1).max = UserMaxHP
'        frmMain.bar_salud(1).value = UserMinHP
'
'        frmMain.bar_salud(0).TextAfterCaption = " / " & UserMaxHP
'        frmMain.bar_salud(1).TextAfterCaption = " / " & UserMaxHP

        'frmMain.HpshpV.Value = (((UserMinHP / 100) / (UserMaxHP / 100)) * 100)
        'Helios Barras
        If UserMaxMAN > 0 Then
'            frmMain.Bar_Mana(0).max = UserMaxMAN
'            frmMain.Bar_Mana(0).value = UserMinMAN
'            frmMain.Bar_Mana(0).TextAfterCaption = " / " & UserMaxMAN
'            frmMain.Bar_Mana(1).max = UserMaxMAN
'            frmMain.Bar_Mana(1).value = UserMinMAN
'            frmMain.Bar_Mana(1).TextAfterCaption = " / " & UserMaxMAN
        Else
'            frmMain.Bar_Mana(0).max = 1
'            frmMain.Bar_Mana(0).value = 0
'            frmMain.Bar_Mana(0).TextAfterCaption = " / 0"
'            frmMain.Bar_Mana(1).max = 1
'            frmMain.Bar_Mana(1).value = 0
'            frmMain.Bar_Mana(1).TextAfterCaption = " / 0"

        End If

    #Else
        
        frmMain.Experiencia.value = UserExp
'        frmMain.bar_salud.max = UserMaxHP
'        frmMain.bar_salud.value = UserMinHP
    
'        frmMain.bar_salud.TextAfterCaption = " / " & UserMaxHP

        'frmMain.HpshpV.Value = (((UserMinHP / 100) / (UserMaxHP / 100)) * 100)
    
        If UserMaxMAN > 0 Then
'            frmMain.Bar_Mana.max = UserMaxMAN
'            frmMain.Bar_Mana.value = UserMinMAN
'            frmMain.Bar_Mana.TextAfterCaption = " / " & UserMaxMAN
        Else
'            frmMain.Bar_Mana.max = 1
'            frmMain.Bar_Mana.value = 0
'            frmMain.Bar_Mana.TextAfterCaption = " / 0"

        End If
    
    #End If
    
    '    If UserMaxMAN > 0 Then
    '        frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 100)
    '    Else
    '        frmMain.MANShp.Width = 0
    '    End If
'    frmMain.bar_sta.max = UserMaxSTA
'    frmMain.bar_sta.value = UserMinSTA
'    frmMain.bar_sta.TextAfterCaption = " / " & UserMaxSTA
    
    ' frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 100)
    
    '    frmMain.lblEnergiaN.Caption = UserMinSTA & "/" & UserMaxSTA
    '    frmMain.lblEnergia.Caption = frmMain.lblEnergiaN.Caption
    '    frmMain.lblVidaN.Caption = UserMinHP & "/" & UserMaxHP
    '    frmMain.lblVida.Caption = frmMain.lblVidaN.Caption
    '    frmMain.lblManaN.Caption = UserMinMAN & "/" & UserMaxMAN
    '    frmMain.lblMana.Caption = frmMain.lblManaN.Caption
    
    If UserGLD <> 0 Then
        frmMain.GldLbl.Caption = Format$(UserGLD, "##,##")
    Else
        frmMain.GldLbl.Caption = UserGLD

    End If
    
    Dim R As Integer, G As Integer, b As Integer

    R = 255
    G = 255 - (UserLvl / 1.9038)
    b = 255 - (UserLvl / 0.3882)

    If (R < 0) Then R = 0: If (R > 255) Then R = 255
    If (G < 0) Then G = 0: If (G > 255) Then G = 255
    If (b < 0) Then b = 0: If (b > 255) Then b = 255
    frmMain.LvlLbl.ForeColor = RGB(R, G, b)
    
    '    frmMain.LvlLbl.Caption = UserLvl
    
    If UserMinHP = 0 Then
        UserEstado = 1
        
        'Sacamos un screenshot si está activado el FragShooter:
        If mOpciones.ScreenShooterAlMorir = True Then
            ScreenShooterCapturePending = True

        End If
        
        If frmMain.macrotrabajo Then frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0

    End If

End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 2 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UsingSkill = incomingData.ReadByte()

    frmMain.MousePointer = 2
    
    Select Case UsingSkill

        Case Magia
            Call AddtoRichPicture(MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)

            'frmMain.tMouse.Enabled = True
        Case Pesca
            'Call AddtoRichPicture(MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)

        Case Robar
            Call AddtoRichPicture(MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)

        Case Talar
            'Call AddtoRichPicture(MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)

        Case Mineria
           'Call AddtoRichPicture(MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)

        Case FundirMetal
            'Call AddtoRichPicture(MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)

        Case Proyectiles
            frmMain.MousePointer = 99
            Call AddtoRichPicture(MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)

            'frmMain.tMouse.Enabled = True
            If UserNavegando = True And UsingSecondSkill = 1 Then
                Call SetCursor(proyectil)
            Else
                Call SetCursor(ProyectilPequena)

            End If

            Exit Sub

    End Select
    
End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowPartyForm()

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim members() As String

    Dim I         As Long
    
    EsPartyLeader = CBool(buffer.ReadByte())
       
    members = Split(buffer.ReadASCIIString(), SEPARATOR)

    For I = 0 To UBound(members())
        Call frmParty.lstMembers.AddItem(members(I))
    Next I
    
    frmParty.lblTotalExp.Caption = buffer.ReadLong
    frmParty.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 23 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot          As Byte

    Dim OBJIndex      As Integer

    Dim Name          As String

    Dim Amount        As Integer

    Dim Equipped      As Boolean

    Dim GrhIndex      As Integer

    Dim ObjType       As Byte

    Dim MaxHit        As Integer

    Dim MinHit        As Integer

    Dim MinDef        As Integer

    Dim MaxDef        As Integer

    Dim value         As Single

    Dim PuedeUsarItem As Byte
    
    slot = buffer.ReadByte()
    OBJIndex = buffer.ReadInteger()
    Name = buffer.ReadASCIIString()
    Amount = buffer.ReadInteger()
    Equipped = buffer.ReadBoolean()
    GrhIndex = buffer.ReadInteger()
    ObjType = buffer.ReadByte()
    MaxHit = buffer.ReadInteger()
    MinHit = buffer.ReadInteger()
    MinDef = buffer.ReadInteger()
    MaxDef = buffer.ReadInteger()
    value = buffer.ReadSingle()
    PuedeUsarItem = buffer.ReadByte()
    
    If Equipped Then

        Select Case ObjType

            Case eObjType.otWeapon
                frmMain.lblArma = MinHit & "/" & MaxHit

                'UserWeaponEqpSlot = slot
            Case eObjType.otArmadura
                frmMain.lblArmadura = MinDef & "/" & MaxDef

                'UserArmourEqpSlot = slot
            Case eObjType.otescudo
                frmMain.lblEscudo = MinDef & "/" & MaxDef

                'UserHelmEqpSlot = slot
            Case eObjType.otcasco
                frmMain.lblCasco = MinDef & "/" & MaxDef

                'UserShieldEqpSlot = slot
        End Select

    Else

        If ObjType > 0 Then

            Select Case ObjType

                Case eObjType.otWeapon
                    frmMain.lblArma = "N/A"

                    'UserWeaponEqpSlot = slot
                Case eObjType.otArmadura
                    frmMain.lblArmadura = "N/A"

                    'UserArmourEqpSlot = slot
                Case eObjType.otescudo
                    frmMain.lblEscudo = "N/A"

                    'UserHelmEqpSlot = slot
                Case eObjType.otcasco
                    frmMain.lblCasco = "N/A"

                    'UserShieldEqpSlot = slot
            End Select

        End If

    End If
    
    Call Inventario.SetItem(slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, MaxDef, value, Name, PuedeUsarItem)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 22 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte

    slot = buffer.ReadByte()
    
    With UserBancoInventory(slot)
        .OBJIndex = buffer.ReadInteger()
        .Name = buffer.ReadASCIIString()
        .Amount = buffer.ReadInteger()
        .GrhIndex = buffer.ReadInteger()
        .ObjType = buffer.ReadByte()
        .MaxHit = buffer.ReadInteger()
        .MinHit = buffer.ReadInteger()
        .MaxDef = buffer.ReadInteger()
        .MinDef = .MaxDef
        .Valor = buffer.ReadLong()
        .PuedeUsarItem = buffer.ReadByte()
        
        If Comerciando Then
            Call InvBanco(0).SetItem(slot, .OBJIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .PuedeUsarItem)

        End If

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo ErrHandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte
    Dim spell_Name  As String
    Dim spell_Gra As Integer
    slot = buffer.ReadByte()
    
    UserHechizos(slot) = buffer.ReadInteger()
    
    spell_Name = buffer.ReadASCIIString()
   
    If slot <= frmMain.hlst.ListCount Then
        frmMain.hlst.List(slot - 1) = spell_Name
    Else
        Call frmMain.hlst.AddItem(spell_Name)
    End If
    
    If spell_Name <> "(None)" Then
      spell_Gra = buffer.ReadInteger()
        If spell_Gra = 0 Then
        Call invSpells.SetItem(slot, UserHechizos(slot), 0, 0, 609, 0, 0, 0, 0, 0, 0, spell_Name, 1)
        Else
         
        Call invSpells.SetItem(slot, UserHechizos(slot), 0, 0, spell_Gra, 0, 0, 0, 0, 0, 0, spell_Name, 1)
        End If
    Else
        Call invSpells.SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 1)
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:
    Dim Error As Long
    Error = err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then _
        err.Raise Error

End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 6 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim I    As Long

    Dim TIPO As Byte
    
    TIPO = incomingData.ReadByte()
    
    If TIPO = 1 Then
        UserAtributos(1) = incomingData.ReadByte()
        UserAtributos(2) = incomingData.ReadByte()
        FuerzaBk = incomingData.ReadByte()
        AgilidadBk = incomingData.ReadByte()
        
        frmMain.lblFuerza.Caption = UserAtributos(1)
        frmMain.lblAgilidad.Caption = UserAtributos(2)
    Else

        For I = 1 To NUMATRIBUTES
            UserAtributos(I) = incomingData.ReadByte()
        Next I

        LlegaronAtrib = True

    End If

End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer

    Dim I     As Long

    Dim J     As Long

    Dim k     As Long
    
    count = buffer.ReadInteger()
    
    ReDim ArmasHerrero(count) As tItemsConstruibles
    ReDim HerreroMejorar(0) As tItemsConstruibles
    
    For I = 1 To count

        With ArmasHerrero(I)
            .Name = buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = buffer.ReadInteger()
            .LinH = buffer.ReadInteger()        'The iron needed
            .LinP = buffer.ReadInteger()        'The silver needed
            .LinO = buffer.ReadInteger()        'The gold needed
            .OBJIndex = buffer.ReadInteger()
            .Upgrade = buffer.ReadInteger()

        End With

    Next I
    
    With frmHerrero
        ' Inicializo los inventarios
        Call InvLingosHerreria(1).Initialize(.picLingotes0, 0, 0, 3, False, , , , , , False)
        Call InvLingosHerreria(2).Initialize(.picLingotes1, 0, 0, 3, False, , , , , , False)
        Call InvLingosHerreria(3).Initialize(.picLingotes2, 0, 0, 3, False, , , , , , False)
        Call InvLingosHerreria(4).Initialize(.picLingotes3, 0, 0, 3, False, , , , , , False)
        
        Call .HideExtraControls(count)
        Call .RenderList(1, True)

    End With
    
    For I = 1 To count

        With ArmasHerrero(I)

            If .Upgrade Then

                For k = 1 To count

                    If .Upgrade = ArmasHerrero(k).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).Name = .Name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).OBJIndex = .OBJIndex
                        HerreroMejorar(J).UpgradeName = ArmasHerrero(k).Name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmasHerrero(k).GrhIndex
                        HerreroMejorar(J).LinH = ArmasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For

                    End If

                Next k

            End If

        End With

    Next I
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    Exit Sub
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer

    Dim I     As Long

    Dim J     As Long

    Dim k     As Long
    
    count = buffer.ReadInteger()
    
    ReDim ArmadurasHerrero(count) As tItemsConstruibles
    
    For I = 1 To count

        With ArmadurasHerrero(I)
            .Name = buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = buffer.ReadInteger()
            .LinH = buffer.ReadInteger()        'The iron needed
            .LinP = buffer.ReadInteger()        'The silver needed
            .LinO = buffer.ReadInteger()        'The gold needed
            .OBJIndex = buffer.ReadInteger()
            .Upgrade = buffer.ReadInteger()

        End With

    Next I
    
    J = UBound(HerreroMejorar)
    
    For I = 1 To count

        With ArmadurasHerrero(I)

            If .Upgrade Then

                For k = 1 To count

                    If .Upgrade = ArmadurasHerrero(k).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).Name = .Name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).OBJIndex = .OBJIndex
                        HerreroMejorar(J).UpgradeName = ArmadurasHerrero(k).Name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmadurasHerrero(k).GrhIndex
                        HerreroMejorar(J).LinH = ArmadurasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmadurasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmadurasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For

                    End If

                Next k

            End If

        End With

    Next I
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    Exit Sub
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer

    Dim I     As Long

    Dim J     As Long

    Dim k     As Long
    
    count = buffer.ReadInteger()
    
    ReDim ObjCarpintero(count) As tItemsConstruibles
    ReDim CarpinteroMejorar(0) As tItemsConstruibles
    
    For I = 1 To count

        With ObjCarpintero(I)
            .Name = buffer.ReadASCIIString()        'Get the object's name
            .GrhIndex = buffer.ReadInteger()
            .Madera = buffer.ReadInteger()          'The wood needed
            .MaderaElfica = buffer.ReadInteger()    'The elfic wood needed
            .OBJIndex = buffer.ReadInteger()
            .Upgrade = buffer.ReadInteger()

        End With

    Next I
    
    With frmCarp
        ' Inicializo los inventarios
        Call InvMaderasCarpinteria(1).Initialize(.picMaderas0, 0, 0, 2, False, , , , , , False)
        Call InvMaderasCarpinteria(2).Initialize(.picMaderas1, 0, 0, 2, False, , , , , , False)
        Call InvMaderasCarpinteria(3).Initialize(.picMaderas2, 0, 0, 2, False, , , , , , False)
        Call InvMaderasCarpinteria(4).Initialize(.picMaderas3, 0, 0, 2, False, , , , , , False)
        
        Call .HideExtraControls(count)
        Call .RenderList(1)

    End With
    
    For I = 1 To count

        With ObjCarpintero(I)

            If .Upgrade Then

                For k = 1 To count

                    If .Upgrade = ObjCarpintero(k).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve CarpinteroMejorar(J) As tItemsConstruibles
                        
                        CarpinteroMejorar(J).Name = .Name
                        CarpinteroMejorar(J).GrhIndex = .GrhIndex
                        CarpinteroMejorar(J).OBJIndex = .OBJIndex
                        CarpinteroMejorar(J).UpgradeName = ObjCarpintero(k).Name
                        CarpinteroMejorar(J).UpgradeGrhIndex = ObjCarpintero(k).GrhIndex
                        CarpinteroMejorar(J).Madera = ObjCarpintero(k).Madera - .Madera * 0.85
                        CarpinteroMejorar(J).MaderaElfica = ObjCarpintero(k).MaderaElfica - .MaderaElfica * 0.85
                        
                        Exit For

                    End If

                Next k

            End If

        End With

    Next I
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar

End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call MessageBox(buffer.ReadASCIIString())
    
    If MostrarEntrar > 0 Then
        CloseSock

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = True
    AlphaCeguera = 255

End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = True

End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 5 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim tmp As String

    tmp = buffer.ReadASCIIString()
    
    Call InitCartel(tmp, buffer.ReadInteger())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 22 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte

    slot = buffer.ReadByte()
    
    With NPCInventory(slot)
        .Name = buffer.ReadASCIIString()
        .Amount = buffer.ReadInteger()
        .Valor = buffer.ReadSingle()
        .GrhIndex = buffer.ReadInteger()
        .OBJIndex = buffer.ReadInteger()
        .ObjType = buffer.ReadByte()
        .MaxHit = buffer.ReadInteger()
        .MinHit = buffer.ReadInteger()
        .MaxDef = buffer.ReadInteger()
        .PuedeUsarItem = buffer.ReadByte()
        
        '.MinDef = Buffer.ReadInteger
    End With
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 5 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxAGU = incomingData.ReadByte()
    UserMinAGU = incomingData.ReadByte()
    UserMaxHAM = incomingData.ReadByte()
    UserMinHAM = incomingData.ReadByte()
    
    '    frmMain.imgSed.Height = 33 - (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 33)
    '    frmMain.imgHambre.Height = 33 - (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 33)
    
'    frmMain.Bar_Agua.max = UserMaxAGU
'    frmMain.Bar_Agua.value = UserMinAGU
'    frmMain.Bar_Agua.TextAfterCaption = " / " & UserMaxAGU
'    frmMain.bar_comida.max = UserMaxHAM
'    frmMain.bar_comida.value = UserMinHAM
'    frmMain.bar_comida.TextAfterCaption = " / " & UserMaxHAM
    
    '    frmMain.lblHambreN.Caption = UserMinHAM
    '    frmMain.lblHambre.Caption = frmMain.lblHambreN.Caption
    '    frmMain.lblSedN.Caption = UserMinAGU
    '    frmMain.lblSed.Caption = frmMain.lblSedN.Caption
End Sub

''
' Handles the Fame message.

Private Sub HandleFame()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 29 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserReputacion
        .AsesinoRep = incomingData.ReadLong()
        .BandidoRep = incomingData.ReadLong()
        .BurguesRep = incomingData.ReadLong()
        .LadronesRep = incomingData.ReadLong()
        .NobleRep = incomingData.ReadLong()
        .PlebeRep = incomingData.ReadLong()
        .Promedio = incomingData.ReadLong()

    End With
    
    LlegoFama = True

End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 20 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .UsuariosMatados = incomingData.ReadLong()
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .PenaCarcel = incomingData.ReadLong()

    End With

End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 4 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).invisible = incomingData.ReadBoolean()

    If charlist(CharIndex).invisible Then
        charlist(CharIndex).Alpha = 0
        charlist(CharIndex).iTick = 0
        charlist(CharIndex).ContadorInvi = INTERVALO_INVI
        charlist(CharIndex).oculto = False
    Else
        charlist(CharIndex).Alpha = 255
        charlist(CharIndex).iTick = 0
        charlist(CharIndex).ContadorInvi = 0
        charlist(CharIndex).oculto = False

    End If
    
    #If SeguridadAlkon Then

        If charlist(CharIndex).invisible Then
            Call MI(CualMI).SetInvisible(CharIndex)
        Else
            Call MI(CualMI).ResetInvisible(CharIndex)

        End If

    #End If

End Sub

Private Sub HandleSetOculto()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 4 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).invisible = incomingData.ReadBoolean()
    
    If charlist(CharIndex).invisible Then
        charlist(CharIndex).Alpha = 0
        charlist(CharIndex).oculto = True
    Else
        charlist(CharIndex).oculto = False

    End If
    
    #If SeguridadAlkon Then

        If charlist(CharIndex).invisible Then
            Call MI(CualMI).SetOculto(CharIndex)
        Else
            Call MI(CualMI).ResetOculto(CharIndex)

        End If

    #End If

End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMeditar = Not UserMeditar

End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = False
    AlphaCeguera = 0

End Sub

Private Sub HandleAtaca()

    If incomingData.Length < 4 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
     
    Dim X    As Integer

    Dim TIPO As Byte

    X = incomingData.ReadInteger()
    TIPO = incomingData.ReadByte()
    
    If TIPO = 1 Then
        If charlist(X).Arma.WeaponWalk(charlist(X).Heading).GrhIndex > 0 Then
            charlist(X).Arma.WeaponWalk(charlist(X).Heading).Started = 1
            charlist(X).Arma.WeaponAttack = 1

        End If

    End If

    If charlist(X).Escudo.ShieldWalk(charlist(X).Heading).GrhIndex > 0 Then
        charlist(X).Escudo.ShieldWalk(charlist(X).Heading).Started = 1
        charlist(X).Escudo.ShieldAttack = 1

    End If
   
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = False

End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 1 + NUMSKILLS Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim I As Long
    
    SkillPoints = incomingData.ReadInteger()
    SPLibres = SkillPoints

    For I = 1 To NUMSKILLS
        UserSkills(I) = incomingData.ReadByte()
        UserSkillsMod(I) = UserSkills(I)
    Next I

    LlegaronSkills = True

End Sub

Private Sub HandleRetosAbre()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 2 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Crea As Boolean

    Dim Vs   As Byte

    Crea = buffer.ReadBoolean
    
    If Crea Then
        Call frmRetos.Iniciar(1, False, 0, UserName, False, "", False, "", False, "", False, False)
        
    Else
        Vs = buffer.ReadByte

        If Vs = 1 Then
            Call frmRetos.Iniciar(Vs, buffer.ReadBoolean, buffer.ReadLong, buffer.ReadASCIIString, buffer.ReadBoolean, buffer.ReadASCIIString, buffer.ReadBoolean, "", False, "", False, buffer.ReadBoolean)
        Else
            Call frmRetos.Iniciar(Vs, buffer.ReadBoolean, buffer.ReadLong, buffer.ReadASCIIString, buffer.ReadBoolean, buffer.ReadASCIIString, buffer.ReadBoolean, buffer.ReadASCIIString, buffer.ReadBoolean, buffer.ReadASCIIString, buffer.ReadBoolean, buffer.ReadBoolean)

        End If

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    frmRetos.Show , frmMain
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

Private Sub HandleRetosRespuesta()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 2 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer  As New clsByteQueue

    Dim msg     As Byte

    Dim Mensaje As String

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    msg = buffer.ReadByte
    
    Select Case msg

        Case 1
            Unload frmRetos

        Case 2
            Mensaje = "No tienes suficiente oro para realizar el reto."

        Case 3
            Mensaje = "Todos los personajes deben estar en zonas seguras para poder ingresar al reto."

        Case 4
            Mensaje = "Todas las salas de retos están ocupadas, espere un momento y vuelva a intentarlo."

        Case 5
            Mensaje = "Alguno de los personajes ya tiene una invitación pendiente para un reto."

        Case 6
            Mensaje = "La apuesta máxima es de 500.000 monedas"

        Case 7
            Mensaje = "Hay demasiados solicitudes de retos creadas, espere un momento y vuelva a intentarlo."

        Case 8
            Mensaje = "Los jugadores deben ser distintos."

        Case 9
            Mensaje = "El mínimo por apuesta es de 5.000 monedas"

        Case 10
            TiempoRetos = (GetTickCount() And &H7FFFFFFF)

        Case 11
            Mensaje = "No podés aceptar un reto estando muerto."

        Case 12
            Mensaje = "¡Estás encarcelado! No podés aceptar el reto."

        Case 13
            Mensaje = "No podés aceptar un reto mientras estás comerciando."

        Case 14
            Mensaje = "No podés aceptar un reto mientras navegás."

        Case 15
            Mensaje = "Alguno de los personajes no cumple con las condiciones para ingresar al reto, verifique que todos tengan el oro suficiente, estén en zonas seguras, se encuentren vivos y listos."

        Case 16
            Mensaje = "El nivel de los participantes debe ser de 15 o más."

    End Select
    
    If Mensaje <> "" Then
        frmMensaje.msg.Caption = Mensaje
        frmMensaje.Show

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim creatures() As String

    Dim I           As Long
    
    creatures = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For I = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(I))
    Next I

    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 7 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim guildList() As String

    Dim I           As Long

    Dim Stemp       As String
    
    'Get news' string
    frmGuildNews.news = buffer.ReadASCIIString()
    
    'Get Enemy guilds list
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For I = 0 To UBound(guildList)
        Stemp = frmGuildNews.txtClanesGuerra.Text
        frmGuildNews.txtClanesGuerra.Text = Stemp & guildList(I) & vbCrLf
    Next I
    
    'Get Allied guilds list
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For I = 0 To UBound(guildList)
        Stemp = frmGuildNews.txtClanesAliados.Text
        frmGuildNews.txtClanesAliados.Text = Stemp & guildList(I) & vbCrLf
    Next I
    
    If mOpciones.GuildNews = True Then
        frmGuildNews.Show vbModeless, frmMain

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim guildList() As String

    Dim I           As Long
    
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For I = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(I))
    Next I
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim guildList() As String

    Dim I           As Long
    
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For I = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(I))
    Next I
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 35 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmCharInfo

        If .frmType = CharInfoFrmType.frmMembers Then
            .imgRechazar.Visible = False
            .imgAceptar.Visible = False
            .imgEchar.Visible = True
            .imgPeticion.Visible = False
        Else
            .imgRechazar.Visible = True
            .imgAceptar.Visible = True
            .imgEchar.Visible = False
            .imgPeticion.Visible = True

        End If
        
        .nombre.Caption = buffer.ReadASCIIString()
        .Raza.Caption = ListaRazas(buffer.ReadByte())
        .Clase.Caption = ListaClases(buffer.ReadByte())
        
        If buffer.ReadByte() = 1 Then
            .Genero.Caption = "Hombre"
        Else
            .Genero.Caption = "Mujer"

        End If
        
        .Nivel.Caption = buffer.ReadByte()
        .Oro.Caption = buffer.ReadLong()
        .Banco.Caption = buffer.ReadLong()
        
        Dim reputation As Long

        reputation = buffer.ReadLong()
        
        .reputacion.Caption = reputation
        
        .txtPeticiones.Text = buffer.ReadASCIIString()
        .guildactual.Caption = buffer.ReadASCIIString()
        .txtMiembro.Text = buffer.ReadASCIIString()
        
        Dim armada As Boolean

        Dim caos   As Boolean
        
        armada = buffer.ReadBoolean()
        caos = buffer.ReadBoolean()
        
        If armada Then
            .ejercito.Caption = "Armada Real"
        ElseIf caos Then
            .ejercito.Caption = "Legión Oscura"

        End If
        
        .Ciudadanos.Caption = CStr(buffer.ReadLong())
        .criminales.Caption = CStr(buffer.ReadLong())
        
        If reputation > 0 Then
            .status.Caption = " Ciudadano"
            .status.ForeColor = vbBlue
        Else
            .status.Caption = " Criminal"
            .status.ForeColor = vbRed

        End If
        
        Call .Show(vbModeless, frmMain)

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 9 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim I      As Long

    Dim List() As String
    
    With frmGuildLeader
        'Get list of existing guilds
        GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For I = 0 To UBound(GuildNames())
            Call .guildslist.AddItem(GuildNames(I))
        Next I
        
        'Get list of guild's members
        GuildMembers = Split(buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .members.Clear
        
        For I = 0 To UBound(GuildMembers())
            Call .members.AddItem(GuildMembers(I))
        Next I
        
        .txtguildnews = buffer.ReadASCIIString()
        
        'Get list of join requests
        List = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear
        
        For I = 0 To UBound(List())
            Call .solicitudes.AddItem(List(I))
        Next I
        
        .Show , frmMain

    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

Private Sub HandleGuildMemberInfo()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildMember
        'Clear guild's list
        .lstClanes.Clear
        
        GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Dim I As Long

        For I = 0 To UBound(GuildNames())
            Call .lstClanes.AddItem(GuildNames(I))
        Next I
        
        'Get list of guild's members
        GuildMembers = Split(buffer.ReadASCIIString(), SEPARATOR)
        .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .lstMiembros.Clear
        
        For I = 0 To UBound(GuildMembers())
            Call .lstMiembros.AddItem(GuildMembers(I))
        Next I
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(buffer)
        
        .Show vbModeless, frmMain

    End With
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 26 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildBrief
        .imgDeclararGuerra.Visible = .EsLeader
        .imgOfrecerAlianza.Visible = .EsLeader
        .imgOfrecerPaz.Visible = .EsLeader
        
        .nombre.Caption = buffer.ReadASCIIString()
        .fundador.Caption = buffer.ReadASCIIString()
        .creacion.Caption = buffer.ReadASCIIString()
        .lider.Caption = buffer.ReadASCIIString()
        .web.Caption = buffer.ReadASCIIString()
        .Miembros.Caption = buffer.ReadInteger()
        
        If buffer.ReadBoolean() Then
            .eleccion.Caption = "ABIERTA"
        Else
            .eleccion.Caption = "CERRADA"

        End If
        
        .lblAlineacion.Caption = buffer.ReadASCIIString()
        .Enemigos.Caption = buffer.ReadInteger()
        .Aliados.Caption = buffer.ReadInteger()
        .antifaccion.Caption = buffer.ReadASCIIString()
        
        Dim codexStr() As String

        Dim I          As Long
        
        codexStr = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        For I = 0 To 7
            .Codex(I).Caption = codexStr(I)
        Next I
        
        .Desc.Text = buffer.ReadASCIIString()

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildFoundation.Show , frmMain

End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    UserParalizado = Not UserParalizado
   
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmComerciar.Visible Then

        Dim I As Long
        
        'Update user inventory
        For I = 1 To MAX_INVENTORY_SLOTS

            ' Agrego o quito un item en su totalidad
            If Inventario.OBJIndex(I) <> InvComUsu.OBJIndex(I) Then

                With Inventario
                    Call InvComUsu.SetItem(I, .OBJIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), .PuedeUsarItem(I))

                End With

                ' Vendio o compro cierta cantidad de un item que ya tenia
            ElseIf Inventario.Amount(I) <> InvComUsu.Amount(I) Then
                Call InvComUsu.ChangeSlotItemAmount(I, Inventario.Amount(I))

            End If

        Next I
        
        ' Fill Npc inventory
        For I = 1 To 20

            ' Compraron la totalidad de un item, o vendieron un item que el npc no tenia
            If NPCInventory(I).OBJIndex <> InvComNpc.OBJIndex(I) Then

                With NPCInventory(I)
                    Call InvComNpc.SetItem(I, .OBJIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .PuedeUsarItem)

                End With

                ' Compraron o vendieron cierta cantidad (no su totalidad)
            ElseIf NPCInventory(I).Amount <> InvComNpc.Amount(I) Then
                Call InvComNpc.ChangeSlotItemAmount(I, NPCInventory(I).Amount)

            End If

        Next I

    End If

End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim I As Long
    
    If frmBancoObj.Visible Then
        
        For I = 1 To Inventario.MaxObjs

            With Inventario
                Call InvBanco(1).SetItem(I, .OBJIndex(I), .Amount(I), .Equipped(I), .GrhIndex(I), .ObjType(I), .MaxHit(I), .MinHit(I), .MaxDef(I), .MinDef(I), .Valor(I), .ItemName(I), 0)

            End With

        Next I
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmBancoObj.LasActionBuy Then
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
        Else

            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
        End If
        
        frmBancoObj.NoPuedeMover = False

    End If
       
End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 22 Then
        NotEnoughData = True
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    Dim OfferSlot As Byte
    
    'Remove packet ID
    Call buffer.ReadByte
    
    OfferSlot = buffer.ReadByte
    
    With buffer

        If OfferSlot = GOLD_OFFER_SLOT Then
            Call InvOroComUsu(2).SetItem(1, .ReadInteger(), .ReadLong(), 0, .ReadInteger(), .ReadByte(), .ReadInteger(), .ReadInteger(), .ReadInteger(), 0, .ReadLong(), .ReadASCIIString(), 0)
        Else
            Call InvOfferComUsu(1).SetItem(OfferSlot, .ReadInteger(), .ReadLong(), 0, .ReadInteger(), .ReadByte(), .ReadInteger(), .ReadInteger(), .ReadInteger(), 0, .ReadLong(), .ReadASCIIString(), 0)

        End If

    End With
    
    Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " ha modificado su oferta.", FontTypeNames.FONTTYPE_VENENO)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim creatureList() As String

    Dim I              As Long
    
    creatureList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For I = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(I))
    Next I

    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim sosList() As String

    Dim I         As Long
    
    sosList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For I = 0 To UBound(sosList())
        Call frmMSG.List1.AddItem(sosList(I))
    Next I
    
    frmMSG.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmPanelGm.Show vbModal

End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim userList() As String

    Dim I          As Long
    
    userList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear

        For I = 0 To UBound(userList())
            Call frmPanelGm.cboListaUsus.AddItem(userList(I))
        Next I

        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Call incomingData.ReadByte
    
    Call AddtoRichPicture("El ping es " & ((GetTickCount() And &H7FFFFFFF) - pingTime) & " ms.", 255, 0, 0, True, False, False)
    
    pingTime = 0

End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 6 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim CharIndex As Integer

    Dim Criminal  As Boolean

    Dim userTag   As String
    
    CharIndex = buffer.ReadInteger()
    Criminal = buffer.ReadBoolean()
    userTag = buffer.ReadASCIIString()
    
    'Update char status adn tag!
    With charlist(CharIndex)

        If Criminal Then
            .Criminal = 1
            UserFaccion = 1
        Else
            .Criminal = 0
            UserFaccion = 2

        End If
        
        .nombre = userTag
        
        If CharIndex = UserCharIndex Then 'Si cambió de estado cambiamos el cursor
            Call SetCursor(General)

        End If
        
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

Private Sub HandleUsersOnline()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    'Dim UsersOn As Integer

    UsersOn = buffer.ReadInteger()
    frmMain.lblUsers.Caption = UsersOn
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Private Sub HandleSetEquitando()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 4 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).equitando = incomingData.ReadBoolean()
    
    If UserCharIndex = CharIndex Then
        UserEquitando = charlist(CharIndex).equitando

    End If

End Sub

Private Sub HandleSetCongelado()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 4 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).congelado = incomingData.ReadBoolean()
    
    If UserCharIndex = CharIndex Then
        UserCongelado = charlist(CharIndex).congelado

    End If

End Sub

Private Sub HandleSetChiquito()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 4 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).Chiquito = incomingData.ReadBoolean()
    
    If UserCharIndex = CharIndex Then
        UserChiquito = charlist(CharIndex).Chiquito

    End If

End Sub

Private Sub HandlePalabrasMagicas()

    If incomingData.Length < 5 Then
        err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
 
    Dim SpellWords As String

    Dim CharIndex  As Integer
    
    SpellWords = buffer.ReadASCIIString
    CharIndex = buffer.ReadInteger
    
    If SpellWords <> "" Then
        Call Dialogos.CreateDialog(SpellWords, CharIndex, 0, 222, 222)

    End If
    
    Call incomingData.CopyBuffer(buffer)

End Sub

Private Sub HandleCreateAreaFX()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 9 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X     As Integer

    Dim Y     As Integer

    Dim fX    As Integer

    Dim Loops As Integer
    
    X = incomingData.ReadInteger()
    Y = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call SetAreaFx(X, Y, fX, Loops)

End Sub

Private Sub HandleSetNadando()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 4 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    charlist(CharIndex).nadando = incomingData.ReadBoolean()
    
    If UserCharIndex = CharIndex Then
        UserNadando = charlist(CharIndex).nadando

    End If

End Sub

Public Sub HandleMultiMessage()
 
    Dim MessageType As Byte

    Dim Number      As Integer

    Dim nombre      As String

    'Dim SpellIndex As Integer
    
    Call incomingData.ReadByte
    
    MessageType = incomingData.ReadByte

    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)

        Select Case MessageType 'eMessages

            Case eMessages.LanzaHechizoA
                Number = incomingData.ReadInteger()
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg("Le has quitado " & Number & " puntos de vida a " & nombre, .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.TeLanzanHechizo
                Number = incomingData.ReadInteger()
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg(nombre & " te ha quitado " & Number & " puntos de vida.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasDesequipadoEsudoOponente
                Call ShowConsoleMsg("¡Hás logrado desequipar el escudo de tu oponente!!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.TeHanDesequipadoEscudo
                Call ShowConsoleMsg("¡Tu oponente te ha desequipado el escudo", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasDesarmadoAlOponente
                Call ShowConsoleMsg("¡Hás logrado desarmar a tu oponente!!!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.TeHanDesarmado
                Call ShowConsoleMsg("¡Tú oponente te ha desarmado!!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasDesequipadoCascoOponente
                Call ShowConsoleMsg("¡Hás logrado desequipar el casco de tu oponente!!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.TeHanDesequipadoCasco
                Call ShowConsoleMsg("¡Tu oponente te ha desequipado el casco", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.OponenteNoTienEquipadoItems
                Call ShowConsoleMsg("¡Tu oponente no tiene equipado objetos!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoHasLogradoDesarmarATuOponente
                Call ShowConsoleMsg("¡No hás logrado desarmar a tu oponente!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.RealNoRobaCiudadanos
                Call ShowConsoleMsg("Los miembros del ejército real no tienen permitido robarle a ciudadanos.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.CaosNoRobaCaos
                Call ShowConsoleMsg("No puedes robar a otros miembros de la legión oscura.", .red, .green, .blue, .bold, .italic)
                Exit Sub
                
            Case eMessages.HasApunaladoA
                Number = incomingData.ReadInteger()
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg("Hás apuñalado a " & nombre & " por " & Number, 185, 185, 185, True, .italic)
                Call Audio.PlayWave(SND_APUÑALAR)
                Exit Sub

            Case eMessages.TeHanApunalado
                Number = incomingData.ReadInteger()
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg("Te ha apuñalado " & nombre & " por " & Number, .red, .green, .blue, .bold, .italic)
                AlphaBlood = 255
                Call Audio.PlayWave(SND_APUÑALAR)
                Exit Sub

            Case eMessages.HasApunaladoCriatura
                Number = incomingData.ReadInteger()
                Call ShowConsoleMsg("Hás apuñalado la criatura por " & Number & ".", 185, 185, 185, True, .italic)
                Call Audio.PlayWave(SND_APUÑALAR)
                Exit Sub

            Case eMessages.NoHasApunalado
                Call ShowConsoleMsg("¡No has logrado apuñalar a tu enemigo!", .red, .green, .blue, .bold, .italic)
                Exit Sub
                    
            Case eMessages.HasGolpeadoCriticamente
                Number = incomingData.ReadInteger()
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg("Has golpeado críticamente a " & nombre & " por " & Number & ".", .red, .green, .blue, .bold, .italic)
                Exit Sub
                    
            Case eMessages.TeHanGolpeadoCriticamente
                Number = incomingData.ReadInteger()
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg(nombre & " te ha golpeado críticamente por " & Number & ".", .red, .green, .blue, .bold, .italic)
                Exit Sub
                    
            Case eMessages.HasGolpeadoCriticamenteCriatura
                Number = incomingData.ReadInteger()
                Call ShowConsoleMsg("Hás golpeado críticamente a la criatura por " & Number & ".", .red, .green, .blue, .bold, .italic)
                Exit Sub

        End Select

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)

        Select Case MessageType 'eMessages

            Case eMessages.mUserMuerto
                Call ShowConsoleMsg("¡Estás muerto!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.mUserParalizado
                Call ShowConsoleMsg("¡Estás paralizado!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.mUserChiquito
                Call ShowConsoleMsg("¡No puedes realizar esta acción con tu apariencia actual!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.UserEstaLejos
                Call ShowConsoleMsg("¡Estás muy lejos!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.mUserChiquito
                Call ShowConsoleMsg("¡No puedes realizar esta acción con tu apariencia actual!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.mUserComerciando
                Call ShowConsoleMsg("¡No puedes realizar esta acción comerciando!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.mUserCongelado
                Call ShowConsoleMsg("¡No puedes realizar esta acción congelado!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.mUserSaliendo
                Call ShowConsoleMsg("¡No puedes realizar esta acción saliendo del juego!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.mUserEmbarcado
                Call ShowConsoleMsg("¡No puedes realizar esta acción embarcado!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.mUserEquitando
                Call ShowConsoleMsg("¡No puedes realizar esta acción montado!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.DomarAnimalParaMontarlo
                Call ShowConsoleMsg("¡Debes domesticar primero a la criatura para montarlo!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.MonturaNoTeAceptaComoSuAmo
                Call ShowConsoleMsg("¡La criatura no te acepta como su dueño!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.UserHaMontado
                Call ShowConsoleMsg("¡Hás montado la criatura!", .red, .green, .blue, True, .italic)
                Exit Sub

            Case eMessages.NoHayLugarParaDesmontar
                Call ShowConsoleMsg("¡No hay lugar para desmontar!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.UserHaDesmontado
                Call ShowConsoleMsg("¡Hás desmontado la criatura!", .red, .green, .blue, True, .italic)
                Exit Sub

            Case eMessages.UserHaDomado
                Call ShowConsoleMsg("¡Hás domado la criatura!", .red, .green, .blue, True, .italic)
                Exit Sub

            Case eMessages.UserNoPuedeUsarObjetoAqui
                Call ShowConsoleMsg("¡No puedes usar este objeto aquí!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.UserNoTieneConocimientosNecesarios
                Call ShowConsoleMsg("¡No tienes los conocimientos necesarios!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.UserTieneEmbarcacionAnclada
                Call ShowConsoleMsg("¡Tú embarcación se encuentra anclada, vé a buscarla!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.UserSeVuelveVisible
                Call ShowConsoleMsg("¡Has vuelto a ser visible!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.UserRecuperaSuAparienciaNormal
                Call ShowConsoleMsg("¡Hás recuperado tu apariencia normal!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.TeHasEscondidoEntreLasSombras
                Call ShowConsoleMsg("¡¡Te hás escondido entre las sombras!!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoHasLogradoEsconderte
                Call ShowConsoleMsg("¡No hás logrado esconderte!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoTienesSuficienteEnergia
                Call ShowConsoleMsg("¡No tienes suficiente energía!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.YaDomasteALaCriatura
                Call ShowConsoleMsg("¡Ya domaste a esa criatura!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.LaCriaturaTieneAmo
                Call ShowConsoleMsg("¡La criatura ya tiene dueño!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoPodesDomarDosCriaturasIguales
                Call ShowConsoleMsg("¡No podés domar dos criaturas del mismo tipo!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoMascotasZonaSegura
                Call ShowConsoleMsg("¡No se permiten mascotas en zona segura!. Éstas te esperarán afuera!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoHasLogradoDomarlo
                Call ShowConsoleMsg("¡No hás logrado domar a la criatura!!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoPodesControlarMasCriaturas
                Call ShowConsoleMsg("¡No podés controlar más criaturas!!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasPescado
                Call ShowConsoleMsg("¡Has pescado un lindo pez!!", .red, .green, .blue, True, .italic)
                Exit Sub

            Case eMessages.HasPescadoAlgunosPeces
                Call ShowConsoleMsg("¡Has pescado algunos peces!!", .red, .green, .blue, True, .italic)
                Exit Sub

            Case eMessages.NoHasPescado
                Call ShowConsoleMsg("¡No hás pescado nada!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.LaRedSeUsaEnMarAbierto
                Call ShowConsoleMsg("La red sólo puede ser utilizada desde el mar abierto.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasPescadoAlgunosPeces
                Call ShowConsoleMsg("¡Hás pescado algunos peces!", .red, .green, .blue, True, .italic)
                Exit Sub

            Case eMessages.EstasMuyCansadoParaRobar
                Call ShowConsoleMsg("Estás muy cansado para robar.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.EstasMuyCansadaParaRobar
                Call ShowConsoleMsg("Estás muy cansada para robar.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.LeHasRobadoMonedasDeOro
                Number = incomingData.ReadInteger()
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg("Le hás robado " & Number & " monedas de oro a " & nombre, .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoTieneOro
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg(nombre & " no tiene oro.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoHasLogradoRobarNada
                Call ShowConsoleMsg("¡No has logrado robar nada!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HanIntentadoRobarte
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg("¡" & nombre & " ha intentado robarte!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.UserNoTieneObjetos
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg("¡" & nombre & " no tiene objetos!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasRobadoObjetos
                Number = incomingData.ReadInteger()
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg("Has robado " & Number & " " & nombre, .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasHurtadoObjetos
                Number = incomingData.ReadInteger()
                nombre = incomingData.ReadASCIIString()
                Call ShowConsoleMsg("Has hurtado " & Number & " " & nombre, .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasConseguidoLena
                Call ShowConsoleMsg("¡Has conseguido algo de leña!", .red, .green, .blue, True, .italic)
                Exit Sub

            Case eMessages.NoHasConseguidoLena
                Call ShowConsoleMsg("¡No hás conseguido leña!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasExtraidoMinerales
                Call ShowConsoleMsg("¡Has extraido algunos minerales!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoHasExtraidoMinerales
                Call ShowConsoleMsg("¡No has conseguido nada!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.TerminasDeMeditar
                Call ShowConsoleMsg("Has terminado de meditar.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.HasRecuperadoMana
                Number = incomingData.ReadInteger()
                Call ShowConsoleMsg("¡Hás recuperado " & Number & " puntos de mana!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.UserMonturaSiendoAtacada
                Call ShowConsoleMsg("¡Están atacando a la criatura que hás domado!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.CriaturaAgonizando
                Call ShowConsoleMsg("¡La criatura está agonizando! Habrá mas posibilidades de domarla!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.EstasMuyCansadoParaLuchar
                Call ShowConsoleMsg("Estás muy cansado para luchar.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.EstasMuyCansadaParaLuchar
                Call ShowConsoleMsg("Estás muy cansada para luchar.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoPodesAtacarEsteNPC
                Call ShowConsoleMsg("No podés atacar este NPC.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.LejosParaDisparar
                Call ShowConsoleMsg("Estás muy lejos para disparar.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.VictimaMuerto
                Call ShowConsoleMsg("No podés atacar a un espíritu.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoPodesPelearAqui
                Call ShowConsoleMsg("No podés pelear aquí.", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.NoPodesAtacarteAVosMismo
                Call ShowConsoleMsg("¡No puedes atacarte a vos mismo!", .red, .green, .blue, .bold, .italic)
                Exit Sub

            Case eMessages.LejosParaAtacar
                Call ShowConsoleMsg("Estás demasiado lejos para atacar.", .red, .green, .blue, .bold, .italic)
                Exit Sub

        End Select
        
    End With
        
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)

        Select Case MessageType

            Case eMessages.DebesQuitarSeguroParaRobarCiudadano
                Call ShowConsoleMsg("Debes quitarte el seguro para robarle a un ciudadano.", .red, .green, .blue, .bold, .italic)
                Exit Sub

        End Select

    End With

End Sub

Private Sub HandleFirstInfo()

    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserClase = incomingData.ReadByte()
    UserFaccion = incomingData.ReadByte()
    
    If UserClase > 0 Then
        frmMain.imgMiniaturaClase.Picture = LoadPictureEX(UserClase & ".jpg")

    End If
    
    Call SetCursor(General)
    
    SetDayLight (True)
    
    AlphaSalir = 0

End Sub

Private Sub HandleCuentaRegresiva()

    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    Dim counter As Integer

    counter = incomingData.ReadInteger()
  
    If counter <= 9 Then
        AlphaCuenta = 255

        If counter = 0 Then
            Call ShowConsoleMsg("YAAAAA!!", 130, 200, 200, True, False)

        End If

    Else
        Call ShowConsoleMsg(counter, 255, 200, 130, True, False)

    End If
    
    CUENTA = counter
  
End Sub

Private Sub HandlePicInRender()

    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    Dim PicType As Byte

    PicType = incomingData.ReadByte()
  
    Select Case PicType

        Case ePicRenderType.BloodDie
            AlphaBloodUserDie = 255

        Case ePicRenderType.Blood
            AlphaBlood = 255

        Case ePicRenderType.Ceguera
            AlphaCeguera = 255

        Case ePicRenderType.TextKills
            TextKillsType = incomingData.ReadByte()
            AlphaTextKills = 255
            Call Audio.PlayWave(258 + TextKillsType)

    End Select
  
End Sub

Private Sub HandleQuit()

    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    Dim counter    As Byte

    Dim cancelExit As Byte

    counter = incomingData.ReadByte()
    cancelExit = incomingData.ReadByte()
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)

        If cancelExit = 1 Then
            Call ShowConsoleMsg("/Salir cancelado...", .red, .green, .blue, True, .italic)
            AlphaSalir = 0
            Exit Sub
        ElseIf counter > 0 Then
            Call ShowConsoleMsg("Cerrando... En " & counter & " segundos... ", .red, .green, .blue, True, .italic)
            AlphaSalir = 1
            Exit Sub
        Else
            Call ShowConsoleMsg("Cerrando...", .red, .green, .blue, True, .italic)
            AlphaSalir = 0
            Exit Sub

        End If

    End With
  
End Sub

Public Sub WriteCuentaRegresiva(ByVal Second As Byte)
 
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.CuentaRegresiva)
        Call .WriteByte(Second)

    End With

End Sub

Public Sub WriteLoginExistingChar()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LoginExistingChar" message to the outgoing data buffer
    '***************************************************
    Dim I As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingChar)
        Call .WriteASCIIString(UserAccount)
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIStringFixed(UserPassword)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
        #If SeguridadAlkon Then
            Call .WriteASCIIStringFixed(MD5HushYo)
        #End If

    End With

End Sub

Public Sub WriteOpenAccount()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LoginExistingChar" message to the outgoing data buffer
    '***************************************************
    Dim I As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.OpenAccount)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIStringFixed(UserPassword)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
        #If SeguridadAlkon Then
            Call .WriteASCIIStringFixed(MD5HushYo)
        #End If

    End With

End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LoginNewChar" message to the outgoing data buffer
    '***************************************************
    Dim I As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        Call .WriteASCIIString(UserAccount)
        Call .WriteASCIIString(UserName)

        Call .WriteASCIIStringFixed(UserPassword)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
        #If SeguridadAlkon Then
            Call .WriteASCIIStringFixed(MD5HushYo)
        #End If
        
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        
        Call .WriteByte(UserHogar)

    End With

End Sub

Public Sub WriteLoginNewAccount()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LoginNewChar" message to the outgoing data buffer
    '***************************************************
    Dim I As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewAccount)
        Call .WriteASCIIString(UserAccount)
        Call .WriteASCIIString(UserEmail)
        Call .WriteASCIIStringFixed(UserPassword)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
        #If SeguridadAlkon Then
            Call .WriteASCIIStringFixed(MD5HushYo)
        #End If
      
    End With

End Sub

Public Sub WriteBorrarPJ(ByVal pj As String)
    
    With outgoingData
        Call .WriteByte(ClientPacketID.BorrarPJ)
                
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(pj)
        
        Call .WriteASCIIStringFixed(MD5(UserPassword))
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
        #If SeguridadAlkon Then
            Call .WriteASCIIStringFixed(MD5HushYo)
        #End If

    End With

End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Talk" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(chat)

    End With

End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Yell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        
        Call .WriteASCIIString(chat)

    End With

End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharIndex As Integer, ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Whisper" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        
        Call .WriteInteger(CharIndex)
        
        Call .WriteASCIIString(chat)

    End With

End Sub

Public Sub WriteRequestPartyForm()
    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "RequestPartyForm" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPartyForm)

End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Walk" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(Heading)

    End With

End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestPositionUpdate" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)

End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Attack" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Attack)
    Call outgoingData.WriteByte(charlist(UserCharIndex).Heading)

End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PickUp" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PickUp)

End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.SafeToggle)
  

End Sub

''
' Writes the "ResuscitationSafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()
    '**************************************************************
    'Author: Rapsodius
    'Creation Date: 10/10/07
    'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
    '**************************************************************
    Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)

End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)

End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestAtributes" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)

End Sub

''
' Writes the "RequestFame" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestFame()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestFame" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestFame)

End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestSkills" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)

End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestMiniStats" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)

End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)

End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)

End Sub

Public Sub WriteUserCommerceConfirm()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "UserCommerceConfirm" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceConfirm)

End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)

End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/10/07
    'Writes the "UserCommerceOk" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)

End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceReject" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)

End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal slot As Byte, _
                     ByVal Amount As Integer, _
                     Optional ByVal DropX As Integer = 0, _
                     Optional DropY As Integer = 0)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Drop" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
        Call .WriteInteger(DropX)
        Call .WriteInteger(DropY)

    End With

End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(slot)

    End With

End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Integer, ByVal Y As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LeftClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteInteger(X)
        Call .WriteInteger(Y)

    End With

End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Integer, ByVal Y As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DoubleClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteInteger(X)
        Call .WriteInteger(Y)

    End With

End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Work" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)

    End With

End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UseSpellMacro" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)

End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UseItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(slot)

    End With

End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal item As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftBlacksmith" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(item)

    End With

End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal item As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(item)

    End With

End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Integer, _
                              ByVal Y As Integer, _
                              ByVal Skill As eSkill)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkLeftClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        
        Call .WriteByte(Skill)

    End With

End Sub

Public Sub WriteCreateEfectoClient(ByVal X As Integer, _
                                   ByVal Y As Integer, _
                                   ByVal Effect As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkLeftClick" message to the outgoing data buffer
    '***************************************************

    With outgoingData
        Call .WriteByte(ClientPacketID.CreateEfectoClient)
        
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteByte(Effect)

    End With
    
    '    Dim mArroja As New clsArroja
    '    Call mArroja.Init(0, 0, 1, Wave, Effect, UserPos.X, UserPos.Y, X, Y)
    '    Arrojas.Add mArroja
End Sub

Public Sub WriteCreateEfectoClientAction(ByVal Effect As Byte, _
                                         X As Integer, _
                                         ByVal Y As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkLeftClick" message to the outgoing data buffer
    '***************************************************

    With outgoingData
        Call .WriteByte(ClientPacketID.CreateEfectoClientAction)
        
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteByte(Effect)

    End With

End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal Desc As String, _
                               ByVal Name As String, _
                               ByVal Site As String, _
                               ByRef Codex() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNewGuild" message to the outgoing data buffer
    '***************************************************
    Dim temp As String

    Dim I    As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(Desc)
        Call .WriteASCIIString(Name)
        Call .WriteASCIIString(Site)
        
        For I = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(I) & SEPARATOR
        Next I
        
        If Len(temp) Then temp = left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)

    End With

End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpellInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpellInfo)
        
        Call .WriteByte(slot)

    End With

End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "EquipItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(slot)

    End With

End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeHeading" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)

    End With

End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ModifySkills" message to the outgoing data buffer
    '***************************************************
    Dim I As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For I = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(I))
        Next I

    End With

End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Train" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)

    End With

End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal slot As Byte, ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceBuy" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)

    End With

End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal slot As Byte, ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)

    End With

End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal slot As Byte, ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceSell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)

    End With

End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal slot As Byte, ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDeposit" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)

    End With

End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal slotANT As Byte, ByVal slotNW As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MoveSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteByte(slotANT)
        Call .WriteByte(slotNW)

    End With

End Sub



''
' Writes the "MoveBank" message to the outgoing data buffer.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal slot As Byte)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 06/14/09
    'Writes the "MoveBank" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveBank)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)

    End With

End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ClanCodexUpdate" message to the outgoing data buffer
    '***************************************************
    Dim temp As String

    Dim I    As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        
        Call .WriteASCIIString(Desc)
        
        For I = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(I) & SEPARATOR
        Next I
        
        If Len(temp) Then temp = left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)

    End With

End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal slot As Byte, _
                                  ByVal Amount As Long, _
                                  ByVal OfferSlot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceOffer" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(slot)
        Call .WriteLong(Amount)
        Call .WriteByte(OfferSlot)

    End With

End Sub

Public Sub WriteCommerceChat(ByVal chat As String)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    'Writes the "CommerceChat" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceChat)
        
        Call .WriteASCIIString(chat)

    End With

End Sub

''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOfferPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)

    End With

End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOfferAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)

    End With

End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAllianceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildPeaceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAlliancePropList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.guild)
    Call outgoingData.WriteByte(ClientPacketIDGuild.GuildAlliancePropList)

End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildPeacePropList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.guild)
    Call outgoingData.WriteByte(ClientPacketIDGuild.GuildPeacePropList)

End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDeclareWar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal url As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNewWebsite" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildNewWebsite)
        
        Call .WriteASCIIString(url)

    End With

End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectNewMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)

    End With

End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildKickMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildKickMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildUpdateNews" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildUpdateNews)
        
        Call .WriteASCIIString(news)

    End With

End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOpenElections" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.guild)
    Call outgoingData.WriteByte(ClientPacketIDGuild.GuildOpenElections)

End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestMembership" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)

    End With

End Sub

''
' Writes the "ShowGuildNews" message to the outgoing data buffer.
'

Public Sub WriteShowGuildNews()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 21/02/2010
    'Writes the "ShowGuildNews" message to the outgoing data buffer
    '***************************************************
 
    outgoingData.WriteByte (ClientPacketIDGuild.ShowGuildNews)

End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.guild)
        Call .WriteByte(ClientPacketIDGuild.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Online" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Online)

End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/16/08
    'Writes the "Quit" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Quit)

End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeave" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)

End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestAccountState" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)

End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PetStand" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetStand)

End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PetFollow" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)

End Sub

''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.TrainList)

End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Rest" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Rest)

End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Meditate" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Meditate)

End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Resucitate" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)

End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Heal" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Heal)

End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Help" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Help)

End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestStats" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)

End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceStart" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)

End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankStart" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankStart)

End Sub

''
' Writes the "ComandosVarios" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComandosVarios(ByVal TipoComando As Byte, _
                               Optional ByVal Opcion As Byte = 0)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ComandosVarios" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ComandosVarios)
    Call outgoingData.WriteByte(TipoComando)
    Call outgoingData.WriteByte(Opcion)

End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Information" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Information)

End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Reward" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Reward)

End Sub

''
' Writes the "RequestMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMOTD()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestMOTD" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)

End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpTime" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UpTime)

End Sub

''
' Writes the "PartyLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyLeave()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyLeave" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyLeave)

End Sub

''
' Writes the "PartyCreate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyCreate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyCreate" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyCreate)

End Sub

''
' Writes the "PartyJoin" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyJoin()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyJoin" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyJoin)

End Sub

''
' Writes the "Inquiry" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiry()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Inquiry" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Inquiry)

End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "PartyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal Number As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CentinelReport" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(Number)

    End With

End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOnline" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)

End Sub

''
' Writes the "PartyOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyOnline" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyOnline)

End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoleMasterRequest" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMRequest" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)

End Sub

''
' Writes the "BugReport" message to the outgoing data buffer.
'
' @param    message The message explaining the reported bug.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBugReport(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BugReport" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.bugReport)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeDescription" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(Desc)

    End With

End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Punishments" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ChangePassword" message to the outgoing data buffer.
'
' @param    oldPass Previous password.
' @param    newPass New password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/10/07
    'Last Modified By: Rapsodius
    'Writes the "ChangePassword" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangePassword)
        
        Call .WriteASCIIStringFixed(MD5(oldPass))
      
        Call .WriteASCIIStringFixed(MD5(newPass))

    End With

End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Gamble" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)

    End With

End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "InquiryVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InquiryVote)
        
        Call .WriteByte(opt)

    End With

End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LeaveFaction" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)

End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)

    End With

End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDepositGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)

    End With

End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Denounce" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "GuildFundate" message to the outgoing data buffer.
'
' @param    clanType The alignment of the clan to be founded.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate(ByVal clanType As eClanType)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildFundate" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildFundate)
        
        Call .WriteByte(clanType)

    End With

End Sub

''
' Writes the "PartyKick" message to the outgoing data buffer.
'
' @param    username The user to kick fro mthe party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyKick)
            
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "PartySetLeader" message to the outgoing data buffer.
'
' @param    username The user to set as the party's leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySetLeader(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartySetLeader" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartySetLeader)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "PartyAcceptMember" message to the outgoing data buffer.
'
' @param    username The user to accept into the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyAcceptMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyAcceptMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyAcceptMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberList" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.GuildMemberList)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.GMMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowName" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.showName)

End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.OnlineRoyalArmy)

End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineChaosLegion" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.onlineChaosLegion)

End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoNearby" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.GoNearby)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Comment" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.comment)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerTime" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.serverTime)

End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Where" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.Where)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreaturesInMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.CreaturesInMap)
        
        Call .WriteInteger(Map)

    End With

End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget(X As Integer, Y As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarpMeToTarget" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        .WriteByte (ClientPacketID.WarpMeToTarget)
        .WriteInteger (X)
        .WriteInteger (Y)

    End With
    
    MainTimer.Restart (TimersIndex.PuedeMover)
    MainTimer.Restart (TimersIndex.SendRPU)
  
End Sub

Public Sub WriteIntercambiarInv(ByVal Slot1 As Byte, _
                                ByVal Slot2 As Byte, _
                                ByVal Banco As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.IntercambiarInv)
        
        Call .WriteByte(Slot1)
        Call .WriteByte(Slot2)
        Call .WriteBoolean(Banco)

    End With

End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, _
                         ByVal Map As Integer, _
                         ByVal X As Integer, _
                         ByVal Y As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarpChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(Map)
        
        Call .WriteInteger(X)
        Call .WriteInteger(Y)

    End With

End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Silence" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.Silence)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SOSShowList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.SOSShowList)

End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SOSRemove" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.SOSRemove)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoToChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.GoToChar)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "invisible" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.invisible)

End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMPanel" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.GMPanel)

End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestUserList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.RequestUserList)

End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Working" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.Working)

End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Hiding" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.Hiding)

End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal Time As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Jail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .WriteByte(Time)

    End With

End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPC" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.KillNPC)

End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarnUser" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)

    End With

End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, _
                         ByVal EditOption As eEditOptions, _
                         ByVal arg1 As String, _
                         ByVal arg2 As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "EditChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(EditOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)

    End With

End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharStats" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RequestCharStats)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RequestCharGold)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInventory" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharBank" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RequestCharBank)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharSkills" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReviveChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ReviveChar)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineGM" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.OnlineGM)

End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    'Writes the "OnlineMap" message to the outgoing data buffer
    '26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.OnlineMap)
        
        Call .WriteInteger(Map)

    End With

End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Forgive" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.Forgive)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Kick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.Kick)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Execute" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.Execute)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.BanChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(reason)

    End With

End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.UnbanChar)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCFollow" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.NPCFollow)

End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SummonChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.SummonChar)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnListRequest" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.SpawnListRequest)

End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnCreature" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)

    End With

End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetNPCInventory" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ResetNPCInventory)

End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CleanWorld" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.CleanWorld)

End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ServerMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NickToIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.NickToIP)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "IPToNick" message to the outgoing data buffer
    '***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim I As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.IPToNick)
        
        For I = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(I))
        Next I

    End With

End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOnlineMembers" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, _
                               ByVal X As Integer, _
                               ByVal Y As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TeleportCreate" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.TeleportCreate)
        
        Call .WriteInteger(Map)
        
        Call .WriteInteger(X)
        Call .WriteInteger(Y)

    End With

End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TeleportDestroy" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.TeleportDestroy)

End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.RainToggle)

End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetCharDescription" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.SetCharDescription)
        
        Call .WriteASCIIString(Desc)

    End With

End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceMIDIToMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        
        Call .WriteInteger(Map)

    End With

End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, _
                               ByVal Map As Integer, _
                               ByVal X As Integer, _
                               ByVal Y As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceWAVEToMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        
        Call .WriteInteger(Map)
        
        Call .WriteInteger(X)
        Call .WriteInteger(Y)

    End With

End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CitizenMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.CitizenMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CriminalMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.CriminalMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TalkAsNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.TalkAsNPC)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.DestroyAllItemsInArea)

End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ItemsInTheFloor" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ItemsInTheFloor)

End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MakeDumb" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.MakeDumb)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MakeDumbNoMore" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumpIPTables" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.DumpIPTables)

End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.CouncilKick)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetTrigger" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.SetTrigger)
        
        Call .WriteByte(Trigger)

    End With

End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 04/13/07
    'Writes the "AskTrigger" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.AskTrigger)

End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPList" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.BannedIPList)

End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPReload" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.BannedIPReload)

End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildBan" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.GuildBan)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, _
                      ByRef Ip() As Byte, _
                      ByVal Nick As String, _
                      ByVal reason As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanIP" message to the outgoing data buffer
    '***************************************************
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim I As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then

            For I = LBound(Ip()) To UBound(Ip())
                Call .WriteByte(Ip(I))
            Next I

        Else
            Call .WriteASCIIString(Nick)

        End If
        
        Call .WriteASCIIString(reason)

    End With

End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanIP" message to the outgoing data buffer
    '***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim I As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.UnbanIP)
        
        For I = LBound(Ip()) To UBound(Ip())
            Call .WriteByte(Ip(I))
        Next I

    End With

End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As String, ByVal cantidad As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.CreateItem)
        
        Call .WriteInteger(cantidad)
        Call .WriteASCIIString(ItemIndex)

    End With

End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyItems" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.DestroyItems)

End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceWAVEAll" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ForceWAVEAll)
        
        Call .WriteByte(waveID)

    End With

End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, _
                                 ByVal punishment As Byte, _
                                 ByVal NewText As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemovePunishment" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RemovePunishment)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)

    End With

End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TileBlockedToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.TileBlockedToggle)

End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.KillNPCNoRespawn)

End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.KillAllNearbyNPCs)

End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LastIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.LastIP)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMOTD" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ChangeMOTD)

End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetMOTD" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.SetMOTD)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SystemMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.SystemMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.CreateNPC)
        
        Call .WriteInteger(NpcIndex)

    End With

End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.CreateNPCWithRespawn)
        
        Call .WriteInteger(NpcIndex)

    End With

End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ImperialArmour" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ImperialArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)

    End With

End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosArmour" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ChaosArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)

    End With

End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.NavigateToggle)

End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ServerOpenToUsersToggle)

End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnOffServer" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.TurnOffServer)

End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnCriminal" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.TurnCriminal)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetFactions" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ResetFactions)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharMail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.RequestCharMail)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterPassword" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)

    End With

End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterMail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.AlterMail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)

    End With

End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterName" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.AlterName)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)

    End With

End Sub

''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ToggleCentinelActivated)

End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DoBackup" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.DoBackUp)

End Sub

Public Sub WriteRetosAbrir()

    Call outgoingData.WriteByte(ClientPacketID.RetosAbrir)

End Sub

Public Sub WriteRetosCrear(Vs As Byte, _
                           PorItems As Boolean, _
                           Oro As Long, _
                           Pj2 As String, _
                           Pj3 As String, _
                           Pj4 As String)

    Call outgoingData.WriteByte(ClientPacketID.RetosCrear)
    Call outgoingData.WriteByte(Vs)
    Call outgoingData.WriteBoolean(PorItems)
    Call outgoingData.WriteLong(Oro)
    Call outgoingData.WriteASCIIString(Pj2)

    If Vs = 2 Then
        Call outgoingData.WriteASCIIString(Pj3)
        Call outgoingData.WriteASCIIString(Pj4)

    End If

End Sub

Public Sub WriteRetosDecide(Decide As Boolean)

    Call outgoingData.WriteByte(ClientPacketID.RetosDecide)
    Call outgoingData.WriteBoolean(Decide)

End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildMessages" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "InitCrafting" message to the outgoing data buffer.
'
' @param    Cantidad The final aumont of item to craft.
' @param    NroPorCiclo The amount of items to craft per cicle.

Public Sub WriteInitCrafting(ByVal cantidad As Long, ByVal NroPorCiclo As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    'Writes the "InitCrafting" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InitCrafting)
        Call .WriteLong(cantidad)
        
        Call .WriteInteger(NroPorCiclo)

    End With

End Sub

''
' Writes the "ItemUpgrade" message to the outgoing data buffer.
'
' @param    ItemIndex The index to the item to upgrade.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHome()

    '***************************************************
    'Author: Budi
    'Last Modification: 01/06/10
    'Writes the "Home" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Home)

    End With

End Sub

Public Sub WriteItemUpgrade(ByVal ItemIndex As Integer)
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 12/09/09
    'Writes the "ItemUpgrade" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ItemUpgrade)
    Call outgoingData.WriteInteger(ItemIndex)

End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveMap" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.SaveMap)

End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveChars" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.SaveChars)

End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CleanSOS" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.CleanSOS)

End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowServerForm" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ShowServerForm)

End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KickAllChars" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.KickAllChars)

End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadNPCs" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ReloadNPCs)

End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadServerIni" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ReloadServerIni)

End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadSpells" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ReloadSpells)

End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadObjects" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ReloadObjects)

End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Restart" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.Restart)

End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetAutoUpdate" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.ResetAutoUpdate)

End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal R As Byte, ByVal G As Byte, ByVal b As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatColor" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.ChatColor)
        
        Call .WriteByte(R)
        Call .WriteByte(G)
        Call .WriteByte(b)

    End With

End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Ignored" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.gm)
    Call outgoingData.WriteByte(ClientPacketIDGM.Ignored)

End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal slot As Byte)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "CheckSlot" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(slot)

    End With

End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/01/2007
    'Writes the "Ping" message to the outgoing data buffer
    '***************************************************
    'Prevent the timer from being cut
    If pingTime <> 0 Then Exit Sub
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    DoEvents
    
    pingTime = (GetTickCount() And &H7FFFFFFF)

End Sub

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, _
                          ByRef sClave As String, _
                          ByRef sValor As String)

    '***************************************************
    'Author: Brian Chaia (BrianPr)
    'Last Modification: 21/06/2009
    'Writes the "SetIniVar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.SetIniVar)
        
        Call .WriteASCIIString(sLlave)
        Call .WriteASCIIString(sClave)
        Call .WriteASCIIString(sValor)

    End With

End Sub

Public Sub WriteAddGM(ByRef nombre As String, ByVal Rango As Byte)

    '***************************************************
    'Author: Brian Chaia (BrianPr)
    'Last Modification: 21/06/2009
    'Writes the "SetIniVar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.gm)
        Call .WriteByte(ClientPacketIDGM.AddGM)
        
        Call .WriteByte(Rango)
        Call .WriteASCIIString(nombre)

    End With

End Sub

Public Sub WriteEquitar()

    With outgoingData
        Call .WriteByte(ClientPacketID.Equitar)

    End With

End Sub

Public Sub WriteDejarMontura()

    With outgoingData
        Call .WriteByte(ClientPacketID.DejarMontura)

    End With

End Sub

Public Sub WriteAnclarEmbarcacion()

    With outgoingData
        Call .WriteByte(ClientPacketID.AnclarEmbarcacion)

    End With

End Sub

'Public Sub WriteDesAnclarEmbarcacion()
'    With outgoingData
'        Call .WriteByte(ClientPacketID.DesAnclarEmbarcacion)
'    End With
'End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Sends all data existing in the buffer
    '***************************************************
    Dim sndData As String
    
    With outgoingData

        If .Length = 0 Then Exit Sub
        
        sndData = .ReadASCIIStringFixed(.Length)
        
        Call SendData(sndData)

    End With

End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados
    If Not ClientSetup.WinSock Then
        If frmMain.Client.State <> SockState.sckConnected Then Exit Sub
    Else

        If frmMain.WSock.State <> SockState.sckConnected Then Exit Sub

    End If

    Dim Data() As Byte
    
    Data = StrConv(sdData, vbFromUnicode)
    
    Call DataCorrect(DummyCode, Data, iCliente)
    
    sdData = StrConv(Data, vbUnicode)
    
    'Send data!
    If Not ClientSetup.WinSock Then
        Call frmMain.Client.SendData(sdData)
    Else
        Call frmMain.WSock.SendData(sdData)

    End If

End Sub

Public Sub DataCorrect(ByRef CodeKey() As Byte, _
                       ByRef DataIn() As Byte, _
                       ByRef varI As Integer)

    Dim I            As Long

    Dim intXOrValue2 As Integer

    Exit Sub

    For I = 0 To UBound(DataIn)
        varI = (varI + 1) Mod UBound(CodeKey)
        intXOrValue2 = CodeKey(varI)
        DataIn(I) = DataIn(I) Xor intXOrValue2
        
    Next I

End Sub

'quest
'quest

Public Sub WriteQuest()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete Quest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.Quest)

End Sub
 
Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestDetailsRequest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestDetailsRequest)
    
    Call outgoingData.WriteByte(QuestSlot)

End Sub
 
Public Sub WriteQuestAccept(ByVal ListInd As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestAccept al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(ClientPacketID.QuestAccept)
    Call outgoingData.WriteInteger(ListInd)

End Sub
 
Private Sub HandleQuestDetails()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestDetails del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.Length < 15 Then
        err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    Dim tmpStr         As String

    Dim tmpByte        As Byte

    Dim QuestEmpezada  As Boolean

    Dim I              As Integer
    
    Dim cantidadnpc    As Integer

    Dim NpcIndex       As Integer
    
    Dim cantidadobj    As Integer

    Dim OBJIndex       As Integer
    
    Dim AmountHave     As Integer
    
    Dim QuestIndex     As Integer
    
    Dim LevelRequerido As Byte

    Dim QuestRequerida As Integer
    
    FrmQuests.ListView2.ListItems.Clear
    FrmQuests.ListView1.ListItems.Clear
    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear
    
    FrmQuests.Image.BackColor = RGB(11, 11, 11)
    FrmQuests.picture1.BackColor = RGB(19, 14, 11)
    FrmQuests.Image.Refresh
    FrmQuests.picture1.Refresh
    FrmQuests.npclbl.Caption = ""
    FrmQuests.objetolbl.Caption = ""
    
    With buffer
        'Leemos el id del paquete
        Call .ReadByte
        
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(.ReadByte, True, False)
        
        If Not QuestEmpezada Then
        
            QuestIndex = .ReadInteger
        
            FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
           
            'tmpStr = "Mision: " & .ReadASCIIString & vbCrLf
            
            LevelRequerido = .ReadByte
            QuestRequerida = .ReadInteger
           
            If QuestRequerida <> 0 Then
                FrmQuestInfo.Text1.Text = QuestList(QuestIndex).Desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf & "Quest:" & QuestList(QuestRequerida).RequiredQuest
            Else
            
                FrmQuestInfo.Text1.Text = QuestList(QuestIndex).Desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf
            
            End If
           
            tmpByte = .ReadByte

            If tmpByte Then 'Hay NPCs
                If tmpByte > 5 Then
                    FrmQuestInfo.ListView1.FlatScrollBar = False
                Else
                    FrmQuestInfo.ListView1.FlatScrollBar = True
           
                End If

                For I = 1 To tmpByte
                    cantidadnpc = .ReadInteger
                    NpcIndex = .ReadInteger
               
                    ' tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & .ReadASCIIString & "."
                    If QuestEmpezada Then
                        tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & vbCrLf
                    Else
                        tmpStr = tmpStr & vbCrLf
                       
                        Dim subelemento As ListItem

                        Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(NpcIndex).Name)
                       
                        subelemento.SubItems(1) = cantidadnpc
                        subelemento.SubItems(2) = NpcIndex
                        subelemento.SubItems(3) = 0

                    End If

                Next I

            End If
           
            tmpByte = .ReadByte

            If tmpByte Then 'Hay OBJs

                For I = 1 To tmpByte
               
                    cantidadobj = .ReadInteger
                    OBJIndex = .ReadInteger
                    
                    AmountHave = .ReadInteger
                   
                    Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(OBJIndex).Name)
                    subelemento.SubItems(1) = AmountHave & "/" & cantidadobj
                    subelemento.SubItems(2) = OBJIndex
                    subelemento.SubItems(3) = 1
                Next I

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
            'tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
            'tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
           
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")
                       
            subelemento.SubItems(1) = .ReadLong
            subelemento.SubItems(2) = 12
            subelemento.SubItems(3) = 0
           
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                       
            subelemento.SubItems(1) = .ReadLong
            subelemento.SubItems(2) = 608
            subelemento.SubItems(3) = 1
           
            tmpByte = .ReadByte

            If tmpByte Then

                For I = 1 To tmpByte
                    'tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadInteger & vbCrLf
                   
                    Dim cantidadobjs As Integer

                    Dim obindex      As Integer
                   
                    cantidadobjs = .ReadInteger
                    obindex = .ReadInteger
                   
                    Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(obindex).Name)
                       
                    subelemento.SubItems(1) = cantidadobjs
                    subelemento.SubItems(2) = obindex
                    subelemento.SubItems(3) = 1
                           
                    ' Set subelemento = frmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                               
                    ' subelemento.SubItems(1) = .ReadInteger
                    ' subelemento.SubItems(2) = 0
                    ' subelemento.SubItems(3) = 1
           
                Next I

            End If

        Else
        
            QuestIndex = .ReadInteger
        
            FrmQuests.titulo.Caption = QuestList(QuestIndex).nombre
           
            LevelRequerido = .ReadByte
            QuestRequerida = .ReadInteger
           
            FrmQuests.detalle.Text = QuestList(QuestIndex).Desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf

            If QuestRequerida <> 0 Then
                FrmQuests.detalle.Text = FrmQuests.detalle.Text & vbCrLf & "Quest: " & QuestList(QuestRequerida).nombre

            End If
            
            'tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
            'tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
           
            tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
           
            tmpByte = .ReadByte

            If tmpByte Then 'Hay NPCs

                For I = 1 To tmpByte
                    cantidadnpc = .ReadInteger
                    NpcIndex = .ReadInteger
               
                    Dim matados As Integer
               
                    matados = .ReadInteger
                                     
                    Set subelemento = FrmQuests.ListView1.ListItems.Add(, , NpcData(NpcIndex).Name)
                       
                    Dim cantok As Integer

                    cantok = cantidadnpc - matados
                       
                    If cantok = 0 Then
                        subelemento.SubItems(1) = "OK"
                    Else
                        subelemento.SubItems(1) = matados & "/" & cantidadnpc

                    End If
                        
                    ' subelemento.SubItems(1) = cantidadnpc - matados
                    subelemento.SubItems(2) = NpcIndex
                    subelemento.SubItems(3) = 0
                    'End If
                Next I

            End If
           
            tmpByte = .ReadByte

            If tmpByte Then 'Hay OBJs

                For I = 1 To tmpByte
               
                    cantidadobj = .ReadInteger
                    OBJIndex = .ReadInteger
                    
                    AmountHave = .ReadInteger
                   
                    Set subelemento = FrmQuests.ListView1.ListItems.Add(, , ObjData(OBJIndex).Name)
                    subelemento.SubItems(1) = AmountHave & "/" & cantidadobj
                    subelemento.SubItems(2) = OBJIndex
                    subelemento.SubItems(3) = 1
                Next I

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
            'tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
            'tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
           
            Dim tmpLong As Long
           
            tmpLong = .ReadLong
           
            If tmpLong <> 0 Then
                Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Oro")
                subelemento.SubItems(1) = tmpLong
                subelemento.SubItems(2) = 12
                subelemento.SubItems(3) = 0

            End If
            
            tmpLong = .ReadLong
           
            If tmpLong <> 0 Then
                Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Experiencia")
                           
                subelemento.SubItems(1) = tmpLong
                subelemento.SubItems(2) = 608
                subelemento.SubItems(3) = 1

            End If
           
            tmpByte = .ReadByte

            If tmpByte Then

                For I = 1 To tmpByte
                    'tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadInteger & vbCrLf
                   
                    cantidadobjs = .ReadInteger
                    obindex = .ReadInteger
                   
                    Set subelemento = FrmQuests.ListView2.ListItems.Add(, , ObjData(obindex).Name)
                       
                    subelemento.SubItems(1) = cantidadobjs
                    subelemento.SubItems(2) = obindex
                    subelemento.SubItems(3) = 1
                           
                    ' Set subelemento = frmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                               
                    ' subelemento.SubItems(1) = .ReadInteger
                    ' subelemento.SubItems(2) = 0
                    ' subelemento.SubItems(3) = 1
           
                Next I

            End If
        
        End If

    End With
    
    'Determinamos que formulario se muestra, segï¿½n si recibimos la informaciï¿½n y la quest estï¿½ empezada o no.
    If QuestEmpezada Then
        FrmQuests.txtInfo.Text = tmpStr
        Call FrmQuests.ListView1_Click
        Call FrmQuests.ListView2_Click
        Call FrmQuests.lstQuests.SetFocus
    Else
        ' frmQuestInfo.txtInfo.Text = tmpStr
        FrmQuestInfo.Show vbModeless, frmMain
        'FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
        Call FrmQuestInfo.ListView1_Click
        Call FrmQuestInfo.ListView2_Click

    End If
    
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    If err.Number <> 0 And err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then err.Raise Error

End Sub
 
Public Sub HandleQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.Length < 1 Then
        err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    Dim I       As Integer

    Dim tmpByte As Byte

    Dim tmpStr  As String
    
    'Leemos el id del paquete
    Call buffer.ReadByte
     
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = buffer.ReadByte
    
    'Limpiamos el ListBox y el TextBox del formulario
    FrmQuests.lstQuests.Clear
    FrmQuests.txtInfo.Text = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = buffer.ReadASCIIString
        
        'Agregamos los items
        For I = 1 To tmpByte
            FrmQuests.lstQuests.AddItem ReadField(I, tmpStr, 45)
        Next I

    End If
    
    'Mostramos el formulario
    
    COLOR_AZUL = RGB(0, 0, 0)
    'Call Establecer_Borde(FrmQuests.lstQuests, FrmQuests, COLOR_AZUL, 0, 0)
    'FrmQuests.Picture = LoadInterface("ventanadetallemision.bmp")
    FrmQuests.Show vbModeless, frmMain
    
    'Pedimos la informaciï¿½n de la primer quest (si la hay)
    If tmpByte Then Call Protocol.WriteQuestDetailsRequest(1)
    
    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(buffer)
 
ErrHandler:

    If err.Number <> 0 And err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then err.Raise Error

End Sub

Public Sub WriteQuestListRequest()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestListRequest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestListRequest)

End Sub
 
Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestAbandon al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el ID del paquete.
    Call outgoingData.WriteByte(ClientPacketID.QuestAbandon)
    
    'Escribe el Slot de Quest.
    Call outgoingData.WriteByte(QuestSlot)

End Sub

'questtt
'quest

Public Sub HandleNpcQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.Length < 14 Then
        err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    Dim tmpStr         As String

    Dim tmpByte        As Byte

    Dim QuestEmpezada  As Boolean

    Dim I              As Integer
    
    Dim J              As Byte
    
    Dim cantidadnpc    As Integer

    Dim NpcIndex       As Integer
    
    Dim cantidadobj    As Integer

    Dim OBJIndex       As Integer
    
    Dim QuestIndex     As Integer
    
    Dim estado         As Byte
    
    Dim LevelRequerido As Byte

    Dim QuestRequerida As Integer
    
    Dim CantidadQuest  As Byte

    Dim subelemento    As ListItem
    
    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear
    
    With buffer
        'Leemos el id del paquete
        Call .ReadByte
        
        CantidadQuest = .ReadByte
            
        For J = 1 To CantidadQuest
        
            QuestIndex = .ReadInteger
            
            FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
               
            'tmpStr = "Mision: " & .ReadASCIIString & vbCrLf
               
            QuestList(QuestIndex).RequiredLevel = .ReadByte
                
            QuestList(QuestIndex).RequiredQuest = .ReadInteger
                
            ' FrmQuestInfo.Text1 = QuestList(QuestIndex).desc & vbCrLf & "Nivel requerido: " & QuestList(QuestIndex).RequiredLevel & vbCrLf
            'tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
            'tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
               
            tmpByte = .ReadByte
    
            If tmpByte Then 'Hay NPCs
                If tmpByte > 5 Then
                    FrmQuestInfo.ListView1.FlatScrollBar = False
                Else
                    FrmQuestInfo.ListView1.FlatScrollBar = True
               
                End If
                    
                ReDim QuestList(QuestIndex).RequiredNPC(1 To tmpByte)
                    
                For I = 1 To tmpByte
                                                
                    QuestList(QuestIndex).RequiredNPC(I).Amount = .ReadInteger
                    QuestList(QuestIndex).RequiredNPC(I).NpcIndex = .ReadInteger

                    '
    
                    '  Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(QuestList(QuestIndex).RequiredNPC(i).NpcIndex).Name)
                           
                    '   subelemento.SubItems(1) = QuestList(QuestIndex).RequiredNPC(i).Amount
                    '   subelemento.SubItems(2) = QuestList(QuestIndex).RequiredNPC(i).NpcIndex
                    '  subelemento.SubItems(3) = 0
    
                Next I

            Else
                ReDim QuestList(QuestIndex).RequiredNPC(0)

            End If
               
            tmpByte = .ReadByte
    
            If tmpByte Then 'Hay OBJs
                ReDim QuestList(QuestIndex).RequiredOBJ(1 To tmpByte)
    
                For I = 1 To tmpByte
                   
                    QuestList(QuestIndex).RequiredOBJ(I).Amount = .ReadInteger
                    QuestList(QuestIndex).RequiredOBJ(I).OBJIndex = .ReadInteger
                       
                    ' Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(QuestList(QuestIndex).RequiredOBJ(i).OBJIndex).Name)
                    ' subelemento.SubItems(1) = QuestList(QuestIndex).RequiredOBJ(i).Amount
                    ' subelemento.SubItems(2) = QuestList(QuestIndex).RequiredOBJ(i).OBJIndex
                    ' subelemento.SubItems(3) = 1
                Next I

            Else
                ReDim QuestList(QuestIndex).RequiredOBJ(0)
    
            End If
               
            QuestList(QuestIndex).RewardGLD = .ReadLong
            ' Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")
                           
            '  subelemento.SubItems(1) = QuestList(QuestIndex).RewardGLD
            ' subelemento.SubItems(2) = 12
            ' subelemento.SubItems(3) = 0
               
            '  Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                           
            QuestList(QuestIndex).RewardEXP = .ReadLong
            'subelemento.SubItems(1) = QuestList(QuestIndex).RewardEXP
            ' subelemento.SubItems(2) = 608
            ' subelemento.SubItems(3) = 1
               
            tmpByte = .ReadByte
    
            If tmpByte Then
                
                ReDim QuestList(QuestIndex).RewardOBJ(1 To tmpByte)
    
                For I = 1 To tmpByte
                                              
                    QuestList(QuestIndex).RewardOBJ(I).Amount = .ReadInteger
                    QuestList(QuestIndex).RewardOBJ(I).OBJIndex = .ReadInteger
                       
                    'Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(QuestList(QuestIndex).RewardOBJ(i).OBJIndex).Name)
                           
                    'subelemento.SubItems(1) = QuestList(QuestIndex).RewardOBJ(i).Amount
                    'subelemento.SubItems(2) = QuestList(QuestIndex).RewardOBJ(i).OBJIndex
                    'subelemento.SubItems(3) = 1
               
                Next I

            Else
                ReDim QuestList(QuestIndex).RewardOBJ(0)
    
            End If
                
            estado = .ReadByte
                
            Select Case estado
                
                Case 0
                    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                    subelemento.SubItems(1) = "Disponible"
                    subelemento.SubItems(2) = QuestIndex
                    subelemento.ForeColor = vbWhite
                    subelemento.ListSubItems(1).ForeColor = vbWhite
                        
                    'FrmQuestInfo.lstQuests.AddItem QuestIndex & "-" & QuestList(QuestIndex).nombre & "(Disponible)"
                Case 1
                    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                    subelemento.SubItems(1) = "En Curso"
                    subelemento.ForeColor = RGB(255, 175, 10)
                    subelemento.SubItems(2) = QuestIndex
                    subelemento.ListSubItems(1).ForeColor = RGB(255, 175, 10)
                    FrmQuestInfo.ListViewQuest.Refresh

                    'FrmQuestInfo.lstQuests.AddItem QuestIndex & "-" & QuestList(QuestIndex).nombre & "(En curso)"
                Case 2
                    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                    subelemento.SubItems(1) = "Finalizada"
                    subelemento.SubItems(2) = QuestIndex
                    subelemento.ForeColor = RGB(15, 140, 50)
                    subelemento.ListSubItems(1).ForeColor = RGB(15, 140, 50)
                    FrmQuestInfo.ListViewQuest.Refresh

                    ' FrmQuestInfo.lstQuests.AddItem QuestIndex & "-" & QuestList(QuestIndex).nombre & "(Realizada)"
                Case 3
                    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                    subelemento.SubItems(1) = "No disponible"
                    subelemento.SubItems(2) = QuestIndex
                    subelemento.ForeColor = RGB(255, 10, 10)
                    subelemento.ListSubItems(1).ForeColor = RGB(255, 10, 10)
                    FrmQuestInfo.ListViewQuest.Refresh
                
            End Select
                
        Next J

    End With
    
    'Determinamos que formulario se muestra, segï¿½n si recibimos la informaciï¿½n y la quest estï¿½ empezada o no.

    ' frmQuestInfo.txtInfo.Text = tmpStr
    FrmQuestInfo.Show vbModeless, frmMain
    'FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    'Call FrmQuestInfo.ListView1_Click
    'Call FrmQuestInfo.ListView2_Click
    
    Call incomingData.CopyBuffer(buffer)
    
ErrHandler:

    If err.Number <> 0 And err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then err.Raise Error
    
End Sub

Private Sub HandleUpdateNPCSimbolo()
    
    On Error GoTo HandleUpdateNPCSimbolo_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 4 Then
        err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim NpcIndex As Integer

    Dim simbolo  As Byte
    
    NpcIndex = incomingData.ReadInteger()
    
    simbolo = incomingData.ReadByte()

    charlist(NpcIndex).simbolo = simbolo
    
    Exit Sub

HandleUpdateNPCSimbolo_Err:

    ' Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateNPCSimbolo", Erl)
    Resume Next
    
End Sub

'aura
Private Sub HandleSendAura()

    On Error GoTo err

    Dim charindexx As Integer
 
    Call incomingData.ReadByte   'packetid
    charindexx = incomingData.ReadInteger
 
    With charlist(charindexx)

        Dim parte As UpdateAuras, aura1 As Byte

        parte = incomingData.ReadByte
        aura1 = incomingData.ReadByte
     
        If Not aura1 = 0 Then
            .aura(parte).AuraGrh = Auras(aura1).AuraGrh
            .aura(parte).R = Auras(aura1).R
            .aura(parte).G = Auras(aura1).G
            .aura(parte).b = Auras(aura1).b
            .aura(parte).Giratoria = Auras(aura1).Giratoria
            .aura(parte).OffSetX = Auras(aura1).OffSetX
            .aura(parte).OffSetY = Auras(aura1).OffSetY
        Else
            .aura(parte).AuraGrh = 0
            .aura(parte).R = 0
            .aura(parte).G = 0
            .aura(parte).b = 0
            .aura(parte).Giratoria = 0
            .aura(parte).OffSetX = 0
            .aura(parte).OffSetY = 0

        End If

        .aura(parte).Color = D3DColorXRGB(.aura(parte).R, .aura(parte).G, .aura(parte).b)

    End With
 
    Exit Sub
err:
    MsgBox "Error en HandleAura: (" & err.Number & ") " & err.Description

End Sub

'aura

' PAQUETES DE LOS EVENTOS '
Public Sub WriteNewEvent(ByVal Modality As eModalityEvent, _
                         ByVal Quotas As Byte, _
                         ByVal MinLvl As Byte, _
                         ByVal MaxLvl As Byte, _
                         ByVal GldInscription As Long, _
                         ByVal DspInscription As Long, _
                         ByVal CanjeInscription As Long, _
                         ByVal TimeInit As Long, _
                         ByVal TimeCancel As Long, _
                         ByVal TeamCant As Byte, _
                         ByRef AllowedClasses() As Byte)

    Dim LoopC As Integer
    
    With outgoingData
        Call .WriteByte(ClientPacketID.EventPacket)
        Call .WriteByte(EventPacketID.NewEvent)
        Call .WriteByte(Modality)
        Call .WriteByte(Quotas)
        Call .WriteByte(MinLvl)
        Call .WriteByte(MaxLvl)
        Call .WriteLong(GldInscription)
        Call .WriteLong(DspInscription)
        Call .WriteLong(CanjeInscription)
        Call .WriteLong(TimeInit)
        Call .WriteLong(TimeCancel)
        
        For LoopC = LBound(AllowedClasses()) To UBound(AllowedClasses())
            Call .WriteByte(AllowedClasses(LoopC))
        Next LoopC
        
        Call .WriteByte(TeamCant)

    End With

End Sub

Public Sub WriteCloseEvent(ByVal slot As Byte)

    With outgoingData
        Call .WriteByte(ClientPacketID.EventPacket)
        Call .WriteByte(EventPacketID.CloseEvent)
        Call .WriteByte(slot)

    End With

End Sub

Public Sub WriteRequiredEvents()

    With outgoingData
        Call .WriteByte(ClientPacketID.EventPacket)
        Call .WriteByte(EventPacketID.RequiredEvents)

    End With

End Sub

Public Sub WriteRequiredDataEvent(ByVal slot As Byte)

    With outgoingData
        Call .WriteByte(ClientPacketID.EventPacket)
        Call .WriteByte(EventPacketID.RequiredDataEvent)
        Call .WriteByte(slot)

    End With

End Sub

Public Sub WriteParticipeEvent(ByVal SlotEvent As String)

    With outgoingData
        Call .WriteByte(ClientPacketID.EventPacket)
        Call .WriteByte(EventPacketID.ParticipeEvent)
        Call .WriteASCIIString(SlotEvent)

    End With

End Sub

Public Sub WriteAbandonateEvent()

    With outgoingData
        Call .WriteByte(ClientPacketID.EventPacket)
        Call .WriteByte(EventPacketID.AbandonateEvent)

    End With

End Sub

Public Sub HandleEventPacketSv()

    If incomingData.Length < 2 Then
        err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler
    
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)

    Call buffer.ReadByte
        
    Dim PacketID As Byte

    Dim LoopC    As Integer

    Dim Modality As eModalityEvent

    Dim List()   As String
    
    PacketID = buffer.ReadByte()
    
    Select Case PacketID

        Case SvEventPacketID.SendListEvent
            frmPanelTorneo.cmbModalityCurso.Clear

            For LoopC = 1 To MAX_EVENT_SIMULTANEO
                Modality = buffer.ReadByte
                
                If Modality > 0 Then
                    frmPanelTorneo.cmbModalityCurso.AddItem strModality(Modality)
                Else
                    frmPanelTorneo.cmbModalityCurso.AddItem "Vacio"

                End If

            Next LoopC
            
        Case SvEventPacketID.SendDataEvent

            With frmPanelTorneo
                .lblQuotasCurso.Caption = "Inscriptos/Cupos: " & buffer.ReadByte & "/" & buffer.ReadByte
                .lblNivelCurso.Caption = "Nivel mínimo/máximo: " & buffer.ReadByte & "/" & buffer.ReadByte
                .lblOroCurso.Caption = "Oro acumulado: " & buffer.ReadLong
                .lblDspCurso.Caption = "Dsp acumulado: " & buffer.ReadLong
                .lblCanjeCurso.Caption = "Canje acumulado: " & buffer.ReadLong
                List = Split(buffer.ReadASCIIString, "-")
                
                .lstUsers.Clear
                
                For LoopC = LBound(List()) To UBound(List())
                    .lstUsers.AddItem UCase$(List(LoopC))
                Next LoopC
                
            End With

    End Select
    
    Call incomingData.CopyBuffer(buffer)

ErrHandler:

    Dim Error As Long

    Error = err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then err.Raise Error

End Sub

Private Function FormatChat(ByRef chat As String) As String()

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 07/28/07
    'Formats a dialog into different text lines.
    '**************************************************************
    Dim word        As String

    Dim curPos      As Long

    Dim Length      As Long

    Dim acumLength  As Long

    Dim lineLength  As Long

    Dim wordLength  As Long

    Dim curLine     As Long

    Dim chatLines() As String
    
    'Initialize variables
    curLine = 0
    curPos = 1
    Length = Len(chat)
    acumLength = 0
    lineLength = -1
    
    ReDim chatLines(FieldCount(chat, 32)) As String
    
    'Start formating
    Do While acumLength < Length
        word = Trim(ReadField(curPos, chat, 32))
        
        wordLength = Len(word)
        
        ' Is the first word of the first line? (it's the only that can start at -1)
        If lineLength = -1 Then
            chatLines(curLine) = word
            
            lineLength = wordLength
            acumLength = wordLength
        Else

            ' Is the word too long to fit in this line?
            If lineLength + wordLength + 1 > MAX_LENGTH Then
                'Put it in the next line
                curLine = curLine + 1
                chatLines(curLine) = word
                
                lineLength = wordLength
            Else
                'Add it to this line
                chatLines(curLine) = chatLines(curLine) & " " & word
                
                lineLength = lineLength + wordLength + 1

            End If
            
            acumLength = acumLength + wordLength + 1

        End If
        
        'Increase to search for next word
        curPos = curPos + 1
    Loop
    
    ' If it's only one line, center text
    If curLine = 0 And Length < MAX_LENGTH Then
        chatLines(curLine) = String((MAX_LENGTH - Length) \ 2 + 1, " ") & chatLines(curLine)

    End If
    
    'Resize array to fit
    ReDim Preserve chatLines(curLine) As String
    
    FormatChat = chatLines

End Function


Public Sub WriteSolicitarRanking(ByVal TIPO As eRanking)
    With outgoingData
        Call .WriteByte(ClientPacketID.SolicitaRranking)
        Call .WriteByte(TIPO)
    End With
End Sub

Public Sub HandleRecibirRanking()

'Recibimos el ranking
'
'
    If incomingData.Length < 3 Then
        err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)

    Dim Arrai() As String
    Dim Arrai2() As String
    Dim Mensaje As String
    Dim I As Integer

    Dim Cadena As String
    Dim Cadena1 As String

    'Leemos el id del paquete
    Call buffer.ReadByte

    'Leemos el string
    Cadena = buffer.ReadASCIIString
    Cadena1 = buffer.ReadASCIIString

    Arrai = Split(Cadena, "-")


    'redimensiono el array de listaprocesos
    ReDim Arrai2(LBound(Arrai()) To UBound(Arrai()))

    For I = 0 To 9
        Arrai2(I) = Arrai(I)
        Ranking.nombre(I) = Arrai2(I)
    Next I

    Arrai = Split(Cadena1, "-")

    For I = 0 To 9
        Arrai2(I) = Arrai(I)
        Ranking.value(I) = Arrai(I)
    Next I

    For I = 0 To 9
        If Ranking.nombre(I) = vbNullString Then
            FrmRanking2.Label1(I).Caption = "<Vacante>"
        Else
            If RankingOro = "$" Then
                FrmRanking2.Label1(I).Caption = Ranking.nombre(I) & " : $" & Ranking.value(I)
            Else
                FrmRanking2.Label1(I).Caption = Ranking.nombre(I) & " : " & Ranking.value(I)
            End If
        End If
        'Call ShowConsoleMsg(Ranking.Nombre(i) & "-" & Ranking.value(i))
    Next I

    Call FrmRanking2.Show(vbModeless, frmMain)

    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(buffer)

ErrHandler:
    Dim Error As Long
    Error = err.Number
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
       err.Raise Error
End Sub

