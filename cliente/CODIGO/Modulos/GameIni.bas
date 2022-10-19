Attribute VB_Name = "GameIni"
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

Public Type tCabecera 'Cabecera de los con

    Desc As String * 255
    CRC As Long
    MagicWord As Long

End Type

Public Type tGameIni

    Puerto As Long
    Musica As Byte
    fX As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer

End Type

Public Type tSetupMods

    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean
    bNoRes      As Boolean ' 24/06/2006 - ^[GS]^
    bNoSoundEffects As Boolean
    sGraficos   As String * 13
    bFPS As Boolean
    WinSock As Boolean

End Type

Public ClientSetup   As tSetupMods

Public MiCabecera    As tCabecera

Public Config_Inicio As tGameIni

Dim Lector           As New clsIniManager

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10

End Sub

Public Function LeerGameIni() As tGameIni

    Dim N       As Integer

    Dim GameIni As tGameIni

    N = FreeFile
    Open PathInit & "\Inicio.con" For Binary As #N
    Get #N, , MiCabecera
    
    Get #N, , GameIni
    
    Close #N
    LeerGameIni = GameIni

End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
    On Local Error Resume Next

    Dim N As Integer

    N = FreeFile
    Open PathInit & "\Inicio.con" For Binary As #N
    Put #N, , MiCabecera
    Put #N, , GameIniConfiguration
    Close #N

End Sub

Public Sub SaveConfig()
    On Local Error GoTo fileErr:
    
    'Set Lector = New clsIniManager
    Call Lector.Initialize(App.path & "/INIT/Config.ini")

    With mOpciones
        
        ' AUDIO
        Call Lector.ChangeValue("AUDIO", "Music", IIf(mOpciones.Music, "True", "False"))
        Call Lector.ChangeValue("AUDIO", "Sound", IIf(mOpciones.sound, "True", "False"))
        Call Lector.ChangeValue("AUDIO", "SoundEffects", IIf(mOpciones.SoundEffects, "True", "False"))
        Call Lector.ChangeValue("AUDIO", "VolMusic", mOpciones.VolMusic)
        Call Lector.ChangeValue("AUDIO", "VolSound", mOpciones.VolSound)
        
        ' GUILD
        Call Lector.ChangeValue("GUILD", "GuildNews", IIf(.GuildNews, "True", "False"))
        Call Lector.ChangeValue("GUILD", "DialogConsole", IIf(.DialogConsole, "True", "False"))
        Call Lector.ChangeValue("GUILD", "DialogCantMessages", CByte(.DialogCantMessages))
        
        ' SCREENSHOOTER
        Call Lector.ChangeValue("SCREENSHOOTER", "ScreenShooterNivelSuperior", IIf(mOpciones.ScreenShooterNivelSuperior, "True", "False"))
        Call Lector.ChangeValue("SCREENSHOOTER", "ScreenShooterNivelSuperiorIndex", mOpciones.ScreenShooterNivelSuperiorIndex)
        Call Lector.ChangeValue("SCREENSHOOTER", "ScreenShooterAlMorir", IIf(mOpciones.ScreenShooterAlMorir, "True", "False"))
        
        ' RECORDAR
        Call Lector.ChangeValue("CUENTA", "Recordar", IIf(mOpciones.Recordar, "True", "False"))
        Call Lector.ChangeValue("CUENTA", "RecordarUsuario", mOpciones.RecordarUsuario)
        Call Lector.ChangeValue("CUENTA", "RecordarPassword", mOpciones.RecordarPassword)
        
        ' VIDEO
        Call Lector.ChangeValue("VIDEO", "TransparencyTree", IIf(mOpciones.TransparencyTree, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "Shadows", IIf(mOpciones.Shadows, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "BlurEffects", IIf(mOpciones.BlurEffects, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "Niebla", IIf(mOpciones.Niebla, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "MostrarAyuda", IIf(mOpciones.MostrarAyuda, "True", "False"))
        ' OTROS
        Call Lector.ChangeValue("OTROS", "CursorFaccionario", IIf(mOpciones.CursorFaccionario, "True", "False"))
        
    End With
    
    Call Lector.DumpFile(App.path & "/INIT/Config.ini")
fileErr:

    If err.Number <> 0 Then
        MsgBox ("Ha ocurrido un error al guardar la configuracion del cliente. Error " & err.Number & " : " & err.Description)

    End If

End Sub

Public Sub ReadConfig()
'On Local Error GoTo fileErr:

' Set Lector = New clsIniManager
    Call Lector.Initialize(App.path & "/INIT/Config.ini")

    'With mOpciones

    ' AUDIO
    mOpciones.Music = Lector.GetValue("AUDIO", "Music")
    mOpciones.sound = Lector.GetValue("AUDIO", "Sound")
    mOpciones.SoundEffects = Lector.GetValue("AUDIO", "SoundEffects")
    mOpciones.VolMusic = Lector.GetValue("AUDIO", "VolMusic")
    mOpciones.VolSound = Lector.GetValue("AUDIO", "VolSound")

    ' GUILD
    mOpciones.GuildNews = Lector.GetValue("GUILD", "GuildNews")
    mOpciones.DialogConsole = Lector.GetValue("GUILD", "DialogConsole")
    mOpciones.DialogCantMessages = Lector.GetValue("GUILD", "DialogCantMessages")

    ' SCREENSHOOTER
    mOpciones.ScreenShooterNivelSuperior = Lector.GetValue("SCREENSHOOTER", "ScreenShooterNivelSuperior")
    mOpciones.ScreenShooterNivelSuperiorIndex = Lector.GetValue("SCREENSHOOTER", "ScreenShooterNivelSuperiorIndex")
    mOpciones.ScreenShooterAlMorir = Lector.GetValue("SCREENSHOOTER", "ScreenShooterAlMorir")

    ' RECORDAR
    mOpciones.Recordar = Lector.GetValue("CUENTA", "Recordar")
    mOpciones.RecordarUsuario = Lector.GetValue("CUENTA", "RecordarUsuario")
    mOpciones.RecordarPassword = Lector.GetValue("CUENTA", "RecordarPassword")

    ' VIDEO
    mOpciones.TransparencyTree = Lector.GetValue("VIDEO", "TransparencyTree")
    mOpciones.Shadows = Lector.GetValue("VIDEO", "Shadows")
    mOpciones.BlurEffects = Lector.GetValue("VIDEO", "BlurEffects")
    mOpciones.Niebla = Lector.GetValue("VIDEO", "Niebla")
    mOpciones.MostrarAyuda = Lector.GetValue("VIDEO", "MostrarAyuda")
    ' OTROS
    mOpciones.CursorFaccionario = Lector.GetValue("OTROS", "CursorFaccionario")

    #If Debugging Then

        PathGraficos = Lector.GetValue("PATH", "PathGraficos")
        PathRecursosCliente = Lector.GetValue("PATH", "PathRecursosCliente")
        PathWav = Lector.GetValue("PATH", "PathWav")
        PathInterface = Lector.GetValue("PATH", "PathInterface")
        PathInit = Lector.GetValue("PATH", "PathInit")
        IpServidor = Lector.GetValue("SERVIDOR", "IP")
        PuertoServidor = Lector.GetValue("SERVIDOR", "PUERTO")
    #Else
        IpServidor = "127.0.0.1"
        PuertoServidor = 7222
    #End If

    ' End With

    Call Lector.DumpFile(App.path & "/INIT/Config.ini")
fileErr:

    If err.Number <> 0 Then
        Call mOpciones_Default
        MsgBox ("ERROR - Config.ini cargado por defecto." & err.Description)

    End If

End Sub
