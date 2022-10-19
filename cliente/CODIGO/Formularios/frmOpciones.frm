VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   9345
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Caption         =   "Auras"
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   240
      TabIndex        =   32
      Top             =   5760
      Width           =   4215
      Begin VB.CheckBox ActAura 
         BackColor       =   &H80000012&
         Caption         =   "Activar Auras"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox RotaAura 
         BackColor       =   &H80000012&
         Caption         =   "Rotación"
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   2040
         TabIndex        =   33
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame frameOtros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Otros"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   6480
      Width           =   3615
      Begin VB.CheckBox ChkMostrarAyuda 
         BackColor       =   &H00000000&
         Caption         =   "Mostrar Ayudas"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1921
         TabIndex        =   35
         Top             =   40
         Width           =   1693
      End
      Begin VB.CheckBox chkCursorFaccionario 
         BackColor       =   &H00000000&
         Caption         =   "Cursor Faccionario"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   40
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdConfigDefault 
      Caption         =   "Restaurar configuración por defecto"
      Height          =   345
      Left            =   250
      MouseIcon       =   "frmOpciones.frx":18F33
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   8280
      Width           =   4215
   End
   Begin VB.Frame frameScreenShooter 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Screen Shooter"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   19
      Top             =   3840
      Width           =   4095
      Begin VB.TextBox txtScreenShooterNivel 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   22
         Text            =   "15"
         Top             =   110
         Width           =   375
      End
      Begin VB.CheckBox chkScreenShooterAlMorir 
         BackColor       =   &H00000000&
         Caption         =   "Al morir"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   470
         Width           =   3255
      End
      Begin VB.CheckBox chkScreenShooterNivelSuperior 
         BackColor       =   &H00000000&
         Caption         =   "Al matar personajes con Nivel superior a"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   110
         Width           =   3255
      End
   End
   Begin VB.Frame frameBotones 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   400
      TabIndex        =   14
      Top             =   7030
      Width           =   3975
      Begin VB.CommandButton cmdManual 
         Caption         =   "Manual de Argentum"
         Height          =   375
         Left            =   2040
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdChangePassword 
         Caption         =   "Cambiar Contraseña"
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton customMsgCmd 
         Caption         =   "Mensajes Personalizados"
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdCustomKeys 
         Caption         =   "Configurar Teclas"
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame frameNoticiasClan 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Noticias del Clan"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   4095
      Begin VB.CheckBox chkNoticiasClanNoMostrar 
         BackColor       =   &H00000000&
         Caption         =   "No Mostrar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox chkNoticiasClanMostrar 
         BackColor       =   &H00000000&
         Caption         =   "Mostrar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Audio"
      ForeColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   335
      TabIndex        =   6
      Top             =   840
      Width           =   4095
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Sonido 3D"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   850
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Sonidos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Musica"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   80
         Width           =   855
      End
      Begin MSComctlLib.Slider slMusica 
         Height          =   375
         Left            =   1320
         TabIndex        =   30
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Max             =   100
      End
      Begin MSComctlLib.Slider slSound 
         Height          =   375
         Left            =   1320
         TabIndex        =   31
         Top             =   465
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Max             =   100
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Diálogos de clan"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   4095
      Begin VB.TextBox txtCantMensajes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "5"
         Top             =   120
         Width           =   450
      End
      Begin VB.OptionButton optPantalla 
         BackColor       =   &H00000000&
         Caption         =   "En pantalla,"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optConsola 
         BackColor       =   &H00000000&
         Caption         =   "En consola"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   110
         TabIndex        =   2
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mensajes"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3240
         TabIndex        =   5
         Top             =   120
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar y Cerrar"
      Height          =   345
      Left            =   250
      MouseIcon       =   "frmOpciones.frx":19085
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   8760
      Width           =   4215
   End
   Begin VB.Frame frameVideo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Video"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   25
      Top             =   5050
      Width           =   4095
      Begin VB.CheckBox chkVideoShadows 
         BackColor       =   &H00000000&
         Caption         =   "Sombras"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   280
         Width           =   2055
      End
      Begin VB.CheckBox chkNiebla 
         BackColor       =   &H00000000&
         Caption         =   "Niebla"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkVideoBlurEffects 
         BackColor       =   &H00000000&
         Caption         =   "Efecto Diseminado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   0
         Width           =   1695
      End
      Begin VB.CheckBox chkVideoTransparencyTree 
         BackColor       =   &H00000000&
         Caption         =   "Transparencia en Arboles"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loading As Boolean

Private Sub ActAura_Click()

    If ActAura.Value = vbUnchecked Then
        WriteVar OpcionesPath, "AURAS", "AuraActiva", 0
        ActivarAuras = GetVar(OpcionesPath, "AURAS", "AuraActiva")
    Else
        WriteVar OpcionesPath, "AURAS", "AuraActiva", 1
        ActivarAuras = GetVar(OpcionesPath, "AURAS", "AuraActiva")

    End If

End Sub

Private Sub Check1_Click(Index As Integer)

    Select Case Index

        Case 0

            If Check1(0).Value = vbChecked Then
                mOpciones.Music = True
            Else
                mOpciones.Music = False

            End If
            
            Audio.MusicActivated = mOpciones.Music
            
        Case 1

            If Check1(0).Value = vbChecked Then
                mOpciones.sound = True
            Else
                mOpciones.sound = False

            End If
            
            Audio.SoundActivated = mOpciones.sound
            
        Case 2

            If Check1(0).Value = vbChecked Then
                mOpciones.SoundEffects = True
            Else
                mOpciones.SoundEffects = False

            End If
            
            Audio.SoundEffectsActivated = mOpciones.SoundEffects

    End Select

End Sub

Private Sub cmdConfigDefault_Click()
    Call mOpciones_Default
    Call LoadOptionsValues

End Sub

Private Sub cmdCustomKeys_Click()

    If Not loading Then Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)

End Sub

Private Sub cmdManual_Click()

    If Not loading Then Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://wiki.aoyind.com", "", App.path, SW_SHOWNORMAL)

End Sub

Private Sub cmdChangePassword_Click()
    Call frmNewPassword.Show(vbModal, Me)

End Sub

Private Sub Command2_Click()
    mOpciones.Music = Check1(0).Value
    mOpciones.sound = Check1(1).Value
    mOpciones.VolMusic = slMusica.Value
    mOpciones.VolSound = slSound.Value
    mOpciones.DialogConsole = Not optConsola.Value

    If IsNumeric(txtCantMensajes.Text) Then
        mOpciones.DialogCantMessages = CInt(txtCantMensajes.Text)
        DialogosClanes.CantidadDialogos = mOpciones.DialogCantMessages

    End If

    mOpciones.GuildNews = chkNoticiasClanMostrar.Value
    mOpciones.ScreenShooterNivelSuperior = chkScreenShooterNivelSuperior.Value

    If IsNumeric(txtScreenShooterNivel.Text) Then
        mOpciones.ScreenShooterNivelSuperiorIndex = CInt(txtScreenShooterNivel.Text)

    End If

    mOpciones.ScreenShooterAlMorir = chkScreenShooterAlMorir.Value
    mOpciones.SoundEffects = Check1(2).Value

    Audio.SoundActivated = mOpciones.sound
    Audio.MusicActivated = mOpciones.Music

    Audio.MusicVolume = mOpciones.VolMusic
    Audio.SoundVolume = mOpciones.VolSound
    Audio.SoundEffectsActivated = mOpciones.SoundEffects

    'VIDEO
    mOpciones.TransparencyTree = chkVideoTransparencyTree.Value
    mOpciones.Shadows = chkVideoShadows.Value
    mOpciones.BlurEffects = chkVideoBlurEffects.Value
    mOpciones.Niebla = chkNiebla.Value
    mOpciones.MostrarAyuda = ChkMostrarAyuda.Value

    'OTROS
    mOpciones.CursorFaccionario = chkCursorFaccionario.Value

    Call SetCursor(General)

    Call SaveConfig

    Unload Me
    frmMain.SetFocus

End Sub

Private Sub customMsgCmd_Click()
    Call frmMessageTxt.Show(vbModeless, Me)

End Sub

Private Sub Form_Load()
    loading = True      'Prevent sounds when setting check's values
    
    Me.Picture = LoadPictureEX("VENTANAOPCIONES.jpg")
    
    Call LoadOptionsValues
    
    loading = False     'Enable sounds when setting check's values

End Sub

Private Sub RotaAura_Click()

    If RotaAura.Value = vbUnchecked Then
        WriteVar OpcionesPath, "AURAS", "rotacion", 0
        RotarActivado = GetVar(OpcionesPath, "AURAS", "rotacion")
    Else
        WriteVar OpcionesPath, "AURAS", "rotacion", 1
        RotarActivado = GetVar(OpcionesPath, "AURAS", "rotacion")

    End If

End Sub

Private Sub slSound_Change()
    mOpciones.VolSound = slSound.Value
    Audio.SoundVolume = mOpciones.VolSound

End Sub

Private Sub slMusica_Change()
    mOpciones.VolMusic = slMusica.Value
    Audio.MusicVolume = mOpciones.VolMusic

End Sub

Private Sub txtCantMensajes_LostFocus()
    txtCantMensajes.Text = Trim$(txtCantMensajes.Text)

    If IsNumeric(txtCantMensajes.Text) Then
        DialogosClanes.CantidadDialogos = Trim$(txtCantMensajes.Text)
    Else
        txtCantMensajes.Text = 5

    End If

End Sub

Public Sub LoadOptionsValues()

    If mOpciones.Music = True Then
        Audio.MusicActivated = True
        Check1(0).Value = vbChecked
        slMusica.Value = mOpciones.VolMusic
    Else
        Audio.MusicActivated = False
        Check1(0).Value = vbUnchecked
        slMusica.Value = mOpciones.VolMusic

    End If
     
    If mOpciones.sound = True Then
        Audio.SoundActivated = True
        Check1(1).Value = vbChecked
        slSound.Value = mOpciones.VolSound
    Else
        Audio.SoundActivated = False
        Check1(1).Value = vbUnchecked
        slSound.Value = mOpciones.VolSound

    End If
     
    If mOpciones.SoundEffects = True Then
        Check1(2).Value = vbChecked
    Else
        Check1(2).Value = vbUnchecked

    End If
     
    If mOpciones.DialogConsole = True Then
        optConsola.Value = False
        optPantalla.Value = True
    Else
        optConsola.Value = True
        optPantalla.Value = False

    End If
     
    If (mOpciones.DialogCantMessages > 0) Then
        txtCantMensajes.Text = CStr(mOpciones.DialogCantMessages)
    Else
        txtCantMensajes.Text = 0

    End If
     
    If mOpciones.GuildNews = True Then
        chkNoticiasClanMostrar.Value = vbChecked
    Else
        chkNoticiasClanMostrar.Value = vbUnchecked

    End If
     
    If mOpciones.ScreenShooterNivelSuperior = True Then
        chkScreenShooterNivelSuperior.Value = vbChecked
        txtScreenShooterNivel = mOpciones.ScreenShooterNivelSuperiorIndex
    Else
        chkScreenShooterNivelSuperior.Value = vbUnchecked
        txtScreenShooterNivel = mOpciones.ScreenShooterNivelSuperiorIndex

    End If
     
    If mOpciones.ScreenShooterAlMorir = True Then
        chkScreenShooterAlMorir.Value = vbChecked
    Else
        chkScreenShooterAlMorir.Value = vbUnchecked

    End If
     
    Audio.MusicVolume = mOpciones.VolMusic
    Audio.SoundVolume = mOpciones.VolSound
    Audio.SoundEffectsActivated = mOpciones.SoundEffects
    
    'VIDEO
    If mOpciones.TransparencyTree = True Then
        chkVideoTransparencyTree.Value = vbChecked
    Else
        chkVideoTransparencyTree.Value = vbUnchecked

    End If
    
    If mOpciones.Shadows = True Then
        chkVideoShadows.Value = vbChecked
    Else
        chkVideoShadows.Value = vbUnchecked

    End If
    
     If mOpciones.MostrarAyuda = True Then
        ChkMostrarAyuda.Value = vbChecked
    Else
        ChkMostrarAyuda.Value = vbUnchecked

    End If
    
    If mOpciones.BlurEffects = True Then
        chkVideoBlurEffects.Value = vbChecked
    Else
        chkVideoBlurEffects.Value = vbUnchecked

    End If
    
    If mOpciones.Niebla = True Then
        chkNiebla.Value = vbChecked
    Else
        chkNiebla.Value = vbUnchecked

    End If
    
    'OTROS
    If mOpciones.CursorFaccionario = True Then
        chkCursorFaccionario.Value = vbChecked
    Else
        chkCursorFaccionario.Value = vbUnchecked

    End If
    
    If ActivarAuras = "1" Then
        frmOpciones.ActAura.Value = vbChecked
    Else
        frmOpciones.ActAura.Value = vbUnchecked

    End If
   
    If RotarActivado = "1" Then
        frmOpciones.RotaAura.Value = vbChecked
    Else
        frmOpciones.RotaAura.Value = vbUnchecked

    End If
          
End Sub

Private Sub txtScreenShooterNivel_LostFocus()
    txtScreenShooterNivel.Text = Trim$(txtScreenShooterNivel.Text)

    If Not IsNumeric(txtScreenShooterNivel.Text) Then
        txtScreenShooterNivel.Text = 15

    End If

End Sub
