VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "AoYind 3"
   ClientHeight    =   11685
   ClientLeft      =   0
   ClientTop       =   555
   ClientWidth     =   15330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainN2.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   779
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1022
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TCod 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5640
      MaxLength       =   160
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox barritaa 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   15180
      Picture         =   "frmMainN2.frx":0CCA
      ScaleHeight     =   5175
      ScaleWidth      =   180
      TabIndex        =   47
      Top             =   3480
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picarmadura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1425
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   10815
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdEfecto 
      Caption         =   "Efecto"
      Enabled         =   0   'False
      Height          =   315
      Left            =   12615
      TabIndex        =   20
      Top             =   10815
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   9120
      Top             =   240
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   7680
      Top             =   240
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   8160
      Top             =   240
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   9600
      Top             =   240
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   630
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   10800
      Visible         =   0   'False
      Width           =   11820
   End
   Begin VB.TextBox SendCMSTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   480
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   9600
      Visible         =   0   'False
      Width           =   11820
   End
   Begin VB.PictureBox pRender 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12225
      Left            =   -480
      ScaleHeight     =   815
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1080
      TabIndex        =   4
      Top             =   0
      Width           =   16200
      Begin VB.TextBox tRePass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6360
         MaxLength       =   160
         PasswordChar    =   "*"
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Chat"
         Top             =   3120
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.PictureBox PicSpells 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2385
         Left            =   14670
         ScaleHeight     =   159
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   48
         Top             =   6120
         Visible         =   0   'False
         Width           =   1920
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00E0E0E0&
            Height          =   495
            Left            =   480
            TabIndex        =   49
            Top             =   3600
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2385
         Left            =   14670
         ScaleHeight     =   159
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   44
         Top             =   3675
         Visible         =   0   'False
         Width           =   1920
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00E0E0E0&
            Height          =   495
            Left            =   480
            TabIndex        =   45
            Top             =   3600
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.PictureBox cmdinfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   15180
         Picture         =   "frmMainN2.frx":7180
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3480
         Width           =   120
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   15135
         ScaleHeight     =   660
         ScaleWidth      =   300
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   6510
         Width           =   300
         Begin VB.PictureBox cmdMoverHechi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   0
            Picture         =   "frmMainN2.frx":732C
            ScaleHeight     =   285
            ScaleWidth      =   210
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   350
            Width           =   210
         End
         Begin VB.PictureBox cmdMoverHechi 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   282
            Index           =   1
            Left            =   0
            Picture         =   "frmMainN2.frx":76B4
            ScaleHeight     =   285
            ScaleWidth      =   210
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   0
            Width           =   210
         End
      End
      Begin VB.PictureBox LanzarImg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   13200
         ScaleHeight     =   645
         ScaleWidth      =   645
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   5235
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.PictureBox BarraHechiz 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   15135
         Picture         =   "frmMainN2.frx":7A3C
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3885
         Visible         =   0   'False
         Width           =   240
         Begin VB.PictureBox BarritaHechiz 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   105
            Left            =   15
            Picture         =   "frmMainN2.frx":9AEE
            ScaleHeight     =   7
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   240
            Width           =   210
         End
      End
      Begin VB.PictureBox picHechiz 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         DrawStyle       =   3  'Dash-Dot
         ForeColor       =   &H00FFFFFF&
         Height          =   3405
         Left            =   12120
         MousePointer    =   99  'Custom
         ScaleHeight     =   227
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   28
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox invHechisos 
         BorderStyle     =   0  'None
         Height          =   3465
         Left            =   10680
         Picture         =   "frmMainN2.frx":9C64
         ScaleHeight     =   3465
         ScaleWidth      =   1140
         TabIndex        =   32
         Top             =   3720
         Visible         =   0   'False
         Width           =   1140
         Begin VB.Image CmdLanzar 
            Height          =   315
            Left            =   15
            MousePointer    =   99  'Custom
            Picture         =   "frmMainN2.frx":FAF1
            Top             =   3375
            Visible         =   0   'False
            Width           =   1125
         End
      End
      Begin VB.PictureBox imgMiniMapa 
         BorderStyle     =   0  'None
         Height          =   1498
         Left            =   13440
         Picture         =   "frmMainN2.frx":12594
         ScaleHeight     =   1500
         ScaleMode       =   0  'User
         ScaleWidth      =   1100
         TabIndex        =   29
         Top             =   570
         Width           =   1500
         Begin VB.Shape shpMiniMapaVision 
            Height          =   315
            Left            =   520
            Top             =   614
            Width           =   375
         End
         Begin VB.Shape shpMiniMapaUser 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000000C0&
            FillColor       =   &H000000FF&
            Height          =   45
            Left            =   695
            Top             =   750
            Width           =   45
         End
      End
      Begin VB.PictureBox picfondoinve 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   13920
         Picture         =   "frmMainN2.frx":19B08
         ScaleHeight     =   570
         ScaleWidth      =   1125
         TabIndex        =   27
         Top             =   3135
         Visible         =   0   'False
         Width           =   1125
         Begin VB.Label Lblmagia 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   585
            TabIndex        =   34
            Top             =   135
            Width           =   480
         End
         Begin VB.Label lblinve 
            BackStyle       =   0  'Transparent
            Height          =   405
            Left            =   150
            TabIndex        =   33
            Top             =   90
            Width           =   360
         End
      End
      Begin VB.PictureBox pConsola 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1348
         Left            =   2040
         ScaleHeight     =   90
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   788
         TabIndex        =   26
         Top             =   3120
         Visible         =   0   'False
         Width           =   11820
         Begin VB.Timer tMouse 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   8400
            Top             =   0
         End
         Begin VB.Timer tRelampago 
            Enabled         =   0   'False
            Interval        =   7500
            Left            =   0
            Top             =   0
         End
         Begin MSWinsockLib.Winsock WSock 
            Left            =   6960
            Top             =   0
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
      End
      Begin VB.TextBox tEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6360
         MaxLength       =   160
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Chat"
         Top             =   2280
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.TextBox tPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6420
         MaxLength       =   160
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Chat"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.TextBox tUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6420
         MaxLength       =   160
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Chat"
         Top             =   9285
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.PictureBox barritaaa 
         Height          =   2295
         Left            =   0
         ScaleHeight     =   2235
         ScaleWidth      =   315
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.PictureBox BarraConsola 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1367
      Left            =   8880
      Picture         =   "frmMainN2.frx":1D637
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7680
      Width           =   270
      Begin VB.PictureBox BarritaConsola 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   30
         Picture         =   "frmMainN2.frx":1EA61
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1020
         Width           =   210
      End
   End
   Begin VB.Label lblItemInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11295
      TabIndex        =   36
      Top             =   10650
      Width           =   3135
   End
   Begin VB.Label lblItem 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "55"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   3000
      TabIndex        =   7
      Top             =   3900
      Width           =   360
   End
   Begin VB.Label lbStats 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   12300
      TabIndex        =   24
      Top             =   7485
      Width           =   1035
   End
   Begin VB.Label MenuF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   13980
      TabIndex        =   23
      Top             =   7515
      Width           =   1005
   End
   Begin VB.Label exp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "300/300"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   6375
      TabIndex        =   22
      Top             =   195
      Width           =   765
   End
   Begin VB.Label lblDIATEST 
      Height          =   495
      Left            =   6225
      TabIndex        =   21
      Top             =   5580
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Image imgMiniaturaClase 
      Height          =   1080
      Left            =   12315
      Top             =   75
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblPosTest 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3240
      TabIndex        =   17
      Top             =   10260
      Width           =   1935
   End
   Begin VB.Image picResu 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   11100
      Stretch         =   -1  'True
      Top             =   11295
      Width           =   375
   End
   Begin VB.Label lblEscudo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6030
      TabIndex        =   16
      Top             =   11355
      Width           =   1095
   End
   Begin VB.Label lblCasco 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5055
      TabIndex        =   15
      Top             =   11355
      Width           =   1095
   End
   Begin VB.Label lblArma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4005
      TabIndex        =   14
      Top             =   11370
      Width           =   1095
   End
   Begin VB.Label lblArmadura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2910
      TabIndex        =   13
      Top             =   11355
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   14880
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8550
      TabIndex        =   12
      Top             =   11355
      Width           =   210
   End
   Begin VB.Label lblFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7770
      TabIndex        =   11
      Top             =   10575
      Width           =   210
   End
   Begin VB.Label lblUsers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   13455
      TabIndex        =   10
      Top             =   11400
      Width           =   570
   End
   Begin VB.Label Coord 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(000,000)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   13275
      TabIndex        =   8
      Top             =   6960
      Width           =   780
   End
   Begin VB.Image PicSeg 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   11505
      Stretch         =   -1  'True
      Top             =   11235
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   15120
      Top             =   0
      Width           =   375
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000000"
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   13155
      TabIndex        =   6
      Top             =   8040
      Width           =   840
   End
   Begin VB.Image Image3 
      Height          =   315
      Index           =   0
      Left            =   12690
      Top             =   7995
      Width           =   390
   End
   Begin VB.Label lbCRIATURA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   120
      Left            =   9000
      TabIndex        =   5
      Top             =   2400
      Width           =   30
   End
   Begin VB.Image iBEXPE 
      Height          =   135
      Left            =   13125
      Top             =   1665
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label lblSedN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   14055
      TabIndex        =   9
      Top             =   10560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image iBEXP 
      Height          =   300
      Left            =   13320
      Top             =   1440
      Visible         =   0   'False
      Width           =   1605
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AoYind 3.0
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


Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim s_Index As Integer    '< Número de slot.
Dim PicMoveX As Single
Dim PicMoveY As Single

Public tX As Integer

Public tY As Integer

Public MouseX As Integer

Public MouseY As Integer

Public MouseBoton As Long

Public MouseShift As Long

Private clicX As Long

Private clicY As Long

Public SinOrtografia As Boolean

'Dim gDSB As DirectSoundBuffer
'Dim gD As DSBUFFERDESC
'Dim gW As WAVEFORMATEX
Dim gFileName As String

'Dim dsE As DirectSoundEnum
'Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As PlayLoop

Dim PuedeMacrear As Boolean

Dim OldYConsola As Integer

Public hlst As clsGraphicalList

Dim InvX As Integer

Dim InvY As Integer

Public WithEvents Client As CSocketMaster
Attribute Client.VB_VarHelpID = -1
Public WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1

Private Sub Bar_Mana_Click(Index As Integer)

    If UserMinMAN = UserMaxMAN Then Exit Sub
            
    If UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If
                
    If Not PuedeMacrear Then
        AddtoRichPicture "¡No puedes usar el macro tan rápido!", 255, 255, 255, True, False, False
    ElseIf charlist(UserCharIndex).Moving = 0 Then
        Call WriteMeditate
        PuedeMacrear = False

    End If

End Sub

Private Sub Bar_Mana_DblClick(Index As Integer)

'    If Index = 0 Then
'        If Bar_Mana(1).Visible = True Then
'            Bar_Mana(1).Visible = False
'        Else
'            Bar_Mana(1).Visible = True
'
'        End If
'
'    End If

End Sub

Private Sub bar_salud_DblClick(Index As Integer)

'    If Index = 0 Then
'
'        'Helios Barras
'        If bar_salud(1).Visible = True Then
'            bar_salud(1).Visible = False
'        Else
'            bar_salud(1).Visible = True
'
'        End If
'
'    End If

End Sub

Private Sub Client_Connect()
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
    #If SeguridadAlkon Then
        Call ConnectionStablished(Socket1.PeerAddress)
    #End If
    
    Second.Enabled = True

    Select Case EstadoLogin

        Case E_MODO.CrearNuevoPj
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login
        
        Case E_MODO.Normal
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login
            iServer = 0
            iCliente = 0
            DummyCode = StrConv("damn" & StrReverse(UCase$(UserName)) & "you", vbFromUnicode)

        Case E_MODO.Cuentas
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login
    
        Case E_MODO.CrearCuenta
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login
    
        Case E_MODO.BorrarPersonaje
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login

    End Select

End Sub

Private Sub Client_CloseSck()

    Dim I As Long

    Client.CloseSck
    
    Call ClosePj

End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)

    Dim RD     As String

    Dim Data() As Byte
    
    Client.GetData RD
    Data = StrConv(RD, vbFromUnicode)
    
    Call DataCorrect(DummyCode, Data, iServer)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    NotEnoughData = False
    
    'Send buffer to Handle data
    Call HandleIncomingData

End Sub

Private Sub Client_Error(ByVal Number As Integer, _
                         Description As String, _
                         ByVal sCode As Long, _
                         ByVal Source As String, _
                         ByVal HelpFile As String, _
                         ByVal HelpContext As Long, _
                         CancelDisplay As Boolean)

    '*********************************************
    'Handle socket errors
    '*********************************************
    If Number = 24036 Then
        Call MessageBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    ElseIf Number = 10049 Then
        Call MessageBox("Su equipo no soporta la API de Socket, se cambiará su configuración a Winsock, si problema persiste contacte soporte.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    End If
    
    Call MessageBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    Second.Enabled = False

    Client.CloseSck

    If Not frmCrearPersonaje.Visible And Not Conectar Then
        Call ClosePj
    Else
        frmCrearPersonaje.MousePointer = 0

    End If

End Sub

Private Sub BarraConsola_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)

    Dim TempY  As Integer

    Dim TamCon As Integer

    TempY = Y - 3
    TamCon = (LineasConsola - 6)

    If TamCon > 0 Then
        If TempY < 16 Then
            If OffSetConsola > 0 Then OffSetConsola = OffSetConsola - 1
            TempY = 16 + OffSetConsola * 52 / TamCon
        ElseIf TempY > 68 Then

            If OffSetConsola < TamCon Then OffSetConsola = OffSetConsola + 1
            TempY = 16 + OffSetConsola * 52 / TamCon
        Else

            If LineasConsola <= 6 Then TempY = 68
            OffSetConsola = Int((TempY - 16) * TamCon / 52)

        End If

    Else
        TempY = 68

    End If

    BarritaConsola.top = TempY
    ReDrawConsola

End Sub

Private Sub BarraHechiz_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    Dim TempY  As Integer

    Dim TamCon As Integer

    TempY = Y - 3

    Dim MaxItems As Integer

    MaxItems = Int(picHechiz.Height / hlst.Pixel_Alto)
    TamCon = (hlst.ListCount - MaxItems)
    
    If TamCon > 0 Then
        If TempY < 16 Then
            If hlst.Scroll > 0 Then hlst.Scroll = hlst.Scroll - 1
            TempY = 16 + hlst.Scroll * 134 / TamCon
        ElseIf TempY > 150 Then

            If hlst.Scroll < TamCon Then hlst.Scroll = hlst.Scroll + 1
            TempY = 16 + hlst.Scroll * 134 / TamCon
        Else

            If hlst.ListCount <= MaxItems Then TempY = 150
            hlst.Scroll = Int((TempY - 16) * TamCon / 134)

        End If

    Else
        TempY = 150

    End If

    BarritaHechiz.top = TempY

End Sub

Private Sub BarritaConsola_MouseDown(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)

    If Button = 1 Then
        OldYConsola = Y

    End If

End Sub

Private Sub BarritaConsola_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)

    If Button = 1 Then

        Dim TempY As Integer

        TempY = BarritaConsola.top + (Y - OldYConsola)

        If TempY < 16 Then TempY = 16
        If TempY > 68 Then TempY = 68
        If LineasConsola <= 6 Then TempY = 68
        OffSetConsola = Int((TempY - 16) * (LineasConsola - 6) / 52)
        BarritaConsola.top = TempY
        ReDrawConsola

    End If

End Sub

Private Sub BarritaHechiz_MouseDown(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

    If Button = 1 Then
        hlst.OldY = Y

    End If

End Sub

Private Sub BarritaHechiz_MouseMove(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

    If Button = 1 Then

        Dim TempY    As Integer

        Dim MaxItems As Integer

        MaxItems = Int(picHechiz.Height / hlst.Pixel_Alto)
        TempY = BarritaHechiz.top + (Y - hlst.OldY)

        If TempY < 16 Then TempY = 16
        If TempY > 150 Then TempY = 150
        If hlst.ListCount <= MaxItems Then TempY = 150
        hlst.Scroll = Int((TempY - 16) * (hlst.ListCount - MaxItems) / 134)
        BarritaHechiz.top = TempY

    End If

End Sub

'Private Sub btnHechizos_Click()
'    Call Audio.PlayWave(SND_CLICK)
'    'picInv.Visible = False
'    hlst.Visible = True
'    cmdinfo.Visible = True
'    ' CmdLanzar.Visible = True
'    Coord.Visible = False
'
'    cmdMoverHechi(0).Visible = True
'    cmdMoverHechi(1).Visible = True
'
'    cmdMoverHechi(0).Enabled = True
'    cmdMoverHechi(1).Enabled = True
'
'    'btnInventario.Visible = True
'    'btnHechizos.Visible = False
'
'    'btnHechizos.Picture = LoadPictureEX("btnHechizos_R.jpg")
'    'btnInventario.Picture = LoadPictureEX("btnInventario.jpg")
'    BarraHechiz.Visible = True
'    lblItemInfo.Visible = False
'
'End Sub

Private Sub btnInventario_Click()
    Call Audio.PlayWave(SND_CLICK)
    'picInv.Visible = True

    hlst.Visible = False
    cmdinfo.Visible = True
    'CmdLanzar.Visible = False
    Coord.Visible = True
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    '    cmdMoverHechi(0).Enabled = False
    '    cmdMoverHechi(1).Enabled = False
    
    'btnInventario.Visible = False
    'btnHechizos.Visible = True
    
    'btnInventario.Picture = LoadPictureEX("btnInventario_R.jpg")
    'btnHechizos.Picture = LoadPictureEX("btnHechizos.jpg")
    BarraHechiz.Visible = False
    cmdinfo.Visible = True
    Picture1.Visible = True
    lblItemInfo.Visible = True
    
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)

    If hlst.ListIndex = -1 Then Exit Sub

    Dim Stemp As String

    Select Case Index

        Case 1 'subir

            If hlst.ListIndex = 0 Then Exit Sub

        Case 0 'bajar

            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub

    End Select

    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    
    Select Case Index

        Case 1 'subir
            Stemp = hlst.List(hlst.ListIndex - 1)
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = Stemp
            hlst.ListIndex = hlst.ListIndex - 1

        Case 0 'bajar
            Stemp = hlst.List(hlst.ListIndex + 1)
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = Stemp
            hlst.ListIndex = hlst.ListIndex + 1

    End Select

End Sub

Public Sub DibujarSeguro()
    PicSeg.Visible = True

End Sub

Public Sub DesDibujarSeguro()
    PicSeg.Visible = False

End Sub

Private Sub Command1_Click()

    Hora = Hora + 1
    Call SetDayLight(True)

    'ScreenShooterCapturePending = True

    'GoingHome = 1
    'Dim i As Integer
    'i = Val(Text1.Text)
    'If i > 0 Then
    'Call DrawTextPergamino(Tutoriales(i).Linea1 & vbCrLf & Tutoriales(i).Linea2 & vbCrLf & Tutoriales(i).Linea3, 0, 0)
    'End If
End Sub

Private Sub Command2_Click()
    AlphaRelampago = 150

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case CustomKeys.BindedKey(eKeyType.mKeyVerMapa)

            If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
                VerMapa = True

            End If
   
        Case vbKeyEscape

            If Conectar Then
                If GTCPres < 10000 Then
                    GTCInicial = GTCInicial - (10000 - GTCPres)
                    Call Audio.PlayBackgroundMusic("10", MusicTypes.Mp3)
                
                ElseIf MostrarEntrar > 0 Then

                    If GTCPres - MostrarEntrar > 1000 Then
                        MostrarEntrar = -GTCPres
                        tUser.Visible = False
                        tPass.Visible = False
                        tEmail.Visible = False
                        tRePass.Visible = False
                        TCod.Visible = False
                        Call Audio.PlayWave(SND_CADENAS)
                    
                        If MostrarCrearCuenta = True Then
                            If mOpciones.Recordar = True Then
                                tPass.Text = mOpciones.RecordarPassword
                                tUser.Text = mOpciones.RecordarUsuario

                            End If

                        End If

                    End If
                
                Else
                    prgRun = False
                    Audio.StopMidi

                End If

            End If

    End Select

End Sub

Public Sub SetRender(Full As Boolean)
    #If renderful = 0 Then
    
        'If Full Then
        pRender.Move 0, 0, 1024, 782
        'Else
        ' pRender.Move 2, 125, 800, 608

        'pRender.Move 13, 169, 768, 576
        'pRender.Move 11, 133, Render_Width, Render_Height
        'End If
        
    #Else

        If Full Then
            pRender.Move 0, 0, 1024, 782
        Else
            pRender.Move 2, 125, 800, 608

            'pRender.Move 13, 169, 768, 576
            'pRender.Move 11, 133, Render_Width, Render_Height
        End If

    #End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Conectar Then Exit Sub

    #If SeguridadAlkon Then

        If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
    #End If

    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then

        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

            Select Case KeyCode

            Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                mOpciones.Music = Not mOpciones.Music
                Audio.MusicActivated = mOpciones.Music

            Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                Call AgarrarItem

            Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                Call EquiparItem

            Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                Nombres = Not Nombres



            Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)

                If UserEstado = 1 Then

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                Else
                    Call WriteWork(eSkill.Domar)

                End If

            Case CustomKeys.BindedKey(eKeyType.mKeySteal)

                If UserEstado = 1 Then

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                Else
                    Call WriteWork(eSkill.Robar)

                End If

            Case CustomKeys.BindedKey(eKeyType.mKeyHide)

                If UserEstado = 1 Then

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                Else

                    If MainTimer.Check(TimersIndex.Hide) Then
                        Call WriteWork(eSkill.Ocultarse)

                    End If

                End If

            Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                Call TirarItem

            Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)

                If macrotrabajo.Enabled Then DesactivarMacroTrabajo

                If MainTimer.Check(TimersIndex.PuedeUsar) Then
                    Call UsarItem

                    If InStr(Inventario.ItemName(Inventario.SelectedItem), "Bala") > 0 Then
                        If Inventario.Equipped(Inventario.SelectedItem) Then
                            UsingSecondSkill = 1

                        End If

                    End If

                End If

            Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)

                If UserMoving = 1 Then

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Para actualizar la posición debes estar quieto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                If MainTimer.Check(TimersIndex.SendRPU) And Not UserEmbarcado Then
                    Call WriteRequestPositionUpdate
                    Beep

                End If

            Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                Call WriteResuscitationToggle

            End Select

        Else

            Select Case KeyCode

                'Custom messages!
            Case vbKey0 To vbKey9

                If LenB(CustomMessages.Message((KeyCode - 39) Mod 10)) <> 0 Then
                    Call WriteTalk(CustomMessages.Message((KeyCode - 39) Mod 10))

                End If

            End Select

        End If

    End If

    Select Case KeyCode

    Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)

        If SendTxt.Visible Then Exit Sub

        If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And (Not frmBancoObj.Visible) And (Not frmMSG.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
            SendCMSTXT.Visible = True
            SendCMSTXT.SetFocus

        End If

    Case CustomKeys.BindedKey(eKeyType.mKeyVerMapa)
        VerMapa = False

    Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
        Call ScreenCapture

    Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
        FPSFLAG = Not FPSFLAG

    Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
        Call frmOpciones.Show(vbModeless, frmMain)

    Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)

        If UserMinMAN = UserMaxMAN Then Exit Sub

        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

            End With

            Exit Sub

        End If

        If Not PuedeMacrear Then
            AddtoRichPicture "¡No puedes usar el macro tan rápido!", 255, 255, 255, True, False, False
        ElseIf charlist(UserCharIndex).Moving = 0 Then
            Call WriteMeditate
            PuedeMacrear = False

        End If

    Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)

        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

            End With

            Exit Sub

        End If

        If macrotrabajo.Enabled Then
            DesactivarMacroTrabajo
        Else
            ActivarMacroTrabajo

        End If

    Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)

        If frmMain.macrotrabajo.Enabled Then DesactivarMacroTrabajo
        Call WriteQuit

    Case CustomKeys.BindedKey(eKeyType.mKeyAttack)

        If Shift <> 0 Then Exit Sub

        If Not MainTimer.Check(TimersIndex.PuedeGolpe) Or UserDescansar Or UserMeditar Then Exit Sub

        If macrotrabajo.Enabled Then DesactivarMacroTrabajo
        Call WriteAttack

    Case CustomKeys.BindedKey(eKeyType.mKeyTalk)

        If SendCMSTXT.Visible Then Exit Sub

        If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And (Not frmBancoObj.Visible) And (Not frmMSG.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
            SendTxt.Visible = True
            SendTxt.SetFocus

        End If

    Case CustomKeys.BindedKey(eKeyType.mKeyMontar)

        If SendCMSTXT.Visible Then Exit Sub
        If SendTxt.Visible Then Exit Sub

        If MainTimer.Check(TimersIndex.Montar) Then
            Call WriteEquitar

        End If

    Case CustomKeys.BindedKey(eKeyType.mKeyAnclar)

        If SendCMSTXT.Visible Then Exit Sub
        If SendTxt.Visible Then Exit Sub
        If UserMoving = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Debes detener la embarcación para saltar al agua!!", .red, .green, .blue, .bold, .italic)

            End With

            Exit Sub

        End If

        If MainTimer.Check(TimersIndex.Anclar) Then
            Call WriteAnclarEmbarcacion

        End If

    Case CustomKeys.BindedKey(eKeyType.mKeyPanelGM)

        If SendCMSTXT.Visible Then Exit Sub
        If SendTxt.Visible Then Exit Sub

        frmPanelGm.Show

    End Select

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Y < 24 And NoRes Then
        MoverVentana (Me.hwnd)

    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If UserPasarNivel > 0 Then
        frmMain.exp.Caption = Round((UserExp / UserPasarNivel) * 100, 2) & "%"
        frmMain.exp.Caption = Round((UserExp / UserPasarNivel) * 100, 2) & "%"

    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

End Sub

Private Sub Image2_Click()
    prgRun = False

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Image4_Click()
    Me.WindowState = vbMinimized

End Sub

Private Sub lblEXP_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
    frmMain.exp.Caption = UserExp & "/" & UserPasarNivel

End Sub

Private Sub lblinve_Click()

    Call Audio.PlayWave(SND_CLICKNEW)
     
    If invHechisos.Visible = True Then
        invHechisos.Visible = False
       ' picHechiz.Visible = False
        'CmdLanzar.Visible = False
        BarraHechiz.Visible = False
        LanzarImg.Visible = False
        'picfondoinve.Visible = True
        picInv.Visible = True
        
        Exit Sub
    Else

        If picInv.Visible = True Then
            'picfondoinve.Visible = False
            picInv.Visible = False
            Exit Sub

        End If

        'picfondoinve.Visible = True
        picInv.Visible = True
        Exit Sub

    End If
       
End Sub

Private Sub Lblmagia_Click()

    Call Audio.PlayWave(214)

    If invHechisos.Visible = True Then
        invHechisos.Visible = False
        'picHechiz.Visible = False
        'CmdLanzar.Visible = False
        BarraHechiz.Visible = False
        Picture1.Visible = False
        cmdinfo.Visible = False
        'picfondoinve.Visible = True
        LanzarImg.Visible = False
        
        Exit Sub

    Else
        invHechisos.Visible = True
        'CmdLanzar.Visible = True
        BarraHechiz.Visible = True
        'picHechiz.Visible = True
        Picture1.Visible = True
        cmdinfo.Visible = True
        LanzarImg.Visible = False
        picInv.Visible = False
       
        Exit Sub

    End If

End Sub



Private Sub PicSpells_Click()
  s_Index = invSpells.SelectedItem

        'No hay
        If (s_Index = 0) Then Exit Sub

        hlst.ListIndex = s_Index - 1

        Call cmdLanzar_Click
End Sub

Private Sub PicSpells_DblClick()


       
If invSpells.SelectedItem = 0 Then Exit Sub
    
        Call WriteSpellInfo(invSpells.SelectedItem)
End Sub

Private Sub PicSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)



    If InvX >= Inventario.OffSetX And InvY >= Inventario.OffSetY Then
        Call Audio.PlayWave(SND_CLICK)
    End If

   
    
    
End Sub

Private Sub PicSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
InvX = X
    InvY = Y
 
    
    If Button = 2 And Not Comerciando Then
        If invSpells.GrhIndex(invSpells.SelectedItem) > 0 Then
          
            DragAndDrop = True
            Me.MouseIcon = GetIcon(invSpells.Grafico(GrhData(invSpells.GrhIndex(invSpells.SelectedItem)).FileNum), 0, 0, Halftone, True, RGB(255, 0, 255))
            Me.MousePointer = 99

        End If

    End If
End Sub

Private Sub PicSpells_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Stemp As String
   s_Index = invSpells.SelectedItem
If DragAndDrop Then
        frmMain.MouseIcon = Nothing
        frmMain.MousePointer = 99
        Call SetCursor(General)

    End If

    If Button = 2 And DragAndDrop And Inventario.SelectedItem > 0 And Not Comerciando Then
        If X >= invSpells.OffSetX And Y >= invSpells.OffSetY And X <= PicSpells.Width And Y <= PicSpells.Height Then

            Dim NewPosInv As Integer
              
            NewPosInv = invSpells.ClickItem(X, Y)

            If NewPosInv > 0 Then
              
                
                  Dim NewLugar As Integer
    Dim AntLugar As Integer
      Dim AntText As String
    NewLugar = NewPosInv
    AntLugar = s_Index
   
      Call WriteMoveSpell(AntLugar, NewLugar)
   
   AntText = hlst.List(NewLugar - 1)
    hlst.List(NewLugar - 1) = hlst.List(AntLugar - 1)
    hlst.List(AntLugar - 1) = AntText
   
                Call WriteIntercambiarInv(invSpells.SelectedItem, NewPosInv, False)
                Call invSpells.Intercambiar(NewPosInv)

            End If
    
       End If
End If
           

    DragAndDrop = False
End Sub


Private Sub pRender_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case CustomKeys.BindedKey(eKeyType.mKeyVerMapa)

            If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
                VerMapa = True

            End If
   
        Case vbKeyEscape

            If Conectar Then
                If GTCPres < 10000 Then
                    GTCInicial = GTCInicial - (10000 - GTCPres)
                    Call Audio.PlayBackgroundMusic("10", MusicTypes.Mp3)
                
                ElseIf MostrarEntrar > 0 Then

                    If GTCPres - MostrarEntrar > 1000 Then
                        MostrarEntrar = -GTCPres
                        tUser.Visible = False
                        tPass.Visible = False
                        tEmail.Visible = False
                        tRePass.Visible = False
                        TCod.Visible = False
                        Call Audio.PlayWave(SND_CADENAS)
                    
                        If MostrarCrearCuenta = True Then
                            If mOpciones.Recordar = True Then
                                tPass.Text = mOpciones.RecordarPassword
                                tUser.Text = mOpciones.RecordarUsuario

                            End If

                        End If

                    End If
                
                Else
                    prgRun = False
                    Audio.StopMidi

                End If

            End If

    End Select

End Sub

Private Sub pRender_KeyUp(KeyCode As Integer, Shift As Integer)

    If Conectar Then Exit Sub

    #If SeguridadAlkon Then

        If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
    #End If
    
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

            Select Case KeyCode

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    mOpciones.Music = Not mOpciones.Music
                    Audio.MusicActivated = mOpciones.Music
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem

                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
               
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                        End With

                    Else
                        Call WriteWork(eSkill.Domar)

                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                        End With

                    Else
                        Call WriteWork(eSkill.Robar)

                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)

                    If UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                        End With

                    Else

                        If MainTimer.Check(TimersIndex.Hide) Then
                            Call WriteWork(eSkill.Ocultarse)

                        End If

                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)

                    If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.PuedeUsar) Then
                        Call UsarItem
                        
                        If InStr(Inventario.ItemName(Inventario.SelectedItem), "Bala") > 0 Then
                            If Inventario.Equipped(Inventario.SelectedItem) Then
                                UsingSecondSkill = 1

                            End If

                        End If
                        
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)

                    If UserMoving = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Para actualizar la posición debes estar quieto!!", .red, .green, .blue, .bold, .italic)

                        End With

                        Exit Sub

                    End If

                    If MainTimer.Check(TimersIndex.SendRPU) And Not UserEmbarcado Then
                        Call WriteRequestPositionUpdate
                        Beep

                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle

            End Select

        Else

            Select Case KeyCode

                    'Custom messages!
                Case vbKey0 To vbKey9

                    If LenB(CustomMessages.Message((KeyCode - 39) Mod 10)) <> 0 Then
                        Call WriteTalk(CustomMessages.Message((KeyCode - 39) Mod 10))

                    End If

            End Select

        End If

    End If
    
    Select Case KeyCode

        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)

            If SendTxt.Visible Then Exit Sub
            
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And (Not frmBancoObj.Visible) And (Not frmMSG.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendCMSTXT.Visible = True
                SendCMSTXT.SetFocus

            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyVerMapa)
            VerMapa = False

        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture
        
        Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
            'prue
            '      FPSFLAG = Not FPSFLAG
            
        Case CustomKeys.BindedKey(eKeyType.mKeySalir)
            prgRun = False
             
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)

            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                End With

                Exit Sub

            End If
                
            If Not PuedeMacrear Then
                AddtoRichPicture "¡No puedes usar el macro tan rápido!", 255, 255, 255, True, False, False
            ElseIf charlist(UserCharIndex).Moving = 0 Then
                Call WriteMeditate
                PuedeMacrear = False

            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)

            If UserEstado = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                End With

                Exit Sub

            End If
            
            If macrotrabajo.Enabled Then
                DesactivarMacroTrabajo
            Else
                ActivarMacroTrabajo

            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)

            If frmMain.macrotrabajo.Enabled Then DesactivarMacroTrabajo
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)

            If Shift <> 0 Then Exit Sub
         
            If Not MainTimer.Check(TimersIndex.PuedeGolpe) Or UserDescansar Or UserMeditar Then Exit Sub
            
            If macrotrabajo.Enabled Then DesactivarMacroTrabajo
            Call WriteAttack
                    
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)

            If SendCMSTXT.Visible Then Exit Sub
            
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And (Not frmBancoObj.Visible) And (Not frmMSG.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus

            End If
            
        Case CustomKeys.BindedKey(eKeyType.mKeyMontar)

            If SendCMSTXT.Visible Then Exit Sub
            If SendTxt.Visible Then Exit Sub
        
            If MainTimer.Check(TimersIndex.Montar) Then
                Call WriteEquitar

            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyAnclar)

            If SendCMSTXT.Visible Then Exit Sub
            If SendTxt.Visible Then Exit Sub
            If UserMoving = 1 Then

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Debes detener la embarcación para saltar al agua!!", .red, .green, .blue, .bold, .italic)

                End With

                Exit Sub

            End If
            
            If MainTimer.Check(TimersIndex.Anclar) Then
                Call WriteAnclarEmbarcacion

            End If
            
        Case CustomKeys.BindedKey(eKeyType.mKeyPanelGM)

            If SendCMSTXT.Visible Then Exit Sub
            If SendTxt.Visible Then Exit Sub
        
            frmPanelGm.Show
            
    End Select

End Sub

'Private Sub QuestBoton_MouseMove(Button As Integer, _
'                                 Shift As Integer, _
'                                 X As Single, _
'                                 Y As Single)
'
'    On Error GoTo QuestBoton_MouseMove_Err
'
'    If QuestBoton.Tag = "0" Then
'        'QuestBoton.Picture = LoadInterface("questover.bmp")
'        QuestBoton.Tag = "1"
'
'    End If
'
'    Exit Sub
'
'QuestBoton_MouseMove_Err:
'
'    'Call RegistrarError(Err.Number, Err.Description, "frmMain.QuestBoton_MouseMove", Erl)
'    Resume Next
'
'End Sub

Private Sub QuestBoton_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    On Error GoTo QuestBoton_MouseUp_Err
    
    If pausa Then Exit Sub
    
    Call WriteQuestListRequest
    
    Exit Sub

QuestBoton_MouseUp_Err:

    ' Call RegistrarError(Err.Number, Err.Description, "frmMain.QuestBoton_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Macro_Timer()
    PuedeMacrear = True

End Sub

Private Sub macrotrabajo_Timer()

    If Inventario.SelectedItem = 0 Then
        DesactivarMacroTrabajo
        Exit Sub

    End If
    
    'Macros are disabled if not using Argentum!
    'If Not Api.IsAppActive() Then
    '    DesactivarMacroTrabajo
    '    Exit Sub
    'End If
    
    If (UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal Or UsingSkill = eSkill.Herreria) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0

    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
    Call UsarItem

End Sub

Public Sub ControlSeguroResu(ByVal Mostrar As Boolean)

    If Mostrar Then
        
        SeguroResu = 14812
        'If Not PicResu.Visible Then
        '    PicResu.Visible = True
        'End If
    Else

        'If PicResu.Visible Then
        '    PicResu.Visible = False
        'End If
        SeguroResu = 14811
    End If

End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    Call AddtoRichPicture(">>MACRO TRABAJO ACTIVADO<<", 0, 200, 200, True, True, False)

End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    Call AddtoRichPicture(">>MACRO TRABAJO DESACTIVADO<<", 0, 200, 200, True, True, False)

End Sub

Public Sub mnuEquipar_Click()
    Call EquiparItem

End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart

End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)

End Sub

Private Sub mnuTirar_Click()
    Call TirarItem

End Sub

Private Sub mnuUsar_Click()
    Call UsarItem

End Sub

Private Sub MenuF_Click()
    'Menu.Visible = True
    'Helios Barras
'    bar_salud(0).Visible = False
'    Bar_Mana(0).Visible = False
'
'    bar_sta.Visible = False
'    bar_comida.Visible = False
'    Bar_Agua.Visible = False

End Sub

Private Sub lbStats_Click()
    'Menu.Visible = False
    'Helios Barras
'    bar_salud(0).Visible = True
'    Bar_Mana(0).Visible = True
'    bar_sta.Visible = True
'    bar_comida.Visible = True
'    Bar_Agua.Visible = True

End Sub

Private Sub picHechiz_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    Call Audio.PlayWave(SND_CLICK)

    If Y < 0 Then Y = 0
    If Y > 228 Then Y = 228
    hlst.ListIndex = Int(Y / hlst.Pixel_Alto) + hlst.Scroll

End Sub

Private Sub picHechiz_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    If Button = 1 Then
        If Y < 0 Then Y = 0
        If Y > 228 Then Y = 228
        hlst.ListIndex = Int(Y / hlst.Pixel_Alto) + hlst.Scroll

    End If

End Sub

Private Sub picInv_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    If InvX >= Inventario.OffSetX And InvY >= Inventario.OffSetY Then
        Call Audio.PlayWave(SND_CLICK)

    End If

End Sub

Private Sub PicInv_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    InvX = X
    InvY = Y

    If Button = 2 And Not Comerciando Then
        If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then
            DragAndDrop = True
            Me.MouseIcon = GetIcon(Inventario.Grafico(GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum), 0, 0, Halftone, True, RGB(255, 0, 255))
            Me.MousePointer = 99

        End If

    End If

End Sub

Private Sub picResu_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteResuscitationToggle

End Sub

Private Sub PicSeg_Click()
    Call Audio.PlayWave(SND_CLICK)
    'AddtoRichPicture "El dibujo de la llave indica que tienes activado el seguro, esto evitará que por accidente ataques a un ciudadano y te conviertas en criminal. Para activarlo o desactivarlo utiliza el comando /SEG", 255, 255, 255, False, False, False
    Call WriteSafeToggle

End Sub

Private Sub Coord_Click()
    AddtoRichPicture "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False

End Sub

Private Function InGameArea() As Boolean

    '***************************************************
    'Author: NicoNZ
    'Last Modification: 04/07/08
    'Checks if last click was performed within or outside the game area.
    '***************************************************
    If clicX < 0 Or clicX > pRender.Width Then Exit Function
    If clicY < 0 Or clicY > pRender.Height Then Exit Function
    
    InGameArea = True

End Function

Private Sub pRender_Click()

    If Conectar Then Exit Sub
    If Cartel Then Cartel = False



    If MouseX > 21 And MouseX < 46 And MouseY > 737 And MouseY < 765 Then    'mensajes

        Call Audio.PlayWave(SND_CLICK)
        If FrmMensajes.Visible = True Then
            Unload FrmMensajes
            Consolacom = 14816
            Exit Sub
        Else
            Call FrmMensajes.Show(vbModeless, frmMain)
            Consolacom = 14817
            Exit Sub
        End If
    End If


    If MouseX > 551 And MouseX < 622 And MouseY > 3 And MouseY < 33 Then    'Online

        Call Audio.PlayWave(SND_CLICK)
        Call WriteOnline
        Exit Sub
    End If

    If MouseX > 952 And MouseX < 976 And MouseY > 737 And MouseY < 765 Then    'Seguro Resu

        Call Audio.PlayWave(SND_CLICK)
        Call WriteResuscitationToggle
        Exit Sub
    End If

    If MouseX > 981 And MouseX < 1005 And MouseY > 737 And MouseY < 765 Then    'Seguro Combate

        Call Audio.PlayWave(SND_CLICK)
        Call WriteSafeToggle

        Exit Sub
    End If



    If MouseX > 113 And MouseX < 310 And MouseY > 49 And MouseY < 67 Then
        Call Audio.PlayWave(SND_CLICKNEW)

        If UserMinMAN = UserMaxMAN Or charlist(UserCharIndex).Moving Then Exit Sub

        If UserEstado = 1 Then    'Muerto

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

            End With

            Exit Sub

        End If

        Call WriteMeditate
        Exit Sub

    End If


    If MouseX > 113 And MouseX < 310 And MouseY > 19 And MouseY < 36 Then
        Call Audio.PlayWave(SND_CLICKNEW)
        If Vidarender = True Then
            Vidarender = False
            Exit Sub
        Else
            Vidarender = True
            Exit Sub
        End If
    End If

    If MouseX > 633 And MouseX < 658 And MouseY > 6 And MouseY < 27 Then    'Ranking
        Call ImgLanzar_Click(4)
        Call Audio.PlayWave(SND_CLICKNEW)


        Exit Sub

    End If



    'Helios 03/06/2021 "coordenadas del raton"
    If MouseX > 661 And MouseX < 681 And MouseY > 5 And MouseY < 25 Then
        Call ImgLanzar_Click(2)    'Helios Clanes 04/06/2021
        Call Audio.PlayWave(SND_CLICKNEW)
        Exit Sub

    End If


    If MouseX > 13 And MouseX < 40 And MouseY > 120 And MouseY < 144 Then    'oro

        Call Audio.PlayWave(SND_CLICKNEW)
        Inventario.SelectGold

        If UserGLD > 0 Then
            frmCantidad.Show , frmMain

        End If
    End If

    If MouseX > 688 And MouseX < 712 And MouseY > 5 And MouseY < 25 Then
        Call ImgLanzar_Click(1)    ' Estadisticas
        Call Audio.PlayWave(SND_CLICKNEW)
        Exit Sub

    End If

    'Mostrar el Minipa helios 06/06/2021
    If MouseX > 719 And MouseX < 739 And MouseY > 5 And MouseY < 25 Then
        Call Audio.PlayWave(SND_CLICKNEW)

        If frmMain.imgMiniMapa.Visible = True Then
            frmMain.imgMiniMapa.Visible = False

        Else

            frmMain.imgMiniMapa.Visible = True

        End If

        Exit Sub

    End If

    'Helios Misiones 04/06/2021
    If MouseX > 749 And MouseX < 764 And MouseY > 5 And MouseY < 25 Then
        Call Audio.PlayWave(SND_CLICKNEW)

        If pausa Then Exit Sub

        'helios 06/06/2021
        If FrmQuests.Visible = True Then Unload FrmQuests: Exit Sub
        Call WriteQuestListRequest

        Exit Sub

    End If

    If MouseX > 774 And MouseX < 794 And MouseY > 5 And MouseY < 25 Then
        Call Audio.PlayWave(SND_CLICKNEW)

        'helios 06/06/2021
        If frmParty.Visible = True Then Unload frmParty: Exit Sub
        Call ImgLanzar_Click(3)    ' Helios PArty 04/06/2021
        Exit Sub

    End If

    If MouseX > 803 And MouseX < 824 And MouseY > 5 And MouseY < 25 Then
        Call ImgLanzar_Click(0)    ' Helios Opciones 04/06/2021
        Call Audio.PlayWave(SND_CLICKNEW)
        Exit Sub

    End If

    If MouseX > 854 And MouseX < 878 And MouseY > 5 And MouseY < 25 Then
        CTextos = CTextos + 1
        Call Audio.PlayWave(SND_CLICK)
        If CTextos = 1 Then
            sintextos = False

            Dim a As Byte

            For a = 1 To 6
                Call ShowConsoleMsg(" ", , , , , 0)
            Next a

        Else

            sintextos = True
            CTextos = 0

        End If

    End If

    'helios esconder Barras 07/06/2021
    If MouseX > 882 And MouseX < 903 And MouseY > 5 And MouseY < 25 Then
        'PulsarEsconder = PulsarEsconder + 1
        Call Audio.PlayWave(SND_CLICKNEW)

        If MmenuBarras = True Then
          MmenuBarras = False
            Exit Sub
        Else
           MmenuBarras = True
           Exit Sub
        End If




    End If

    If MouseX > 1003 And MouseX < 1018 And MouseY > 6 And MouseY < 19 Then
        'helios 06/06/2021
        Call Audio.PlayWave(SND_CLICKNEW)    ' Desconectar

        If frmCerrar.Visible Then Exit Sub

        Dim mForm As Form

        For Each mForm In Forms

            If mForm.hwnd <> Me.hwnd Then Unload mForm
            Set mForm = Nothing
        Next
        frmCerrar.Show , Me

        'Call WriteQuit
        Exit Sub

    End If

    If MouseX > 828 And MouseX < 850 And MouseY > 5 And MouseY < 25 Then



        ContarClip = ContarClip + 1

        If ContarClip = 1 Then

            Call Audio.PlayWave(SND_CLICKNEW)
            'invHechisos.Visible = True
            'picHechiz.Visible = True
            'CmdLanzar.Visible = False
            'BarraHechiz.Visible = True
            'LanzarImg.Visible = True
            'picfondoinve.Visible = True
            picInv.Visible = True
            PicSpells.Visible = True
            MostrarMenuInventario = True

            barritaa.Visible = True

        Else
            PicSpells.Visible = False
            invHechisos.Visible = False
            'picHechiz.Visible = False
            'CmdLanzar.Visible = False
            BarraHechiz.Visible = False
            LanzarImg.Visible = False
            'picfondoinve.Visible = True
            picInv.Visible = False
            MostrarMenuInventario = False
            barritaa.Visible = False
            ContarClip = 0

        End If

        Exit Sub

    End If

    'helios 06/06/2021
    If MouseX > 902 And MouseX < 925 And MouseY > 182 And MouseY < 1237 Then
        If frmMain.invHechisos.Visible = True Then
            Call Audio.PlayWave(SND_CLICKNEW)

            Call LanzarImg_Click    ' Lanzar Magia

        End If

        Exit Sub

    End If

    'helios 06/06/2021

    If ContarClip = 1 Then

        If MouseX > 977 And MouseX < 989 And MouseY > 230 And MouseY < 242 Then
            Call Audio.PlayWave(SND_CLICK)

            If picInv.left = 978 Then

                picInv.left = 946
                Exit Sub

            End If

            If picInv.left = 946 Then
                picInv.left = 914
                Exit Sub
            End If

            If picInv.left = 914 Then
                picInv.left = 882
                Exit Sub
            End If

        End If

        If MouseX > 993 And MouseX < 1007 And MouseY > 228 And MouseY < 243 Then

            Call Audio.PlayWave(SND_CLICK)

            If picInv.left = 882 Then

                picInv.left = 914

                Exit Sub
            End If

            If picInv.left = 914 Then
                picInv.left = 946
                Exit Sub
            End If

            If picInv.left = 946 Then
                picInv.left = 978
                Exit Sub
            End If

        End If

    End If

    If ContarClip = 1 Then

        If MouseX > 977 And MouseX < 989 And MouseY > 568 And MouseY < 581 Then
            Call Audio.PlayWave(SND_CLICK)

            If PicSpells.left = 978 Then

                PicSpells.left = 946
                Exit Sub

            End If

            If PicSpells.left = 946 Then
                PicSpells.left = 914
                Exit Sub
            End If

            If PicSpells.left = 914 Then
                PicSpells.left = 882
                Exit Sub
            End If

        End If

        If MouseX > 993 And MouseX < 1007 And MouseY > 567 And MouseY < 579 Then

            Call Audio.PlayWave(SND_CLICK)

            If PicSpells.left = 882 Then

                PicSpells.left = 914

                Exit Sub
            End If

            If PicSpells.left = 914 Then
                PicSpells.left = 946
                Exit Sub
            End If

            If PicSpells.left = 946 Then
                PicSpells.left = 978
                Exit Sub
            End If

        End If

    End If

    'helios 06/06/2021

    'Fin Helios 03/06/2021 "coordenadas del raton"

    If UserEmbarcado Then
        If Not Barco(0) Is Nothing Then
            If Barco(0).TickPuerto = 0 And Barco(0).Embarcado = True Then
                Exit Sub

            End If

        End If

        If Not Barco(1) Is Nothing Then
            If Barco(1).TickPuerto = 0 And Barco(1).Embarcado = True Then
                Exit Sub

            End If

        End If

    End If

    #If SeguridadAlkon Then

        If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
    #End If

    If Not Comerciando Then

        If SendTxt.Visible = True Then
            SendTxt.SetFocus
        ElseIf SendCMSTXT.Visible = True Then
            SendCMSTXT.SetFocus

        End If

        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        Debug.Print "Coordenada X: " & MouseX & " Coordenada Y: " & MouseY & " --- " & tX & " - " & tY

        If Not InGameArea() Then Exit Sub
        If Not InMapBounds(tX, tY) Then Exit Sub

        If MouseShift = 0 And (MapData(tX, tY).Graphic(4).GrhIndex = 0 Or bTecho) Then

            'If MouseBoton <> vbRightButton Then
            If MouseBoton Then
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else

                    If macrotrabajo.Enabled Then DesactivarMacroTrabajo

                    '         If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                    '            'frmMain.MousePointer = vbDefault
                    '           Call SetCursor(General)
                    '          UsingSkill = 0
                    '         With FontTypes(FontTypeNames.FONTTYPE_TALK)
                    '            Call AddtoRichPicture("No podés lanzar proyectiles tan rapido.", .red, .green, .blue, .bold, .italic)
                    '       End With
                    '      Exit Sub
                    '  End If

                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.PuedeFlechas) Then
                            frmMain.MousePointer = vbDefault
                            Call SetCursor(General)
                            UsingSkill = 0

                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichPicture("No podés lanzar proyectiles tan rapido.", .red, .green, .blue, .bold, .italic)

                            End With

                            Exit Sub

                        End If

                    End If

                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.PuedeLanzarHechizo) Then  'Corto intervalo de Golpe-Magia
                            frmMain.MousePointer = vbDefault
                            Call SetCursor(General)
                            UsingSkill = 0

                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichPicture("No puedes lanzar hechizos tan rápido.", .red, .green, .blue, .bold, .italic)

                            End With

                            Exit Sub

                        End If

                    End If

                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            Call SetCursor(General)
                            UsingSkill = 0
                            Exit Sub

                        End If

                    End If

                    'If frmMain.MousePointer <> 99 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)

                    frmMain.MousePointer = vbDefault

                    If UsingSecondSkill = 1 Then
                        If UserNavegando = True Then
                            Call WriteCreateEfectoClient(tX, tY, eEffects.Bala)
                            UsingSecondSkill = 0
                        Else

                            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                Call AddtoRichPicture("Debes estar navegando para usar el cañon.", .red, .green, .blue, .bold, .italic)

                            End With

                        End If

                    Else
                        Call WriteWorkLeftClick(tX, tY, UsingSkill)

                    End If

                    frmMain.MousePointer = vbDefault
                    Call SetCursor(General)

                    UsingSkill = 0
                    UsingSecondSkill = 0

                End If

            Else
                Call AbrirMenuViewPort

            End If

        ElseIf (MouseShift And 1) = 1 Then

            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MainTimer.Check(TimersIndex.Telep) And Not UserEmbarcado And Not UserNadando And UserMoving = 0 Then
                    If MouseBoton = vbLeftButton And charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 5 Then
                        If VerMapa Then
                            Call WriteWarpMeToTarget(Int((frmMain.MouseX - PosMapX + 32) / RelacionMiniMapa), Int((frmMain.MouseY - PosMapY + 32) / RelacionMiniMapa))
                        Else
                            Call WriteWarpMeToTarget(tX, tY)

                        End If

                    End If

                End If

            End If

        End If

    End If

End Sub

Private Sub pRender_DblClick()

    If Conectar Then Exit Sub
    Call WriteDoubleClick(tX, tY)
    If MouseX > 113 And MouseX < 310 And MouseY > 49 And MouseY < 67 Then
        Call Audio.PlayWave(SND_CLICKNEW)
        If Manarender = True Then
            Manarender = False
            Exit Sub
        Else
            Manarender = True
            Exit Sub
        End If
    End If

End Sub

Private Sub pRender_MouseDown(Button As Integer, _
                              Shift As Integer, _
                            X As Single, _
                            Y As Single)

    If MouseX > 292 And MouseX < 543 And MouseY > 0 And MouseY < 12 Then    'mover pantalla
        If Resolucion = False Then
            If Not Conectar Then
                Call ReleaseCapture
                
                Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
                 
            End If
        End If
    End If


    MouseBoton = Button
    MouseShift = Shift



End Sub

Private Sub pRender_MouseMove(Button As Integer, _
                              Shift As Integer, _
                            X As Single, _
                            Y As Single)
    Dim f As New StdFont
    MouseX = X
    MouseY = Y

    'Recuerdo Flechas inventario
    If MouseX > 977 And MouseX < 989 And MouseY > 230 And MouseY < 242 Then
        If MostrarMenuInventario = True Then
            RecuadroInv = True
            RecuadroX = 975
            RecuadroY = 230
            Exit Sub
        End If
    End If

    If MouseX > 993 And MouseX < 1007 And MouseY > 228 And MouseY < 243 Then
        If MostrarMenuInventario = True Then
            RecuadroInv = True
            RecuadroX = 993
            RecuadroY = 230

            Exit Sub
        End If
    End If



    If MouseX > 977 And MouseX < 989 And MouseY > 568 And MouseY < 581 Then
        If MostrarMenuInventario = True Then
            RecuadroInv = True
            RecuadroX = 975
            RecuadroY = 567
            Exit Sub
        End If
    End If

    If MouseX > 993 And MouseX < 1007 And MouseY > 567 And MouseY < 579 Then
        If MostrarMenuInventario = True Then
            RecuadroInv = True
            RecuadroX = 993
            RecuadroY = 567
            Exit Sub
        End If
    End If

    'Recuadro flechas inventario



    If MouseX > 292 And MouseX < 543 And MouseY > 0 And MouseY < 12 Then    'mover pantalla
        If Resolucion = False Then
            If Not Conectar Then

                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Mover Pantalla"
                    TT2.TipText = "Si estas en modo Ventana y mantienes apretado el Click isquierdo del Mouse se movera la pantalla"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)


                    f.Name = "Augusta"
                    f.Size = 12
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 827, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub
    End If








    If MouseX > 828 And MouseX < 850 And MouseY > 5 And MouseY < 25 Then    'inventario
        RecuadroX = 827
        RecuadroY = 6
        RecuadroON = True
        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Inventario"
                    TT2.TipText = "Inventario AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)


                    f.Name = "Augusta"
                    f.Size = 12
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 827, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub
    End If
    If MouseX > 882 And MouseX < 903 And MouseY > 5 And MouseY < 25 Then    ' barras de vida
        RecuadroX = 881
        RecuadroY = 6
        RecuadroON = True
        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Barras de Vida"
                    TT2.TipText = "Vida y Mana AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)

                    f.Size = 12
                    f.Name = "Augusta"
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 882, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub
    End If

    If MouseX > 661 And MouseX < 681 And MouseY > 5 And MouseY < 25 Then    'clanes
        RecuadroX = 662
        RecuadroY = 6
        RecuadroON = True
        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Clanes"
                    TT2.TipText = "Clanes AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)
                    f.Size = 12
                    f.Name = "Augusta"
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 662, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub

    End If


    If MouseX > 719 And MouseX < 739 And MouseY > 5 And MouseY < 25 Then    ' mini mapa
        RecuadroX = 717
        RecuadroY = 6
        RecuadroON = True
        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Mini Mapa"
                    TT2.TipText = "Inventario AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)

                    f.Size = 12
                    f.Name = "Augusta"
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 717, 0, Me.pRender.hwnd
                End If
            End If
        End If

        Exit Sub

    End If

    If MouseX > 803 And MouseX < 824 And MouseY > 5 And MouseY < 25 Then    'opciones

        RecuadroX = 800
        RecuadroY = 6
        RecuadroON = True
        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Opciones"
                    TT2.TipText = "Opciones AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)

                    f.Size = 12
                    f.Name = "Augusta"
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 800, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub
    End If

    If MouseX > 688 And MouseX < 712 And MouseY > 5 And MouseY < 25 Then    'estadisticas
        RecuadroX = 689
        RecuadroY = 6
        RecuadroON = True
        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Estadisticas"
                    TT2.TipText = "Estadisticas AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)

                    f.Size = 12
                    f.Name = "Augusta"
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 689, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub
    End If

    If MouseX > 774 And MouseX < 794 And MouseY > 5 And MouseY < 25 Then    'partis

        RecuadroX = 772
        RecuadroY = 6
        RecuadroON = True
        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Partis"
                    TT2.TipText = "Partis AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)

                    f.Size = 12
                    f.Name = "Augusta"
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 772, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub
    End If


    If MouseX > 749 And MouseX < 764 And MouseY > 5 And MouseY < 25 Then    ' quees
        RecuadroX = 744
        RecuadroY = 6
        RecuadroON = True

        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Misiones"
                    TT2.TipText = "Misiones AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)
                    f.Size = 12
                    f.Name = "Augusta"
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 744, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub

    End If

    If MouseX > 854 And MouseX < 878 And MouseY > 5 And MouseY < 25 Then    'consola
        RecuadroX = 854
        RecuadroY = 6
        RecuadroON = True
        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Consola"
                    TT2.TipText = "Consola AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)
                    f.Size = 12
                    f.Name = "Augusta"
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 854, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub
    End If

    If MouseX > 633 And MouseX < 658 And MouseY > 6 And MouseY < 27 Then   'Ranking
        RecuadroX = 634
        RecuadroY = 6
        RecuadroON = True

        If Not Conectar Then
            If mOpciones.MostrarAyuda Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Ranking"
                    TT2.TipText = "Ranking AOYind"
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)

                    f.Size = 12
                    f.Name = "Augusta"
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 634, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub
    End If


    If MouseX > 52 And MouseX < 83 And MouseY > 60 And MouseY < 75 Then    'exp pasar lvl

        If Not Conectar Then
            If MmenuBarras Then
                contarr = contarr + 1
                If contarr = 1 Then
                    TT2.Style = TTBalloon
                    TT2.Icon = TTIconInfo
                    TT2.Title = "Exp para pasar Nivel "
                    TT2.TipText = UserExp & "/" & UserPasarNivel
                    TT2.PopupOnDemand = False
                    TT2.ForeColor = vbWhite
                    TT2.BackColor = RGB(13, 13, 13)


                    f.Name = "Augusta"
                    f.Size = 12
                    f.Underline = False
                    TT2.TipFont = f


                    TT2.CreateToolTip pRender.hwnd
                    TT2.Show 827, 0, Me.pRender.hwnd
                End If
            End If
        End If
        Exit Sub
    End If

    If MouseX > 1003 And MouseX < 1018 And MouseY > 6 And MouseY < 19 Then
        RecuadroX = 998
        RecuadroY = 2
        RecuadroSON = True
        Exit Sub

    End If
    contarr = 0
    TT2.Destroy
    RecuadroON = False
    RecuadroSON = False
    RecuadroInv = False
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > pRender.Width Then
        MouseX = pRender.Width

    End If

    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > pRender.Height Then
        MouseY = pRender.Height

    End If

    If Conectar Then Call MouseAction(X, Y, 0)

End Sub

Private Sub pRender_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y

    If Conectar Then Call MouseAction(X, Y, 1)

End Sub



Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False

    End If

End Sub

Private Sub SpoofCheck_Timer()

    Dim IPMMSB As Byte

    Dim IPMSB  As Byte

    Dim IPLSB  As Byte

    Dim IPLLSB As Byte

    IPLSB = 3 + 15
    IPMSB = 32 + 15
    IPMMSB = 200 + 15
    IPLLSB = 74 + 15

    If IpServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) & "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()

    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer

End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()

    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            Call WriteDrop(Inventario.SelectedItem, 1)
        Else

            If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                Inventario.DropX = 0
                Inventario.DropY = 0
                frmCantidad.Show , frmMain

            End If

        End If

    End If

End Sub

Private Sub AgarrarItem()
    Call WritePickUp

End Sub

Public Sub UsarItem()

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then Call WriteUseItem(Inventario.SelectedItem)

End Sub

Public Sub EquiparItem()

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then Call WriteEquipItem(Inventario.SelectedItem)

End Sub

Private Sub LanzarImg_Click()

    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.HabilitaLanzarHechizo) Then
        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

            End With

        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True

        End If

    End If

End Sub

Private Sub LanzarImg_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    UsaMacro = False
    CnTd = 0

End Sub

Private Sub cmdINFO_Click()

    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)

    End If

End Sub

Public Sub ReDrawConsola()
    pConsola.Cls

    Dim I As Long

    For I = OffSetConsola To OffSetConsola + 6

        If I >= 0 And I <= LineasConsola Then
            pConsola.CurrentX = 0
            pConsola.CurrentY = (I - OffSetConsola - 1) * 14
            pConsola.ForeColor = Consola(I).Color
            pConsola.FontBold = CBool(Consola(I).bold)
            pConsola.FontItalic = CBool(Consola(I).italic)
            pConsola.Print Consola(I).Texto
            
        End If

    Next I

End Sub

Private Sub Form_Load()
    
    'frmMain.Caption = "AoYind 3"
    'PanelDer.Picture = LoadPicture(App.path & _
     "\Graficos\Principalnuevo_sin_energia.jpg")
    
    'InvEqu.Picture = LoadPicture(App.path & _
     "\Graficos\Centronuevoinventario.jpg")
   
    'frmMain.iBEXP.Picture = LoadPictureEX("BARRAEXP.jpg")
    
    'Me.Picture = LoadPictureEX("VentanaPrincipalm.bmp")
    'CmdLanzar.Picture = LoadPictureEX("Lanzarbtn.jpg")
    ' CmdLanzar.Picture = LoadPictureEX("btnLanzar.jpg")
    'picInv.Picture = LoadPictureEX("VentanaPrincipalInv.jpg")
    
    'BarraHechiz2.Picture = LoadPictureEX("BarraHechiz2.jpg")
    
    'btnHechizos.Picture = LoadPictureEX("btnHechizos_R.bmp")
    'btnInventario.Picture = LoadPictureEX("btnInventario_R.bmp")
    'neo parcheo ruta
    imgMiniMapa.Picture = LoadPicture(PathRecursosCliente & "\Recursos\minimapadefault.bmp")
    ' imgMiniMapa.Picture = LoadPicture(App.path & "\Recursos\minimapadefault.bmp")
    'MANShp.Picture = LoadPictureEX("barMana.jpg")
    'Hpshp.Picture = LoadPictureEX("BarHp.jpg")
    'STAShp.Picture = LoadPictureEX("BarSta.jpg")
    
    'picInv.Height = 246
    
    'btnInventario.MouseIcon = CmdLanzar.MouseIcon
    'btnHechizos.MouseIcon = CmdLanzar.MouseIcon
   
    Set hlst = New clsGraphicalList
    Call hlst.Initialize(Me.picHechiz, RGB(200, 190, 190))
    
    tUser.BackColor = RGB(200, 200, 200)
    tPass.BackColor = RGB(200, 200, 200)
    tEmail.BackColor = RGB(200, 200, 200)
    tRePass.BackColor = RGB(200, 200, 200)
    TCod.BackColor = RGB(200, 200, 200)
    
    tUser.top = 612
    tPass.top = 637
    tPass.left = 428
    tUser.left = 428
    
    Me.left = 0
    Me.top = 0
   
    If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And (Not frmBancoObj.Visible) And (Not frmMSG.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
        Debug.Print "Precarga"

    End If

    OpcionesPath = App.path & "\init\Config.ini"
    ActivarAuras = GetVar(OpcionesPath, "AURAS", "AuraActiva")
    RotarActivado = GetVar(OpcionesPath, "AURAS", "rotacion")

    If ActivarAuras = "1" Then
        frmOpciones.ActAura.value = vbChecked
    Else
        frmOpciones.ActAura.value = vbUnchecked

    End If
   
    If RotarActivado = "1" Then
        frmOpciones.RotaAura.value = vbChecked
    Else
        frmOpciones.RotaAura.value = vbUnchecked

    End If
    
      
    

End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0

End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

Private Sub ImgLanzar_Click(Index As Integer)
'Call Audio.PlayWave(SND_CLICK)

    Select Case Index

    Case 0

        'helios 06/06/2021
        If frmOpciones.Visible = True Then Unload frmOpciones: Exit Sub
        Call frmOpciones.Show(vbModeless, frmMain)

    Case 1

        'helios 06/06/2021
        If frmEstadisticas.Visible = False Then

            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            Call WriteRequestAtributes
            Call WriteRequestSkills
            Call WriteRequestMiniStats
            Call WriteRequestFame
            Call FlushBuffer

            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents    'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Else
            Unload frmEstadisticas

        End If

    Case 2

        'helios 06/06/2021
        If frmGuildAdm.Visible = False Then
            If frmGuildLeader.Visible Then Unload frmGuildLeader

            Call WriteRequestGuildLeaderInfo
        Else
            Unload frmGuildAdm

        End If

    Case 3
        Call WriteRequestPartyForm

    Case 4
        If FrmRanking.Visible = True Then Unload FrmRanking: Exit Sub
        Call FrmRanking.Show(vbModeless, frmMain)

    End Select

End Sub

Private Sub Image3_Click(Index As Integer)

    Select Case Index

        Case 0
            Inventario.SelectGold

            If UserGLD > 0 Then
                frmCantidad.Show , frmMain

            End If

    End Select

End Sub

Private Sub picInv_DblClick()

    If InvX >= Inventario.OffSetX And InvY >= Inventario.OffSetY Then
        If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

        If Not MainTimer.Check(TimersIndex.PuedeUsarDobleClick) Then Exit Sub

        If macrotrabajo.Enabled Then DesactivarMacroTrabajo

        Select Case Inventario.ObjType(Inventario.SelectedItem)

        Case eObjType.otcasco
            Call EquiparItem

        Case eObjType.otArmadura
            Call EquiparItem

        Case eObjType.otescudo
            Call EquiparItem

        Case eObjType.otAlas
            Call EquiparItem

        Case eObjType.otWeapon

            If InStr(Inventario.ItemName(Inventario.SelectedItem), "Arco") > 0 Then
                If Inventario.Equipped(Inventario.SelectedItem) Then
                    Call UsarItem
                Else
                    Call EquiparItem

                End If

            ElseIf InStr(Inventario.ItemName(Inventario.SelectedItem), "Bala") > 0 Then

                If Inventario.Equipped(Inventario.SelectedItem) Then
                    Call UsarItem
                    UsingSecondSkill = 1
                Else
                    Call EquiparItem

                End If
            ElseIf InStr(Inventario.ItemName(Inventario.SelectedItem), "Caña") > 0 _
                   Or InStr(Inventario.ItemName(Inventario.SelectedItem), "Leñador") > 0 _
                   Or InStr(Inventario.ItemName(Inventario.SelectedItem), "Serrucho") > 0 _
                   Or InStr(Inventario.ItemName(Inventario.SelectedItem), "Herrero") > 0 _
                   Or InStr(Inventario.ItemName(Inventario.SelectedItem), "Minero") > 0 _
                   Or InStr(Inventario.ItemName(Inventario.SelectedItem), "Red") > 0 Then
                   

                If Inventario.Equipped(Inventario.SelectedItem) Then
                    Call UsarItem
                    ActivarMacroTrabajo
                Else
                    Call EquiparItem

                End If


            Else
                Call EquiparItem

            End If

        Case eObjType.otAnillo
            Call EquiparItem

        Case eObjType.otFlechas


            Call EquiparItem




        Case Else
            Call UsarItem

        End Select

    End If

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If DragAndDrop Then
        frmMain.MouseIcon = Nothing
        frmMain.MousePointer = 99
        Call SetCursor(General)

    End If

    If Button = 2 And DragAndDrop And Inventario.SelectedItem > 0 And Not Comerciando Then
        If X >= Inventario.OffSetX And Y >= Inventario.OffSetY And X <= picInv.Width And Y <= picInv.Height Then

            Dim NewPosInv As Integer

            NewPosInv = Inventario.ClickItem(X, Y)

            If NewPosInv > 0 Then
                Call WriteIntercambiarInv(Inventario.SelectedItem, NewPosInv, False)
                Call Inventario.Intercambiar(NewPosInv)

            End If
    
        Else

            Dim DropX As Integer, tmpX As Integer

            Dim DropY As Integer, tmpY As Integer

            tmpX = X + picInv.left - pRender.left
            tmpY = Y + picInv.top - pRender.top
        
            If tmpX > 0 And tmpX < pRender.Width And tmpY > 0 And tmpY < pRender.Height Then
                Call ConvertCPtoTP(tmpX, tmpY, DropX, DropY)
        
                'Solo tira a un tilde de distancia...
                If DropX < UserPos.X - 1 Then
                    DropX = UserPos.X - 1
                    DropY = UserPos.Y
                ElseIf DropX > UserPos.X + 1 Then
                    DropX = UserPos.X + 1
                    DropY = UserPos.Y
                ElseIf DropY < UserPos.Y - 1 Then
                    DropY = UserPos.Y - 1
                    DropX = UserPos.X
                ElseIf DropY > UserPos.Y + 1 Then
                    DropY = UserPos.Y + 1
                    DropX = UserPos.X

                End If
            
                If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                    Call WriteDrop(Inventario.SelectedItem, 1, DropX, DropY)
                Else

                    If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                        Inventario.DropX = DropX
                        Inventario.DropY = DropY
                        frmCantidad.Show , frmMain

                    End If

                End If

            End If

        End If

    End If

    DragAndDrop = False

End Sub

Private Sub SendTxt_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un GM"
    Else

        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I         As Long

        Dim tempstr   As String

        Dim CharAscii As Integer
        
        For I = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, I, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)

            End If

        Next I

        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr

        End If
        
        stxtbuffer = SendTxt.Text

    End If

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then

        'Say
        If stxtbuffercmsg <> "" Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)

        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False

    End If

End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

End Sub

Private Sub SendCMSTXT_Change()

    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else

        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I         As Long

        Dim tempstr   As String

        Dim CharAscii As Integer
        
        For I = 1 To Len(SendCMSTXT.Text)
            CharAscii = Asc(mid$(SendCMSTXT.Text, I, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)

            End If

        Next I
        
        If tempstr <> SendCMSTXT.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.Text = tempstr

        End If
        
        stxtbuffercmsg = SendCMSTXT.Text

    End If

End Sub

Private Sub AbrirMenuViewPort()
    #If (ConMenuseConextuales = 1) Then

        If tX >= 1 And tY >= 1 And tY <= MapInfo.Height And tX <= MapInfo.Width Then

            If MapData(tX, tY).CharIndex > 0 Then
                If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
                    Dim I As Long

                    Dim m As New frmMenuseFashion
            
                    Load m
                    m.SetCallback Me
                    m.SetMenuId 1
                    m.ListaInit 2, False
            
                    If charlist(MapData(tX, tY).CharIndex).nombre <> "" Then
                        m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).nombre, True
                    Else
                        m.ListaSetItem 0, "<NPC>", True

                    End If

                    m.ListaSetItem 1, "Comerciar"
            
                    m.ListaFin
                    m.Show , Me

                End If

            End If

        End If

    #End If

End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)

    Select Case MenuId

        Case 0 'Inventario

            Select Case Sel

                Case 0

                Case 1

                Case 2 'Tirar
                    Call TirarItem

                Case 3 'Usar

                    If MainTimer.Check(TimersIndex.PuedeUsarDobleClick) Then
                        Call UsarItem

                    End If

                Case 3 'equipar
                    Call EquiparItem

            End Select
    
        Case 1 'Menu del ViewPort del engine

            Select Case Sel

                Case 0 'Nombre
                    Call WriteLeftClick(tX, tY)
        
                Case 1 'Comerciar
                    Call WriteLeftClick(tX, tY)
                    Call WriteCommerceStart

            End Select

    End Select

End Sub

Private Sub tMouse_Timer()

    If MainTimer.CheckV(TimersIndex.PuedeLanzarHechizo) And MainTimer.CheckV(TimersIndex.PuedeGolpeMagia) And MainTimer.CheckV(TimersIndex.PuedeFlechas) Then
   
        If UsingSkill = eSkill.Proyectiles Then
            Me.MousePointer = 99

            'Else
            '    Me.MousePointer = 2
        End If
  
        tMouse.Enabled = False
    Else
        Me.MousePointer = 0

    End If

End Sub

Private Sub tPass_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyTab Then
        tUser.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        ClickAbrirCuenta

    End If

End Sub

Private Sub tPass_LostFocus()

    If MostrarCrearCuenta = False And tUser.Visible And frmMensaje.Visible = False Then tUser.SetFocus

End Sub

Private Sub tRelampago_Timer()
    Call DoRelampago

End Sub

Private Sub tUser_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If tPass.Text <> "" Then
            ClickAbrirCuenta
        Else
            tPass.SetFocus

        End If

    End If

End Sub

Private Sub tUser_LostFocus()

    If MostrarCrearCuenta = False And tPass.Visible And frmMensaje.Visible = False Then tPass.SetFocus

End Sub

Private Sub WSock_Close()
    Call ClosePj

End Sub

Private Sub WSock_Connect()
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
    Second.Enabled = True

    Select Case EstadoLogin

        Case E_MODO.CrearNuevoPj
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login
        
        Case E_MODO.Normal
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login
            iServer = 0
            iCliente = 0
            DummyCode = StrConv("damn" & StrReverse(UCase$(UserName)) & "you", vbFromUnicode)

        Case E_MODO.Cuentas
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login
            
        Case E_MODO.CrearCuenta
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login
            
        Case E_MODO.BorrarPersonaje
            #If SeguridadAlkon Then
                Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
            #End If
            Call Login

    End Select

End Sub

Private Sub WSock_DataArrival(ByVal bytesTotal As Long)

    Dim RD     As String

    Dim Data() As Byte
    
    WSock.GetData RD
    Data = StrConv(RD, vbFromUnicode)
    
    Call DataCorrect(DummyCode, Data, iServer)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    NotEnoughData = False
    
    'Send buffer to Handle data
    Call HandleIncomingData

End Sub

Private Sub WSock_Error(ByVal Number As Integer, _
                        Description As String, _
                        ByVal sCode As Long, _
                        ByVal Source As String, _
                        ByVal HelpFile As String, _
                        ByVal HelpContext As Long, _
                        CancelDisplay As Boolean)

    '*********************************************
    'Handle socket errors
    '*********************************************
    If Number = 24036 Then
        Call MessageBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    ElseIf Number = 10049 Then
        Call MessageBox("Su equipo no soporta la API de Socket, se cambiará su configuración a Winsock, si problema persiste contacte soporte.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    End If
    
    Call MessageBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    Second.Enabled = False

    WSock.Close

    If Not frmCrearPersonaje.Visible And Not Conectar Then
        Call ClosePj
    Else
        frmCrearPersonaje.MousePointer = 0

    End If

End Sub

Public Sub RefreshMiniMap()
    frmMain.Coord.Caption = "(" & UserPos.X & "," & UserPos.Y & ")"

    Me.shpMiniMapaUser.left = UserPos.X
    Me.shpMiniMapaUser.top = UserPos.Y
    Me.shpMiniMapaVision.left = UserPos.X - 135
    Me.shpMiniMapaVision.top = UserPos.Y - 135
    Me.imgMiniMapa.Refresh

End Sub

Private Sub imgMiniMapa_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    If X > 1077 Then X = 1077
    If X < 23 Then X = 23
    If Y > 1477 Then Y = 1477
    If Y < 23 Then Y = 23

    If Button = vbRightButton Then
        Call WriteWarpChar("YO", UserMap, CInt(X - 1), CInt(Y - 1))
        Call RefreshMiniMap

    End If

End Sub

Private Sub cmdLanzar_Click()

    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.HabilitaLanzarHechizo) Then
        If UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

            End With

        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True

        End If

    End If

End Sub

Public Sub EnviarCorreoVal()
    If frmMain2.tUser = "" Then
         MessageBox "Ingrese un nombre de cuenta"
        Exit Sub
    End If
    If frmMain2.tEmail = "" Then
        MessageBox "Ingrese una direccion e-Email"
        Exit Sub
    End If
    CodVerificacion = RandomLetrasMayusculas(1) & RandomLetrasMinusculas(1) & Int((9 * Rnd) + 1) & Chr(Int((Rnd * 25) + 65)) & RandomLetrasMinusculas(1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1)
    CorreoVal = frmMain2.tEmail

    Set oMail = New clsCDOmail
    With oMail
        'datos para enviar
        .servidor = "smtp.gmail.com"
        .puerto = 465
        .UseAuntentificacion = True
        .ssl = True
        .Usuario = "soporteaoyind@gmail.com"
        .PassWord = "ruKgym-xisqom-3pothe"

        .Asunto = "Codigo de Validacion"
        '.Adjunto = "c:\archivo.zip"
        .de = "SoporteAoyind3@gmail.com"
        .para = CorreoVal
        .Mensaje = "Se ha solicitado codigo de validacion para la cuenta: " & tUser.Text & " por favor ingrese en el juego el siguiente Codigo: " & CodVerificacion

        .Enviar_Backup    ' manda el mail

    End With

    Set oMail = Nothing
End Sub

