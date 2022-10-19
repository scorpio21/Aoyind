VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BF3128D8-55B8-11D4-8ED4-00E07D815373}#1.0#0"; "mbprgbar.ocx"
Begin VB.Form frmMain1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "AoYind 3"
   ClientHeight    =   11730
   ClientLeft      =   0
   ClientTop       =   555
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmMainN.frx":0CCA
   ScaleHeight     =   782
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picarmadura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2610
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   11175
      Visible         =   0   'False
      Width           =   480
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
      Left            =   14760
      Picture         =   "frmMainN.frx":C490A
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3480
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
         Picture         =   "frmMainN.frx":C69BC
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdEfecto 
      Caption         =   "Efecto"
      Enabled         =   0   'False
      Height          =   315
      Left            =   12615
      TabIndex        =   27
      Top             =   10815
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox pRender 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9093
      Left            =   0
      ScaleHeight     =   606
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   5
      Top             =   1917
      Width           =   12000
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
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Chat"
         Top             =   3120
         Visible         =   0   'False
         Width           =   2700
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
         TabIndex        =   24
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
      Left            =   240
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   788
      TabIndex        =   2
      Top             =   135
      Width           =   11820
      Begin VB.Timer tRelampago 
         Enabled         =   0   'False
         Interval        =   7500
         Left            =   0
         Top             =   0
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
         Left            =   11550
         Picture         =   "frmMainN.frx":C6B32
         ScaleHeight     =   91
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   15
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
            Picture         =   "frmMainN.frx":C7F5C
            ScaleHeight     =   7
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1020
            Width           =   210
         End
      End
      Begin MSWinsockLib.Winsock WSock 
         Left            =   6960
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tMouse 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   8400
         Top             =   0
      End
   End
   Begin VB.PictureBox imgMiniMapa 
      BorderStyle     =   0  'None
      Height          =   1498
      Left            =   11640
      ScaleHeight     =   1500
      ScaleMode       =   0  'User
      ScaleWidth      =   1100
      TabIndex        =   26
      Top             =   75
      Visible         =   0   'False
      Width           =   1500
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
      Begin VB.Shape shpMiniMapaVision 
         Height          =   315
         Left            =   520
         Top             =   614
         Width           =   375
      End
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
      Left            =   240
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1513
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
      Left            =   240
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1513
      Visible         =   0   'False
      Width           =   11820
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
      Height          =   3416
      Left            =   12360
      MousePointer    =   99  'Custom
      ScaleHeight     =   228
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   11
      Top             =   3041
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      Height          =   3595
      Left            =   12345
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   12
      Top             =   2996
      Width           =   2700
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   4080
         Width           =   2175
      End
   End
   Begin MBProgressBar.ProgressBar Experiencia 
      Height          =   135
      Left            =   0
      TabIndex        =   29
      Top             =   11011
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   238
      BorderStyle     =   0
      Value           =   50
      Percentage      =   50
      Smooth          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmMainN.frx":C80D2
      BarPicture      =   "frmMainN.frx":CD51A
      TextAfterCaption=   "%"
      Style           =   1
   End
   Begin MBProgressBar.ProgressBar bar_salud 
      Height          =   225
      Left            =   12420
      TabIndex        =   35
      Top             =   8790
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   397
      BorderStyle     =   0
      CaptionType     =   2
      Value           =   50
      Percentage      =   50
      Smooth          =   -1  'True
      TextColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmMainN.frx":D2962
      BarPicture      =   "frmMainN.frx":DF6BC
      Style           =   1
   End
   Begin MBProgressBar.ProgressBar Bar_Mana 
      Height          =   225
      Left            =   12420
      TabIndex        =   36
      Top             =   9375
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   397
      BorderStyle     =   0
      CaptionType     =   2
      Value           =   50
      Percentage      =   50
      Smooth          =   -1  'True
      TextColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmMainN.frx":E046C
      BarPicture      =   "frmMainN.frx":ED1C6
      Style           =   1
   End
   Begin MBProgressBar.ProgressBar bar_sta 
      Height          =   225
      Left            =   12420
      TabIndex        =   37
      Top             =   9960
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   397
      BorderStyle     =   0
      CaptionType     =   2
      Value           =   50
      Percentage      =   50
      Smooth          =   -1  'True
      TextColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmMainN.frx":EDF76
      BarPicture      =   "frmMainN.frx":FACD0
      Style           =   1
   End
   Begin MBProgressBar.ProgressBar Bar_Agua 
      Height          =   150
      Left            =   13980
      TabIndex        =   38
      Top             =   10545
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   265
      BorderStyle     =   0
      CaptionType     =   2
      Value           =   50
      Percentage      =   50
      Smooth          =   -1  'True
      TextColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmMainN.frx":FBA80
      BarPicture      =   "frmMainN.frx":1087DA
      Style           =   1
   End
   Begin MBProgressBar.ProgressBar bar_comida 
      Height          =   150
      Left            =   12240
      TabIndex        =   39
      Top             =   10545
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   265
      BorderStyle     =   0
      CaptionType     =   2
      Value           =   50
      Percentage      =   50
      Smooth          =   -1  'True
      TextColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "frmMainN.frx":108E56
      BarPicture      =   "frmMainN.frx":115BB0
      Style           =   1
   End
   Begin VB.Label QuestBoton 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   12840
      MouseIcon       =   "frmMainN.frx":11629C
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   10440
      Width           =   1455
   End
   Begin VB.Label lbStats 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   12300
      TabIndex        =   41
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
      TabIndex        =   40
      Top             =   7515
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   3
      Left            =   13350
      MousePointer    =   99  'Custom
      Top             =   9870
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   2
      Left            =   13230
      MousePointer    =   99  'Custom
      Top             =   9285
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   12915
      MousePointer    =   99  'Custom
      Top             =   8685
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image Image1 
      Height          =   270
      Index           =   0
      Left            =   13020
      MousePointer    =   99  'Custom
      Top             =   8100
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Menu 
      Height          =   2940
      Left            =   12000
      Picture         =   "frmMainN.frx":116F66
      Top             =   7995
      Visible         =   0   'False
      Width           =   3360
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
      Left            =   5835
      TabIndex        =   32
      Top             =   11010
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblDIATEST 
      Height          =   495
      Left            =   6225
      TabIndex        =   28
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
      TabIndex        =   23
      Top             =   10260
      Width           =   1935
   End
   Begin VB.Label lblItemInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   12135
      TabIndex        =   22
      Top             =   2730
      Width           =   3135
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   11355
      Width           =   1095
   End
   Begin VB.Image btnHechizos 
      Height          =   540
      Left            =   13680
      MousePointer    =   99  'Custom
      Top             =   2070
      Width           =   1545
   End
   Begin VB.Image btnInventario 
      Height          =   540
      Left            =   12150
      MousePointer    =   99  'Custom
      Top             =   2070
      Visible         =   0   'False
      Width           =   1530
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   10
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
      Left            =   12660
      TabIndex        =   9
      Top             =   1155
      Width           =   360
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   0
      Left            =   15000
      MousePointer    =   99  'Custom
      Top             =   3510
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   1
      Left            =   15000
      MousePointer    =   99  'Custom
      Top             =   3165
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdInfo 
      Height          =   405
      Left            =   15120
      MousePointer    =   99  'Custom
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000000"
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   13155
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   2400
      Width           =   30
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sensui"
      BeginProperty Font 
         Name            =   "Augusta"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   13140
      TabIndex        =   6
      Top             =   1290
      Width           =   1545
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
      TabIndex        =   14
      Top             =   10560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image CmdLanzar 
      Height          =   600
      Left            =   12255
      MousePointer    =   99  'Custom
      Top             =   6765
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image iBEXP 
      Height          =   300
      Left            =   13320
      Top             =   1440
      Visible         =   0   'False
      Width           =   1605
   End
End
Attribute VB_Name = "frmMain1"
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

Public tX                As Integer

Public tY                As Integer

Public MouseX            As Integer

Public MouseY            As Integer

Public MouseBoton        As Long

Public MouseShift        As Long

Private clicX            As Long

Private clicY            As Long

Public SinOrtografia     As Boolean

'Dim gDSB As DirectSoundBuffer
'Dim gD As DSBUFFERDESC
'Dim gW As WAVEFORMATEX
Dim gFileName            As String

'Dim dsE As DirectSoundEnum
'Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying         As PlayLoop

Dim PuedeMacrear         As Boolean

Dim OldYConsola          As Integer

Public hlst              As clsGraphicalList

Dim InvX                 As Integer

Dim InvY                 As Integer

Public WithEvents Client As CSocketMaster
Attribute Client.VB_VarHelpID = -1

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

    BarritaConsola.Top = TempY
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

    BarritaHechiz.Top = TempY

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

        TempY = BarritaConsola.Top + (Y - OldYConsola)

        If TempY < 16 Then TempY = 16
        If TempY > 68 Then TempY = 68
        If LineasConsola <= 6 Then TempY = 68
        OffSetConsola = Int((TempY - 16) * (LineasConsola - 6) / 52)
        BarritaConsola.Top = TempY
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
        TempY = BarritaHechiz.Top + (Y - hlst.OldY)

        If TempY < 16 Then TempY = 16
        If TempY > 150 Then TempY = 150
        If hlst.ListCount <= MaxItems Then TempY = 150
        hlst.Scroll = Int((TempY - 16) * (hlst.ListCount - MaxItems) / 134)
        BarritaHechiz.Top = TempY

    End If

End Sub

Private Sub btnHechizos_Click()
    Call Audio.PlayWave(SND_CLICK)
    'picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    Coord.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    cmdMoverHechi(0).Enabled = True
    cmdMoverHechi(1).Enabled = True
    
    btnInventario.Visible = True
    btnHechizos.Visible = False
    
    'btnHechizos.Picture = LoadPictureEX("btnHechizos_R.jpg")
    'btnInventario.Picture = LoadPictureEX("btnInventario.jpg")
    BarraHechiz.Visible = True
    lblItemInfo.Visible = False
    
End Sub

Private Sub btnInventario_Click()
    Call Audio.PlayWave(SND_CLICK)
    'picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    Coord.Visible = True
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    cmdMoverHechi(0).Enabled = False
    cmdMoverHechi(1).Enabled = False
    
    btnInventario.Visible = False
    btnHechizos.Visible = True
    
    'btnInventario.Picture = LoadPictureEX("btnInventario_R.jpg")
    'btnHechizos.Picture = LoadPictureEX("btnHechizos.jpg")
    BarraHechiz.Visible = False
    
    lblItemInfo.Visible = True
    
End Sub

Private Sub cmdEfecto_Click()
    Hora = Hora + 1
    SetDayLight
    'AlphaSalir = 0
    'AlphaBlood = 255
    'TextKillsType = RandomNumber(2, 9)
    'AlphaTextKills = 255
    'Call Audio.PlayWave(258 + TextKillsType)

    'AlphaCeguera = 255
End Sub

Private Sub cmdMoverHechi_Click(index As Integer)

    If hlst.ListIndex = -1 Then Exit Sub

    Dim sTemp As String

    Select Case index

        Case 1 'subir

            If hlst.ListIndex = 0 Then Exit Sub

        Case 0 'bajar

            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub

    End Select

    Call WriteMoveSpell(index, hlst.ListIndex + 1)
    
    Select Case index

        Case 1 'subir
            sTemp = hlst.List(hlst.ListIndex - 1)
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex - 1

        Case 0 'bajar
            sTemp = hlst.List(hlst.ListIndex + 1)
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
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

    If Full Then
        pRender.Move 0, 0, 1024, 782
    Else
        pRender.Move 2, 125, 800, 608

        'pRender.Move 13, 169, 768, 576
        'pRender.Move 11, 133, Render_Width, Render_Height
    End If

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

Private Sub QuestBoton_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    On Error GoTo QuestBoton_MouseMove_Err

    If QuestBoton.Tag = "0" Then
        'QuestBoton.Picture = LoadInterface("questover.bmp")
        QuestBoton.Tag = "1"

    End If
    
    Exit Sub

QuestBoton_MouseMove_Err:

    'Call RegistrarError(Err.Number, Err.Description, "frmMain.QuestBoton_MouseMove", Erl)
    Resume Next
    
End Sub

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
        'If Not PicResu.Visible Then
        '    PicResu.Visible = True
        'End If
    Else

        'If PicResu.Visible Then
        '    PicResu.Visible = False
        'End If
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

Private Sub mnuEquipar_Click()
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
    #If RenderFull = 1 Then

        Menu.Visible = True
    #End If
    bar_salud.Visible = False
    Bar_Mana.Visible = False
    bar_sta.Visible = False
    bar_comida.Visible = False
    Bar_Agua.Visible = False

    Image1(0).Visible = True
    Image1(1).Visible = True
    Image1(2).Visible = True
    Image1(3).Visible = True

End Sub

Private Sub lbStats_Click()
    #If RenderFull = 1 Then
        Menu.Visible = False
    #End If
    bar_salud.Visible = True
    Bar_Mana.Visible = True
    bar_sta.Visible = True
    bar_comida.Visible = True
    Bar_Agua.Visible = True

    Image1(0).Visible = False
    Image1(1).Visible = False
    Image1(2).Visible = False
    Image1(3).Visible = False

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

Private Sub Picture3_Click()

End Sub

Private Sub pRender_Click()

    If Conectar Then Exit Sub
    If Cartel Then Cartel = False

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
        Debug.Print tX & " - " & tY

        If Not InGameArea() Then Exit Sub
        If Not InMapBounds(tX, tY) Then Exit Sub
        
        If MouseShift = 0 And (MapData(tX, tY).Graphic(4).GrhIndex = 0 Or bTecho) Then

            'If MouseBoton <> vbRightButton Then
            If MouseBoton Then
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                
                    If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                    
                    '                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                    '                        'frmMain.MousePointer = vbDefault
                    '                        Call SetCursor(General)
                    '                        UsingSkill = 0
                    '                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                    '                            Call AddtoRichPicture("No podés lanzar proyectiles tan rapido.", .red, .green, .blue, .bold, .italic)
                    '                        End With
                    '                        Exit Sub
                    '                    End If
                    
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

End Sub

Private Sub pRender_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    MouseBoton = Button
    MouseShift = Shift

End Sub

Private Sub pRender_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    MouseX = X
    MouseY = Y
    
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

Private Sub UsarItem()

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then Call WriteUseItem(Inventario.SelectedItem)

End Sub

Private Sub EquiparItem()

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then Call WriteEquipItem(Inventario.SelectedItem)

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

Private Sub CmdLanzar_MouseMove(Button As Integer, _
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
    
    frmMain.Caption = "AoYind 3"
    'PanelDer.Picture = LoadPicture(App.path & _
     "\Graficos\Principalnuevo_sin_energia.jpg")
    
    'InvEqu.Picture = LoadPicture(App.path & _
     "\Graficos\Centronuevoinventario.jpg")
   
    'frmMain.iBEXP.Picture = LoadPictureEX("BARRAEXP.jpg")
    
    Me.Picture = LoadPictureEX("VentanaPrincipalm.bmp")
    CmdLanzar.Picture = LoadPictureEX("btnLanzar.jpg")
    'picInv.Picture = LoadPictureEX("VentanaPrincipalInv.jpg")
    
    'BarraHechiz.Picture = LoadPictureEX("BarraHechiz.jpg")
    
    btnHechizos.Picture = LoadPictureEX("btnHechizos_R.bmp")
    btnInventario.Picture = LoadPictureEX("btnInventario_R.bmp")
    'neo parcheo ruta
    imgMiniMapa.Picture = LoadPicture(PathRecursosCliente & "\Recursos\minimapadefault.bmp")
    ' imgMiniMapa.Picture = LoadPicture(App.path & "\Recursos\minimapadefault.bmp")
    'MANShp.Picture = LoadPictureEX("barMana.jpg")
    'Hpshp.Picture = LoadPictureEX("BarHp.jpg")
    'STAShp.Picture = LoadPictureEX("BarSta.jpg")
    
    'picInv.Height = 246
    
    btnInventario.MouseIcon = CmdLanzar.MouseIcon
    btnHechizos.MouseIcon = CmdLanzar.MouseIcon
   
    Set hlst = New clsGraphicalList
    Call hlst.Initialize(Me.picHechiz, RGB(200, 190, 190))
    
    tUser.BackColor = RGB(200, 200, 200)
    tPass.BackColor = RGB(200, 200, 200)
    tEmail.BackColor = RGB(200, 200, 200)
    tRePass.BackColor = RGB(200, 200, 200)
   
    tUser.Top = 612
    tPass.Top = 637
    tPass.Left = 428
    tUser.Left = 428
    
    Me.Left = 0
    Me.Top = 0
   
    If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And (Not frmBancoObj.Visible) And (Not frmMSG.Visible) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
        Debug.Print "Precarga"

    End If

    OpcionesPath = App.path & "\init\Config.ini"
    ActivarAuras = GetVar(OpcionesPath, "AURAS", "AuraActiva")
    RotarActivado = GetVar(OpcionesPath, "AURAS", "rotacion")

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

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0

End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

Private Sub Image1_Click(index As Integer)
    'Call Audio.PlayWave(SND_CLICK)

    Select Case index

        Case 0
            Call frmOpciones.Show(vbModeless, frmMain)
            
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            Call WriteRequestAtributes
            Call WriteRequestSkills
            Call WriteRequestMiniStats
            Call WriteRequestFame
            Call FlushBuffer
            
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        
        Case 2

            If frmGuildLeader.Visible Then Unload frmGuildLeader
            
            Call WriteRequestGuildLeaderInfo

        Case 3
            Call WriteRequestPartyForm

    End Select

End Sub

Private Sub Image3_Click(index As Integer)

    Select Case index

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

            tmpX = X + picInv.Left - pRender.Left
            tmpY = Y + picInv.Top - pRender.Top
        
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

    Me.shpMiniMapaUser.Left = UserPos.X
    Me.shpMiniMapaUser.Top = UserPos.Y
    Me.shpMiniMapaVision.Left = UserPos.X - 135
    Me.shpMiniMapaVision.Top = UserPos.Y - 135
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
