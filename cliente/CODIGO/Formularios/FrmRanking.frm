VERSION 5.00
Begin VB.Form FrmRanking 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   Picture         =   "FrmRanking.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   135
   End
   Begin VB.Image ImgOro 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":5917
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Image ImgFrags 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":65E1
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Image ImgReto 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":72AB
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image ImgNivel 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":7F75
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Image ImgTorneo 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":8C3F
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   2295
   End
   Begin VB.Image ImgClan 
      Height          =   255
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":9909
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "FrmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Image1_Click()
Unload Me
frmMain.SetFocus
End Sub






Private Sub ImgClan_Click()
Call Audio.PlayWave(SND_CLICK)

Call WriteSolicitarRanking(TopClanes)
RankingOro = ""
End Sub


Private Sub ImgFrags_Click()

Call Audio.PlayWave(SND_CLICKNEW)
 Call WriteSolicitarRanking(TopFrags)
   FrmRanking2.Picture = LoadPictureEX("RankingAsesinados.jpg")
    RankingOro = ""
   
End Sub

Private Sub ImgNivel_Click()

Call Audio.PlayWave(SND_CLICKNEW)
Call WriteSolicitarRanking(TopLevel)
RankingOro = ""
End Sub

Private Sub ImgOro_Click()
Call Audio.PlayWave(SND_CLICKNEW)
FrmRanking2.Picture = LoadPictureEX("RankingOro_1.jpg")
Call WriteSolicitarRanking(TopOro)
RankingOro = "$"
End Sub

Private Sub ImgReto_Click()
Call Audio.PlayWave(SND_CLICKNEW)
FrmRanking2.Picture = LoadPictureEX("RankingRetos_1.jpg")
Call WriteSolicitarRanking(TopRetos)
RankingOro = ""
End Sub

Private Sub ImgTorneo_Click()
Call Audio.PlayWave(SND_CLICKNEW)
 Call WriteSolicitarRanking(TopLevel)
   FrmRanking2.Picture = LoadPictureEX("RankingLevel.jpg")
    RankingOro = ""
End Sub

Private Sub Label1_Click()
Call Audio.PlayWave(SND_CLICKNEW)
Unload Me
frmMain.SetFocus
End Sub



