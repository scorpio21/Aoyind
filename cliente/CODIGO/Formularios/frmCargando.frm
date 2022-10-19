VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BProg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1185
      Picture         =   "frmCargando.frx":7D89E
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   1
      Top             =   6060
      Width           =   315
   End
   Begin VB.PictureBox BBProg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   960
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim f As Integer

Private Sub Form_Load()

    'Me.Picture = LoadPictureEX("cargando.jpg")
End Sub
 
