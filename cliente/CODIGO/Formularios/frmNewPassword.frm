VERSION 5.00
Begin VB.Form frmNewPassword 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   293
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   293
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1650
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   293
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar contraseña nueva:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña nueva:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   1410
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña anterior:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    If Text2.Text <> Text3.Text Then
        Call MsgBox("Las contraseñas no coinciden", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Cambiar Contraseña")
        Exit Sub

    End If

    If Len(Text2.Text) < 6 Then
        Call MsgBox("La contraseña debe contener al menos 6 caracteres.", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Cambiar Contraseña")
        Exit Sub

    End If
    
    Call WriteChangePassword(Text1.Text, Text2.Text)
    Unload Me

End Sub

Private Sub Command2_Click()
    Unload Me

End Sub

Private Sub Form_Load()
    Me.Picture = LoadPictureEX("VENTANACAMBIARPASS.jpg")

End Sub

