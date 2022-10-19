VERSION 5.00
Begin VB.Form frmNavegador 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   223
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMail 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtUsuario 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton cmdCrear2 
      Caption         =   "Crear Cuenta"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblMail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmNavegador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum eTipo

    Crear = 1
    Recuperar = 2
    Borrar = 3

End Enum

Public TIPO As eTipo

Private Sub cmdCrear2_Click()
    Call Audio.PlayWave(SND_CLICKNEW)
    
    If txtUsuario.Text = "" Then
        MessageBox "Escriba un usuario."
        Exit Sub

    End If
    
    If txtPassword.Text = "" Then
        MessageBox "Escriba un password."
        Exit Sub

    End If
    
    If txtMail.Text = "" Then
        MessageBox "Escriba un mail."
        Exit Sub

    End If
    
    UserAccount = txtUsuario.Text
    UserPassword = MD5(txtPassword.Text)
    UserEmail = txtMail.Text
    
    If Right$(txtUsuario.Text, 1) = " " Then
        UserAccount = RTrim$(UserAccount)
        MessageBox "Nombre invalido, se han removido los espacios al final del nombre"

    End If

    If Len(txtUsuario.Text) > 20 Then
        MessageBox "El nombre es demasiado largo, debe tener como máximo 20 letras."
        Exit Sub

    End If
    
    If Not ClientSetup.WinSock Then
        frmMain.Client.CloseSck
                
        EstadoLogin = E_MODO.CrearCuenta
                
        frmMain.Client.Connect IpServidor, PuertoServidor
                
        If Not frmMain.Client.State <> SockState.sckConnected Then
            
            MessageBox "Error: Se ha perdido la conexion con el server."
            Unload Me
                    
        Else
            Call Login

        End If

    Else
        frmMain.WSock.Close
                
        EstadoLogin = E_MODO.CrearCuenta
                
        frmMain.WSock.Connect IpServidor, PuertoServidor
                
        If Not frmMain.WSock.State <> SockState.sckConnected Then
            
            MessageBox "Error: Se ha perdido la conexion con el server."
            Unload Me
                    
        Else
            Call Login

        End If

    End If
    
    Unload Me
    
End Sub

Private Sub Command1_Click()
    Call Audio.PlayWave(SND_CLICKNEW)
    Unload Me

End Sub

Private Sub Form_Load()

    If TIPO = Crear Then

        'WB.Navigate ("http://www.aoyind.com/crearcuenta.php")
    End If

End Sub
