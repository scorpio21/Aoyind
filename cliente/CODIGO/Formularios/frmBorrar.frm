VERSION 5.00
Begin VB.Form frmBorrar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Borrar Personaje"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Salir"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtNombrePersonaje 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtNombreCuenta 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese su contraseña"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblNombrePersonaje 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblNombreCuenta 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "frmBorrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBorrar_Click()

    If txtNombreCuenta.Text = "" Then
        MessageBox ("Escriba el nombre de su cuenta.")
        Exit Sub

    End If
    
    If txtNombrePersonaje.Text = "" Then
        MessageBox ("Escriba el nombre de su personaje.")
        Exit Sub

    End If
    
    If txtPassword.Text = "" Then
        MessageBox ("Escriba su contraseña.")
        Exit Sub

    End If
    
    If MsgBox("¿Está seguro de borrar al personaje " & txtNombrePersonaje.Text & "?", vbYesNo, "AoYind 3 - Borrar personaje") = vbYes Then
        UserPassword = txtPassword.Text
        UserName = txtNombreCuenta.Text
        
        iServer = 0
        iCliente = 0
        DummyCode = StrConv(StrReverse("conectar") & "CuEnTa", vbFromUnicode)
        
        If Not ClientSetup.WinSock Then
            frmMain.Client.CloseSck
                    
            EstadoLogin = E_MODO.BorrarPersonaje
                    
            frmMain.Client.Connect IpServidor, PuertoServidor
                    
            If Not frmMain.Client.State <> SockState.sckConnected Then
                MessageBox "Error de conexión."
            Else
                Call Login

            End If

        Else
            frmMain.WSock.Close
                    
            EstadoLogin = E_MODO.BorrarPersonaje
                    
            frmMain.WSock.Connect IpServidor, PuertoServidor
                    
            If Not frmMain.WSock.State <> SockState.sckConnected Then
                MessageBox "Error de conexión."
            Else
                Call Login

            End If

        End If

    End If
    
End Sub

Private Sub cmdCerrar_Click()
    Unload Me

End Sub

Private Sub Form_Load()
    Me.Picture = LoadPictureEX("VENTANABORRARPJ.jpg")

End Sub
