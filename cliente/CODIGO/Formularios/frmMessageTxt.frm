VERSION 5.00
Begin VB.Form frmMessageTxt 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mensajes Predefinidos"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   4740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cancelCmd 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2520
      TabIndex        =   21
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton okCmd 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   10
      Top             =   3915
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   9
      Top             =   3555
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   8
      Top             =   3195
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   7
      Top             =   2835
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   6
      Top             =   2475
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   5
      Top             =   2115
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Top             =   1755
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   1395
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   1035
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   675
      Width           =   3400
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 10:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   870
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 9:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 8:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 7:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 6:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 5:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 4:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 3:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 2:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje 1:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   780
   End
End
Attribute VB_Name = "frmMessageTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancelCmd_Click()
    Unload Me

End Sub

Private Sub Form_Load()
    Me.Picture = LoadPictureEX("VENTANAMENSAJESPREDEFINIDOS.jpg")

    Dim I As Long
    
    For I = 0 To 9
        messageTxt(I) = CustomMessages.Message(I)
    Next I

End Sub

Private Sub okCmd_Click()

    On Error GoTo ErrHandler

    Dim I As Long
    
    For I = 0 To 9
        CustomMessages.Message(I) = messageTxt(I)
    Next I
    
    Unload Me
    Exit Sub

ErrHandler:

    'Did detected an invalid message??
    If err.Number = CustomMessages.InvalidMessageErrCode Then
        Call MessageBox("El Mensaje " & CStr(I + 1) & " es inválido. Modifiquelo por favor.")

    End If

End Sub
