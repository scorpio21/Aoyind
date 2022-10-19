VERSION 5.00
Begin VB.Form frmCustomKeys 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   6240
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar y Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   8400
      Width           =   7935
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Otros"
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   2280
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   2280
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   2280
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   2280
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   2280
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   2280
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   2280
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   2280
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   2280
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   2280
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "GM Panel"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   63
         Top             =   3870
         Width           =   1695
      End
      Begin VB.Label lblAnclar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Anclar/Desanclar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   61
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Montar / Desmontar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   57
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Ver Mapa"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   55
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Seg. de Resucitación"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   53
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Salir"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   51
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Capturar Pantalla"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Macro Trabajo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Meditar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar Opciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar/Ocultar FPS"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Hablar"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   1920
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   1920
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Hablar al Clan"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Hablar a Todos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Acciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   1920
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1920
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1920
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1920
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Atacar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Usar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Tirar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Ocultar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Robar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Domar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Equipar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Agarrar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Opciones Personales"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Mostrar/Ocultar Nombres"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Corregir Posicion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Activar/Desactivar Musica"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Movimiento"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Derecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Izquierda"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Abajo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Arriba"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label lblTecladoNormal 
      BackStyle       =   0  'Transparent
      Caption         =   "Teclado Normal"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label lblTecladoAWSD 
      BackStyle       =   0  'Transparent
      Caption         =   "Teclado AWSD"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   58
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Image imgCustomKeyNormal 
      Height          =   1500
      Left            =   120
      Top             =   6600
      Width           =   3750
   End
   Begin VB.Image imgCustomKeyAWSD 
      Height          =   1500
      Left            =   4320
      Top             =   6600
      Width           =   3750
   End
End
Attribute VB_Name = "frmCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

''
'frmCustomKeys - Allows the user to customize keys.
'Implements class clsCustomKeys
'
'@author Rapsodius
'@date 20070805
'@version 1.0.0
'@see clsCustomKeys

Option Explicit

Private Sub Command1_Click()
    Call CustomKeys.LoadDefaults

    Dim I As Long

    For I = 1 To CustomKeys.count
        Text1(I).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(I))
    Next I

End Sub

Private Sub Command2_Click()

    Dim I As Long

    For I = 1 To CustomKeys.count

        If LenB(Text1(I).Text) = 0 Then
            AddtoRichPicture "Hay teclas invalidas. Por favor revise y vuelva a intentarlo!", 200, 200, 200, False, False, False
            Exit Sub

        End If

    Next I

    Call CustomKeys.SaveCustomKeys

    Unload Me

End Sub

Private Sub Form_Load()
    'imgCustomKeyNormal.Picture = LoadPictureEX("frmKeysConfigurationSelectNormalKeyboard.jpg")
    'imgCustomKeyAWSD.Picture = LoadPictureEX("frmKeysConfigurationSelectAlternativeKeyboard.jpg")
    Me.Picture = LoadPictureEX("VENTANACONFIGURARTECLAS.jpg")

    Dim I As Long
    
    For I = 1 To CustomKeys.count
        Text1(I).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(I))
    Next I

End Sub

Private Sub imgCustomKeyAWSD_Click()
    Call CustomKeys.LoadAWSD

    Dim I As Long

    For I = 1 To CustomKeys.count
        Text1(I).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(I))
    Next I

End Sub

Private Sub imgCustomKeyNormal_Click()
    Call CustomKeys.LoadDefaults

    Dim I As Long
    
    For I = 1 To CustomKeys.count
        Text1(I).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(I))
    Next I

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim I As Long
    
    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    
    Text1(Index).Text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
    For I = 1 To CustomKeys.count

        If I <> Index Then
            If CustomKeys.BindedKey(I) = KeyCode Then
                Text1(Index).Text = "" 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub

            End If

        End If

    Next I
    
    CustomKeys.BindedKey(Index) = KeyCode

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0

End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call Text1_KeyDown(Index, KeyCode, Shift)

End Sub
