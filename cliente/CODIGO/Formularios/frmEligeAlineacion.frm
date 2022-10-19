VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgReal 
      Height          =   1005
      Left            =   795
      Tag             =   "1"
      Top             =   120
      Width           =   5745
   End
   Begin VB.Image imgNeutral 
      Height          =   795
      Left            =   810
      Tag             =   "1"
      Top             =   2100
      Width           =   5730
   End
   Begin VB.Image imgLegal 
      Height          =   825
      Left            =   810
      Tag             =   "1"
      Top             =   1200
      Width           =   5715
   End
   Begin VB.Image imgCaos 
      Height          =   795
      Left            =   825
      Tag             =   "1"
      Top             =   4110
      Width           =   5700
   End
   Begin VB.Image imgCriminal 
      Height          =   825
      Left            =   825
      Tag             =   "1"
      Top             =   3030
      Width           =   5745
   End
   Begin VB.Image imgSalir 
      Height          =   315
      Left            =   5650
      Tag             =   "1"
      Top             =   4880
      Width           =   930
   End
End
Attribute VB_Name = "frmEligeAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmEligeAlineacion.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

'
'Private cBotonCriminal As clsGraphicalButton
'Private cBotonCaos As clsGraphicalButton
'Private cBotonLegal As clsGraphicalButton
'Private cBotonNeutral As clsGraphicalButton
'Private cBotonReal As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastPressed  As clsGraphicalButton

Private Enum eAlineacion

    ieREAL = 0
    ieCAOS = 1
    ieNeutral = 2
    ieLegal = 4
    ieCriminal = 5

End Enum

Private Sub Form_Load()
    'Call SetTranslucent(Me.hwnd, NTRANS_GENERAL)

    Me.Picture = LoadPictureEX("VentanaFundarClan.jpg")
    
    Call LoadButtons

End Sub

Private Sub LoadButtons()

    '    Set cBotonCriminal = New clsGraphicalButton
    '    Set cBotonCaos = New clsGraphicalButton
    '    Set cBotonLegal = New clsGraphicalButton
    '    Set cBotonNeutral = New clsGraphicalButton
    '    Set cBotonReal = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    '
    '    Call cBotonCriminal.Initialize(imgCriminal, "", _
    '                                    "BotonCriminal.jpg", _
    '                                    "BotonCriminal.jpg", Me)
    '
    '    Call cBotonCaos.Initialize(imgCaos, "", _
    '                                    "BotonCaos.jpg", _
    '                                    "BotonCaos.jpg", Me)
    '
    '    Call cBotonLegal.Initialize(imgLegal, "", _
    '                                    "BotonLegal.jpg", _
    '                                    "BotonLegal.jpg", Me)
    '
    '    Call cBotonNeutral.Initialize(imgNeutral, "", _
    '                                    "BotonNeutral.jpg", _
    '                                    "BotonNeutral.jpg", Me)
    '
    '    Call cBotonReal.Initialize(imgReal, "", _
    '                                    "BotonReal.jpg", _
    '                                    "BotonReal.jpg", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, "BotonSalirAlineacion.jpg", "BotonSalirRolloverAlineacion.jpg", "BotonSalirClickAlineacion.jpg", Me)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then MoverVentana (Me.hwnd)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub imgCaos_Click()
    Call WriteGuildFundate(eAlineacion.ieCAOS)
    Unload Me

End Sub

Private Sub imgCriminal_Click()
    Call WriteGuildFundate(eAlineacion.ieCriminal)
    Unload Me

End Sub

Private Sub imgLegal_Click()
    Call WriteGuildFundate(eAlineacion.ieLegal)
    Unload Me

End Sub

Private Sub imgNeutral_Click()
    Call WriteGuildFundate(eAlineacion.ieNeutral)
    Unload Me

End Sub

Private Sub imgReal_Click()
    Call WriteGuildFundate(eAlineacion.ieREAL)
    Unload Me

End Sub

Private Sub imgSalir_Click()
    Unload Me

End Sub
