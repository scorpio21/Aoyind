VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   Caption         =   "Administraci�n del Clan"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   494
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFiltrarMiembros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3075
      TabIndex        =   6
      Top             =   2340
      Width           =   2580
   End
   Begin VB.TextBox txtFiltrarClanes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   2340
      Width           =   2580
   End
   Begin VB.TextBox txtguildnews 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   195
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3435
      Width           =   5475
   End
   Begin VB.ListBox solicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      ItemData        =   "frmGuildLeader.frx":0000
      Left            =   195
      List            =   "frmGuildLeader.frx":0002
      TabIndex        =   2
      Top             =   5100
      Width           =   2595
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0004
      Left            =   3060
      List            =   "frmGuildLeader.frx":0006
      TabIndex        =   1
      Top             =   525
      Width           =   2595
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":0008
      Left            =   180
      List            =   "frmGuildLeader.frx":000A
      TabIndex        =   0
      Top             =   525
      Width           =   2595
   End
   Begin VB.Image imgCerrar 
      Height          =   495
      Left            =   2940
      Tag             =   "1"
      Top             =   6765
      Width           =   2775
   End
   Begin VB.Image imgPropuestasAlianzas 
      Height          =   495
      Left            =   2940
      Tag             =   "1"
      Top             =   6255
      Width           =   2775
   End
   Begin VB.Image imgPropuestasPaz 
      Height          =   495
      Left            =   2940
      Tag             =   "1"
      Top             =   5745
      Width           =   2775
   End
   Begin VB.Image imgEditarURL 
      Height          =   495
      Left            =   2940
      Tag             =   "1"
      Top             =   5235
      Width           =   2775
   End
   Begin VB.Image imgEditarCodex 
      Height          =   495
      Left            =   2940
      Tag             =   "1"
      Top             =   4725
      Width           =   2775
   End
   Begin VB.Image imgActualizar 
      Height          =   390
      Left            =   180
      Tag             =   "1"
      Top             =   4230
      Width           =   5550
   End
   Begin VB.Image imgDetallesSolicitudes 
      Height          =   375
      Left            =   165
      Tag             =   "1"
      Top             =   6045
      Width           =   2655
   End
   Begin VB.Image imgDetallesMiembros 
      Height          =   375
      Left            =   3045
      Tag             =   "1"
      Top             =   2700
      Width           =   2655
   End
   Begin VB.Image imgDetallesClan 
      Height          =   375
      Left            =   150
      Tag             =   "1"
      Top             =   2700
      Width           =   2655
   End
   Begin VB.Image imgElecciones 
      Height          =   375
      Left            =   165
      Tag             =   "1"
      Top             =   6885
      Width           =   2655
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1815
      TabIndex        =   3
      Top             =   6510
      Width           =   255
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Private Const MAX_NEWS_LENGTH     As Integer = 512

Private cBotonElecciones          As clsGraphicalButton

Private cBotonActualizar          As clsGraphicalButton

Private cBotonDetallesClan        As clsGraphicalButton

Private cBotonDetallesMiembros    As clsGraphicalButton

Private cBotonDetallesSolicitudes As clsGraphicalButton

Private cBotonEditarCodex         As clsGraphicalButton

Private cBotonEditarURL           As clsGraphicalButton

Private cBotonPropuestasPaz       As clsGraphicalButton

Private cBotonPropuestasAlianzas  As clsGraphicalButton

Private cBotonCerrar              As clsGraphicalButton

Public LastPressed                As clsGraphicalButton

Private Sub Form_Load()
    'Call SetTranslucent(Me.hwnd, NTRANS_GENERAL)
    Me.Picture = LoadPictureEX("VentanaAdministrarClan.jpg")
    
    Call LoadButtons

End Sub

Private Sub LoadButtons()

    Set cBotonElecciones = New clsGraphicalButton
    Set cBotonActualizar = New clsGraphicalButton
    Set cBotonDetallesClan = New clsGraphicalButton
    Set cBotonDetallesMiembros = New clsGraphicalButton
    Set cBotonDetallesSolicitudes = New clsGraphicalButton
    Set cBotonEditarCodex = New clsGraphicalButton
    Set cBotonEditarURL = New clsGraphicalButton
    Set cBotonPropuestasPaz = New clsGraphicalButton
    Set cBotonPropuestasAlianzas = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    Call cBotonElecciones.Initialize(imgElecciones, "BotonElecciones.jpg", "BotonEleccionesRollover.jpg", "BotonEleccionesClick.jpg", Me)

    Call cBotonActualizar.Initialize(imgActualizar, "BotonActualizar.jpg", "BotonActualizarRollover.jpg", "BotonActualizarClick.jpg", Me)

    Call cBotonDetallesClan.Initialize(imgDetallesClan, "BotonDetallesAdministrarClan.jpg", "BotonDetallesRolloverAdministrarClan.jpg", "BotonDetallesClickAdministrarClan.jpg", Me)

    Call cBotonDetallesMiembros.Initialize(imgDetallesMiembros, "BotonDetallesAdministrarClan.jpg", "BotonDetallesRolloverAdministrarClan.jpg", "BotonDetallesClickAdministrarClan.jpg", Me)
                                    
    Call cBotonDetallesSolicitudes.Initialize(imgDetallesSolicitudes, "BotonDetallesAdministrarClan.jpg", "BotonDetallesRolloverAdministrarClan.jpg", "BotonDetallesClickAdministrarClan.jpg", Me)

    Call cBotonEditarCodex.Initialize(imgEditarCodex, "BotonEditarCodex.jpg", "BotonEditarCodexRollover.jpg", "BotonEditarCodexClick.jpg", Me)

    Call cBotonEditarURL.Initialize(imgEditarURL, "BotonEditarURL.jpg", "BotonEditarURLRollover.jpg", "BotonEditarURLClick.jpg", Me)

    Call cBotonPropuestasPaz.Initialize(imgPropuestasPaz, "BotonPropuestaPaz.jpg", "BotonPropuestaPazRollover.jpg", "BotonPropuestaPazClick.jpg", Me)

    Call cBotonPropuestasAlianzas.Initialize(imgPropuestasAlianzas, "BotonPropuestasAlianzas.jpg", "BotonPropuestasAlianzasRollover.jpg", "BotonPropuestasAlianzasClick.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, "BotonCerrarAdministrarClan.jpg", "BotonCerrarRolloverAdministrarClan.jpg", "BotonCerrarClickAdministrarClan.jpg", Me)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then MoverVentana (Me.hwnd)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub guildslist_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub imgActualizar_Click()

    Dim k As String

    k = Replace(txtguildnews, vbCrLf, "�")
    
    Call WriteGuildUpdateNews(k)

End Sub

Private Sub imgCerrar_Click()
    Unload Me
    frmMain.SetFocus

End Sub

Private Sub imgDetallesClan_Click()
    frmGuildBrief.EsLeader = True

    If guildslist.ListIndex = -1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("�Debes selecionar un Clan!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If

    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

End Sub

Private Sub imgDetallesMiembros_Click()

    If members.ListIndex = -1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("�Debes selecionar un Miembro!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))

End Sub

Private Sub imgDetallesSolicitudes_Click()

    If solicitudes.ListIndex = -1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("�Debes selecionar una solicitud!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))

End Sub

Private Sub imgEditarCodex_Click()
    Call frmGuildDetails.Show(vbModal, frmGuildLeader)

End Sub

Private Sub imgEditarURL_Click()
    Call frmGuildURL.Show(vbModeless, frmGuildLeader)

End Sub

Private Sub imgElecciones_Click()
    Call WriteGuildOpenElections
    Unload Me

End Sub

Private Sub imgPropuestasAlianzas_Click()
    Call WriteGuildAlliancePropList

End Sub

Private Sub imgPropuestasPaz_Click()
    Call WriteGuildPeacePropList

End Sub

Private Sub members_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub solicitudes_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub txtguildnews_Change()

    If Len(txtguildnews.Text) > MAX_NEWS_LENGTH Then txtguildnews.Text = Left$(txtguildnews.Text, MAX_NEWS_LENGTH)

End Sub

Private Sub txtguildnews_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub txtFiltrarClanes_Change()
    Call FiltrarListaClanes(txtFiltrarClanes.Text)

End Sub

Private Sub txtFiltrarClanes_GotFocus()

    With txtFiltrarClanes
        .SelStart = 0
        .SelLength = Len(.Text)

    End With

End Sub

Private Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long
    
    With guildslist
        'Limpio la lista
        .Clear
        
        .Visible = False
        
        ' Recorro los arrays
        For lIndex = 0 To UBound(GuildNames)

            ' Si coincide con los patrones
            If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                ' Lo agrego a la lista
                .AddItem GuildNames(lIndex)

            End If

        Next lIndex
        
        .Visible = True

    End With

End Sub

Private Sub txtFiltrarMiembros_Change()
    Call FiltrarListaMiembros(txtFiltrarMiembros.Text)

End Sub

Private Sub txtFiltrarMiembros_GotFocus()

    With txtFiltrarMiembros
        .SelStart = 0
        .SelLength = Len(.Text)

    End With

End Sub

Private Sub FiltrarListaMiembros(ByRef sCompare As String)

    Dim lIndex As Long
    
    With members
        'Limpio la lista
        .Clear
        
        .Visible = False
        
        ' Recorro los arrays
        For lIndex = 0 To UBound(GuildMembers)

            ' Si coincide con los patrones
            If InStr(1, UCase$(GuildMembers(lIndex)), UCase$(sCompare)) Then
                ' Lo agrego a la lista
                .AddItem GuildMembers(lIndex)

            End If

        Next lIndex
        
        .Visible = True

    End With

End Sub

