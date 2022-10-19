VERSION 5.00
Begin VB.Form frmCarp 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtCantItems 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Ingrese la cantidad total de items a construir."
      Top             =   2880
      Width           =   780
   End
   Begin VB.ComboBox cboItemsCiclo 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5355
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4215
      Width           =   735
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMaderas0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.VScrollBar Scroll 
      Height          =   3135
      Left            =   405
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMaderas1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   2355
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMaderas2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   3150
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMaderas3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   3945
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMejorar0 
      Height          =   420
      Left            =   3150
      Top             =   1560
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgCantidadCiclo 
      Height          =   645
      Left            =   5160
      Top             =   3435
      Width           =   1110
   End
   Begin VB.Image imgPestania 
      Height          =   255
      Index           =   1
      Left            =   1680
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image imgPestania 
      Height          =   255
      Index           =   0
      Left            =   720
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   975
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   4
      Left            =   5280
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   3
      Left            =   5280
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   2
      Left            =   5280
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   1
      Left            =   5280
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   4
      Left            =   1560
      Top             =   3780
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   3
      Left            =   1560
      Top             =   2985
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   2
      Left            =   1560
      Top             =   2190
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   1
      Left            =   1560
      Top             =   1395
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   4
      Left            =   720
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   3
      Left            =   720
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   2
      Left            =   720
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   1
      Left            =   720
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   4800
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Image imgConstruir3 
      Height          =   420
      Left            =   3150
      Top             =   3960
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgConstruir2 
      Height          =   420
      Left            =   3150
      Top             =   3180
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgConstruir1 
      Height          =   420
      Left            =   3150
      Top             =   2370
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgConstruir0 
      Height          =   420
      Left            =   3120
      Top             =   1560
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgChkMacro 
      Height          =   420
      Left            =   5445
      MousePointer    =   99  'Custom
      Top             =   1875
      Width           =   435
   End
   Begin VB.Image imgMejorar1 
      Height          =   420
      Left            =   3150
      Top             =   2370
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgMejorar2 
      Height          =   420
      Left            =   3150
      Top             =   3180
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgMejorar3 
      Height          =   420
      Left            =   3150
      Top             =   3960
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "frmCarp"
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

Option Explicit

Dim Cargando As Boolean

Private Enum ePestania

    ieItems
    ieMejorar

End Enum

Private picCheck                As Picture

Private picRecuadroItem         As Picture

Private picRecuadroMaderas      As Picture

Private Pestanias(1)            As Picture

Private UltimaPestania          As Byte

Private cBotonCerrar            As clsGraphicalButton

Private cBotonConstruir(0 To 4) As clsGraphicalButton

Private cBotonMejorar(0 To 4)   As clsGraphicalButton

Public LastPressed              As clsGraphicalButton

Private UsarMacro               As Boolean

Private Sub Form_Load()
    'Call SetTranslucent(Me.hwnd, NTRANS_GENERAL)
    
    Call LoadDefaultValues
    
    Me.Picture = LoadPictureEX("VentanaCarpinteriaItems.jpg")
    LoadButtons

End Sub

Private Sub LoadButtons()

    Dim Index As Long

    Set Pestanias(ePestania.ieItems) = LoadPictureEX("VentanaCarpinteriaItems.jpg")
    Set Pestanias(ePestania.ieMejorar) = LoadPictureEX("VentanaCarpinteriaMejorar.jpg")
    
    Set picCheck = LoadPictureEX("CheckBoxCarpinteria.jpg")
    
    Set picRecuadroItem = LoadPictureEX("RecuadroItemsCarpinteria.jpg")
    Set picRecuadroMaderas = LoadPictureEX("RecuadroMadera.jpg")
    
    For Index = 1 To MAX_LIST_ITEMS
        imgMarcoItem(Index).Picture = picRecuadroItem
        imgMarcoUpgrade(Index).Picture = picRecuadroItem
        imgMarcoMaderas(Index).Picture = picRecuadroMaderas
    Next Index
    
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonConstruir(0) = New clsGraphicalButton
    Set cBotonConstruir(1) = New clsGraphicalButton
    Set cBotonConstruir(2) = New clsGraphicalButton
    Set cBotonConstruir(3) = New clsGraphicalButton
    Set cBotonMejorar(0) = New clsGraphicalButton
    Set cBotonMejorar(1) = New clsGraphicalButton
    Set cBotonMejorar(2) = New clsGraphicalButton
    Set cBotonMejorar(3) = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(imgCerrar, "BotonCerrarCarpinteria.jpg", "BotonCerrarRolloverCarpinteria.jpg", "BotonCerrarClickCarpinteria.jpg", Me)
                                    
    Call cBotonConstruir(0).Initialize(imgConstruir0, "BotonConstruirCarpinteria.jpg", "BotonConstruirRolloverCarpinteria.jpg", "BotonConstruirClickCarpinteria.jpg", Me)
                                    
    Call cBotonConstruir(1).Initialize(imgConstruir1, "BotonConstruirCarpinteria.jpg", "BotonConstruirRolloverCarpinteria.jpg", "BotonConstruirClickCarpinteria.jpg", Me)
                                    
    Call cBotonConstruir(2).Initialize(imgConstruir2, "BotonConstruirCarpinteria.jpg", "BotonConstruirRolloverCarpinteria.jpg", "BotonConstruirClickCarpinteria.jpg", Me)
                                    
    Call cBotonConstruir(3).Initialize(imgConstruir3, "BotonConstruirCarpinteria.jpg", "BotonConstruirRolloverCarpinteria.jpg", "BotonConstruirClickCarpinteria.jpg", Me)
    
    Call cBotonMejorar(0).Initialize(imgMejorar0, "BotonMejorarCarpinteria.jpg", "BotonMejorarRolloverCarpinteria.jpg", "BotonMejorarClickCarpinteria.jpg", Me)
    
    Call cBotonMejorar(1).Initialize(imgMejorar1, "BotonMejorarCarpinteria.jpg", "BotonMejorarRolloverCarpinteria.jpg", "BotonMejorarClickCarpinteria.jpg", Me)
    
    Call cBotonMejorar(2).Initialize(imgMejorar2, "BotonMejorarCarpinteria.jpg", "BotonMejorarRolloverCarpinteria.jpg", "BotonMejorarClickCarpinteria.jpg", Me)
    
    Call cBotonMejorar(3).Initialize(imgMejorar3, "BotonMejorarCarpinteria.jpg", "BotonMejorarRolloverCarpinteria.jpg", "BotonMejorarClickCarpinteria.jpg", Me)
                                    
    imgCantidadCiclo.Picture = LoadPictureEX("ConstruirPorCiclo.jpg")
    
    imgChkMacro.Picture = picCheck
    
    imgPestania(ePestania.ieItems).MouseIcon = picMouseIcon
    imgPestania(ePestania.ieMejorar).MouseIcon = picMouseIcon
    
    imgChkMacro.MouseIcon = picMouseIcon

End Sub

Private Sub LoadDefaultValues()
    
    Dim MaxConstItem As Integer

    Dim I            As Integer

    Cargando = True
    
    MaxConstItem = CInt((UserLvl - 4) / 5)
    MaxConstItem = IIf(MaxConstItem < 1, 1, MaxConstItem)
    
    For I = 1 To MaxConstItem
        cboItemsCiclo.AddItem I
    Next I
    
    cboItemsCiclo.ListIndex = 0
    
    Scroll.Value = 0
    
    UsarMacro = True
    
    UltimaPestania = ePestania.ieItems
    
    Cargando = False

End Sub

Private Sub Construir(ByVal Index As Integer)

    Dim ItemIndex      As Integer

    Dim CantItemsCiclo As Integer
    
    If Scroll.Visible = True Then ItemIndex = Scroll.Value
    ItemIndex = ItemIndex + Index
    
    Select Case UltimaPestania

        Case ePestania.ieItems
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ObjCarpintero(ItemIndex).OBJIndex
                frmMain.ActivarMacroTrabajo
            Else
                ' Que cosntruya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))

            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftCarpenter(ObjCarpintero(ItemIndex).OBJIndex)
            
        Case ePestania.ieMejorar
            Call WriteItemUpgrade(CarpinteroMejorar(ItemIndex).OBJIndex)

    End Select
        
    Unload Me

End Sub

Public Sub HideExtraControls(ByVal NumItems As Integer, _
                             Optional ByVal Upgrading As Boolean = False)

    Dim I As Integer
    
    picMaderas0.Visible = (NumItems >= 1)
    picMaderas1.Visible = (NumItems >= 2)
    picMaderas2.Visible = (NumItems >= 3)
    picMaderas3.Visible = (NumItems >= 4)
    
    imgConstruir0.Visible = (NumItems >= 1 And Not Upgrading)
    imgConstruir1.Visible = (NumItems >= 2 And Not Upgrading)
    imgConstruir2.Visible = (NumItems >= 3 And Not Upgrading)
    imgConstruir3.Visible = (NumItems >= 4 And Not Upgrading)
    
    imgMejorar0.Visible = (NumItems >= 1 And Upgrading)
    imgMejorar1.Visible = (NumItems >= 2 And Upgrading)
    imgMejorar2.Visible = (NumItems >= 3 And Upgrading)
    imgMejorar3.Visible = (NumItems >= 4 And Upgrading)
    
    For I = 1 To MAX_LIST_ITEMS
        picItem(I).Visible = (NumItems >= I)
        imgMarcoItem(I).Visible = (NumItems >= I)
        imgMarcoMaderas(I).Visible = (NumItems >= I)

        ' Upgrade
        imgMarcoUpgrade(I).Visible = (NumItems >= I And Upgrading)
        picUpgrade(I).Visible = (NumItems >= I And Upgrading)
    Next I
    
    If NumItems > MAX_LIST_ITEMS Then
        Scroll.Visible = True
        Cargando = True
        Scroll.max = NumItems - MAX_LIST_ITEMS
        Cargando = False
    Else
        Scroll.Visible = False

    End If
    
    txtCantItems.Visible = Not Upgrading
    cboItemsCiclo.Visible = Not Upgrading And UsarMacro
    imgChkMacro.Visible = Not Upgrading
    imgCantidadCiclo.Visible = Not Upgrading And UsarMacro

End Sub

Private Sub RenderItem(ByRef Pic As PictureBox, ByVal GrhIndex As Long)

    Dim SR As RECT

    Dim DR As RECT

    If GrhIndex > 0 Then

        With GrhData(GrhIndex)
            SR.Left = .sX
            SR.Top = .sY
            SR.Right = SR.Left + .PixelWidth
            SR.Bottom = SR.Top + .PixelHeight

        End With
    
        DR.Left = 0
        DR.Top = 0
        DR.Right = 32
        DR.Bottom = 32
        Pic.Cls
        Call DrawGrhtoHdc(Pic.hDC, GrhIndex, SR, DR)
        Pic.Refresh

    End If

End Sub

Public Sub RenderList(ByVal Inicio As Integer)

    Dim I        As Long

    Dim NumItems As Integer

    NumItems = UBound(ObjCarpintero)
    Inicio = Inicio - 1

    For I = 1 To MAX_LIST_ITEMS

        If I + Inicio <= NumItems Then

            With ObjCarpintero(I + Inicio)
                ' Agrego el item
                Call RenderItem(picItem(I), .GrhIndex)
                picItem(I).ToolTipText = .Name
        
                ' Inventario de leños
                Call InvMaderasCarpinteria(I).SetItem(1, 0, .Madera, 0, MADERA_GRH, 0, 0, 0, 0, 0, 0, "Leña", 0)
                Call InvMaderasCarpinteria(I).SetItem(2, 0, .MaderaElfica, 0, MADERA_ELFICA_GRH, 0, 0, 0, 0, 0, 0, "Leña élfica", 0)

            End With

        End If

    Next I

End Sub

Public Sub RenderUpgradeList(ByVal Inicio As Integer)

    Dim I        As Long

    Dim NumItems As Integer

    NumItems = UBound(CarpinteroMejorar)
    Inicio = Inicio - 1

    For I = 1 To MAX_LIST_ITEMS

        If I + Inicio <= NumItems Then

            With CarpinteroMejorar(I + Inicio)
                ' Agrego el item
                Call RenderItem(picItem(I), .GrhIndex)
                picItem(I).ToolTipText = .Name
            
                Call RenderItem(picUpgrade(I), .UpgradeGrhIndex)
                picUpgrade(I).ToolTipText = .UpgradeName
        
                ' Inventario de leños
                Call InvMaderasCarpinteria(I).SetItem(1, 0, .Madera, 0, MADERA_GRH, 0, 0, 0, 0, 0, 0, "Leña", 0)
                Call InvMaderasCarpinteria(I).SetItem(2, 0, .MaderaElfica, 0, MADERA_ELFICA_GRH, 0, 0, 0, 0, 0, 0, "Leña élfica", 0)

            End With

        End If

    Next I

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then MoverVentana (Me.hwnd)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub imgCerrar_Click()
    Unload Me

End Sub

Private Sub imgChkMacro_Click()
    UsarMacro = Not UsarMacro
    
    If UsarMacro Then
        imgChkMacro.Picture = picCheck
    Else
        Set imgChkMacro.Picture = Nothing

    End If
    
    cboItemsCiclo.Visible = UsarMacro
    imgCantidadCiclo.Visible = UsarMacro

End Sub

Private Sub imgConstruir0_Click()
    Call Construir(1)

End Sub

Private Sub imgConstruir1_Click()
    Call Construir(2)

End Sub

Private Sub imgConstruir2_Click()
    Call Construir(3)

End Sub

Private Sub imgConstruir3_Click()
    Call Construir(4)

End Sub

Private Sub imgMejorar0_Click()
    Call Construir(1)

End Sub

Private Sub imgMejorar1_Click()
    Call Construir(2)

End Sub

Private Sub imgMejorar2_Click()
    Call Construir(3)

End Sub

Private Sub imgMejorar3_Click()
    Call Construir(4)

End Sub

Private Sub imgPestania_Click(Index As Integer)

    Dim I        As Integer

    Dim NumItems As Integer
    
    If Cargando Then Exit Sub
    If UltimaPestania = Index Then Exit Sub
    
    Scroll.Value = 0
    
    Select Case Index

        Case ePestania.ieItems
            ' Background
            Me.Picture = Pestanias(ePestania.ieItems)
            
            NumItems = UBound(ObjCarpintero)
        
            Call HideExtraControls(NumItems)
            
            ' Cargo inventarios e imagenes
            Call RenderList(1)

        Case ePestania.ieMejorar
            ' Background
            Me.Picture = Pestanias(ePestania.ieMejorar)
            
            NumItems = UBound(CarpinteroMejorar)
            
            Call HideExtraControls(NumItems, True)
            
            Call RenderUpgradeList(1)

    End Select

    UltimaPestania = Index

End Sub

Private Sub Scroll_Change()

    Dim I As Long
    
    If Cargando Then Exit Sub
    
    I = Scroll.Value
    ' Cargo inventarios e imagenes
    
    Select Case UltimaPestania

        Case ePestania.ieItems
            Call RenderList(I + 1)

        Case ePestania.ieMejorar
            Call RenderUpgradeList(I + 1)

    End Select

End Sub
