VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQuestInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   Picture         =   "frmQuestInfo.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   7680
      ScaleHeight     =   87
      ScaleMode       =   0  'User
      ScaleWidth      =   109
      TabIndex        =   9
      Top             =   4080
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   3860
      LinkItem        =   "detalle"
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   3320
   End
   Begin VB.ListBox lstQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      Left            =   10560
      TabIndex        =   6
      Top             =   8040
      Width           =   2835
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   10580
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   1830
      Width           =   480
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1560
      Left            =   7320
      TabIndex        =   1
      Top             =   1680
      Width           =   2230
      _ExtentX        =   3942
      _ExtentY        =   2752
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   14737632
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Criatura"
         Object.Width           =   3106
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Tipo"
         Object.Width           =   0
      EndProperty
      Picture         =   "frmQuestInfo.frx":11599
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2280
      Left            =   9840
      TabIndex        =   2
      Top             =   3000
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   4022
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      ForeColor       =   14737632
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Criatura"
         Object.Width           =   2294
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   1113
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Tipo"
         Object.Width           =   0
      EndProperty
      Picture         =   "frmQuestInfo.frx":11E6A
   End
   Begin MSComctlLib.ListView ListViewQuest 
      Height          =   2640
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   4657
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Quest"
         Object.Width           =   4588
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "estado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "id"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   11760
      Top             =   0
      Width           =   495
   End
   Begin VB.Label npclbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   3885
      Width           =   2295
   End
   Begin VB.Label objetolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   6400
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   3880
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Label titulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mision"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Width           =   3375
   End
End
Attribute VB_Name = "FrmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err

    Text1.BackColor = RGB(11, 11, 11)
   
    picture1.BackColor = RGB(19, 14, 11)
    
    Exit Sub

Form_Load_Err:
    
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err

    If (KeyAscii = 27) Then
        Unload Me

    End If
    
    Exit Sub

Form_KeyPress_Err:
    
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Image1.Picture = Nothing
    Image1.Tag = 0
    Image2.Picture = Nothing
    Image2.Tag = 0
    
    Exit Sub

Form_MouseMove_Err:
   
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    
    On Error GoTo Image1_MouseMove_Err

    If Image1.Tag = "0" Then
       
        Image1.Tag = "1"

    End If
    
    Exit Sub

Image1_MouseMove_Err:
 
    Resume Next
    
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Image1_MouseUp_Err
    
    Unload Me
    
    Exit Sub

Image1_MouseUp_Err:
   
    Resume Next
    
End Sub

Private Sub Image2_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    
    On Error GoTo Image2_MouseMove_Err

    If Image2.Tag = "0" Then
       
        Image2.Tag = "1"

    End If
    
    Exit Sub

Image2_MouseMove_Err:
  
    Resume Next
    
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Image2_MouseUp_Err

    If ListViewQuest.SelectedItem.Index > 0 Then
        Call WriteQuestAccept(ListViewQuest.SelectedItem.Index)
        Unload Me

    End If
    
    Exit Sub

Image2_MouseUp_Err:
    
    Resume Next
    
End Sub

Private Sub Image3_Click()
    
    On Error GoTo Image3_Click_Err
    
    Unload Me
    
    Exit Sub

Image3_Click_Err:
  
    Resume Next
    
End Sub

Public Sub ListView1_Click()

    On Error GoTo ListView1_Click_Err
    If ListView1.SelectedItem Is Nothing Then Exit Sub

    If ListView1.SelectedItem.SubItems(2) <> "" Then
        If ListView1.SelectedItem.SubItems(3) = 0 Then
            PlayerView.BackColor = RGB(11, 11, 11)
            Call DibujarBody(ListView1.SelectedItem.SubItems(2), 3)

            npclbl.Caption = NpcData(ListView1.SelectedItem.SubItems(2)).Name & " (" & ListView1.SelectedItem.SubItems(1) & ")"

        Else

            Dim X As Long

            Dim Y As Long

           ' X = (PlayerView.ScaleWidth - GrhData(ListView1.SelectedItem.SubItems(2)).PixelWidth) / 2
           ' Y = (PlayerView.ScaleHeight - GrhData(ListView1.SelectedItem.SubItems(2)).PixelHeight) / 2
            Call Grh_Render_To_Hdc(PlayerView, ObjData(ListView1.SelectedItem.SubItems(2)).GrhIndex, 40, 10, False, RGB(11, 11, 11))

            npclbl.Caption = ObjData(ListView1.SelectedItem.SubItems(2)).Name & " (" & ListView1.SelectedItem.SubItems(1) & ")"

        End If

    End If


    Exit Sub


ListView1_Click_Err:

    Resume Next

End Sub

Public Sub ListView2_Click()
    
    On Error GoTo ListView2_Click_Err

   If ListView2.SelectedItem.SubItems(2) <> "" Then
 
        Call Grh_Render_To_Hdc(picture1, ObjData(ListView2.SelectedItem.SubItems(2)).GrhIndex, 0, 0, False, RGB(19, 14, 11))
    
    End If
    
    objetolbl.Caption = ObjData(ListView2.SelectedItem.SubItems(2)).Name & vbCrLf & " (" & ListView2.SelectedItem.SubItems(1) & ")"

    
    Exit Sub

ListView2_Click_Err:
    
    Resume Next
    
End Sub

Private Sub ListViewQuest_ItemClick(ByVal item As MSComctlLib.ListItem)
    
    On Error GoTo ListViewQuest_ItemClick_Err
    
    If ListViewQuest.SelectedItem Is Nothing Then Exit Sub
    
    If Len(ListViewQuest.SelectedItem.SubItems(2)) <> 0 Then
        
        Dim QuestIndex As Byte

        QuestIndex = ListViewQuest.SelectedItem.SubItems(2)
        
        FrmQuestInfo.ListView2.ListItems.Clear
        FrmQuestInfo.ListView1.ListItems.Clear
            
        FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
        
        FrmQuestInfo.Text1.Text = ""
    
        picture1.BackColor = RGB(19, 14, 11)
        Imagen.Refresh
        picture1.Refresh
        npclbl.Caption = ""
        objetolbl.Caption = ""
        
        If QuestList(QuestIndex).RequiredQuest <> 0 Then
            FrmQuestInfo.Text1.Text = QuestList(QuestIndex).Desc & vbCrLf & vbCrLf & "Requisitos: " & vbCrLf & "Nivel requerido: " & QuestList(QuestIndex).RequiredLevel & vbCrLf & "Quest: " & QuestList(QuestList(QuestIndex).RequiredQuest).nombre
        Else
            FrmQuestInfo.Text1.Text = QuestList(QuestIndex).Desc & vbCrLf & vbCrLf & "Requisitos: " & vbCrLf & "Nivel requerido: " & QuestList(QuestIndex).RequiredLevel & vbCrLf
                
        End If
                
        If UBound(QuestList(QuestIndex).RequiredNPC) > 0 Then 'Hay NPCs
            If UBound(QuestList(QuestIndex).RequiredNPC) > 5 Then
                FrmQuestInfo.ListView1.FlatScrollBar = False
            Else
                FrmQuestInfo.ListView1.FlatScrollBar = True
               
            End If
                    
            For I = 1 To UBound(QuestList(QuestIndex).RequiredNPC)

                Dim subelemento As ListItem
    
                Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(QuestList(QuestIndex).RequiredNPC(I).NpcIndex).Name)
                           
                subelemento.SubItems(1) = QuestList(QuestIndex).RequiredNPC(I).Amount
                subelemento.SubItems(2) = QuestList(QuestIndex).RequiredNPC(I).NpcIndex
                subelemento.SubItems(3) = 0
    
            Next I
    
        End If
    
        If LBound(QuestList(QuestIndex).RequiredOBJ) > 0 Then  'Hay OBJs
    
            For I = 1 To UBound(QuestList(QuestIndex).RequiredOBJ)
                Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(QuestList(QuestIndex).RequiredOBJ(I).OBJIndex).Name)
                subelemento.SubItems(1) = QuestList(QuestIndex).RequiredOBJ(I).Amount
                subelemento.SubItems(2) = QuestList(QuestIndex).RequiredOBJ(I).OBJIndex
                subelemento.SubItems(3) = 1
            Next I
    
        End If
               
        If QuestList(QuestIndex).RewardGLD <> 0 Then
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")
            subelemento.SubItems(1) = QuestList(QuestIndex).RewardGLD
            subelemento.SubItems(2) = 12
            subelemento.SubItems(3) = 0

        End If
                
        If QuestList(QuestIndex).RewardEXP <> 0 Then
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
            subelemento.SubItems(1) = QuestList(QuestIndex).RewardEXP
            subelemento.SubItems(2) = 608
            subelemento.SubItems(3) = 1

        End If

        If UBound(QuestList(QuestIndex).RewardOBJ) > 0 Then
                    
            For I = 1 To UBound(QuestList(QuestIndex).RewardOBJ)
                                                                   
                Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(QuestList(QuestIndex).RewardOBJ(I).OBJIndex).Name)
                           
                subelemento.SubItems(1) = QuestList(QuestIndex).RewardOBJ(I).Amount
                subelemento.SubItems(2) = QuestList(QuestIndex).RewardOBJ(I).OBJIndex
                subelemento.SubItems(3) = 1
               
            Next I
    
        End If
                
        Call ListView1_Click
        Call ListView2_Click

    End If
    
    Exit Sub

ListViewQuest_ItemClick_Err:
   
    Resume Next
    
End Sub

Private Sub lstQuests_Click()
    
    On Error GoTo lstQuests_Click_Err
    
    Dim QuestIndex As Byte

    QuestIndex = Val(ReadField(1, lstQuests.List(lstQuests.ListIndex), Asc("-")))

    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear
            
    FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
                
    FrmQuestInfo.Text1.Text = QuestList(QuestIndex).Desc & vbCrLf & "Nivel requerido: " & QuestList(QuestIndex).RequiredLevel & vbCrLf
                
    If UBound(QuestList(QuestIndex).RequiredNPC) > 0 Then 'Hay NPCs
        If UBound(QuestList(QuestIndex).RequiredNPC) > 5 Then
            FrmQuestInfo.ListView1.FlatScrollBar = False
        Else
            FrmQuestInfo.ListView1.FlatScrollBar = True
               
        End If
                    
        For I = 1 To UBound(QuestList(QuestIndex).RequiredNPC)

            Dim subelemento As ListItem
    
            Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(QuestList(QuestIndex).RequiredNPC(I).NpcIndex).Name)
                           
            subelemento.SubItems(1) = QuestList(QuestIndex).RequiredNPC(I).Amount
            subelemento.SubItems(2) = QuestList(QuestIndex).RequiredNPC(I).NpcIndex
            subelemento.SubItems(3) = 0
    
        Next I
    
    End If
    
    If LBound(QuestList(QuestIndex).RequiredOBJ) > 0 Then  'Hay OBJs
    
        For I = 1 To UBound(QuestList(QuestIndex).RequiredOBJ)
            Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(QuestList(QuestIndex).RequiredOBJ(I).OBJIndex).Name)
            subelemento.SubItems(1) = QuestList(QuestIndex).RequiredOBJ(I).Amount
            subelemento.SubItems(2) = QuestList(QuestIndex).RequiredOBJ(I).OBJIndex
            subelemento.SubItems(3) = 1
        Next I
    
    End If
               
    Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")
                           
    subelemento.SubItems(1) = QuestList(QuestIndex).RewardGLD
    subelemento.SubItems(2) = 12
    subelemento.SubItems(3) = 0
               
    Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                           
    subelemento.SubItems(1) = QuestList(QuestIndex).RewardEXP
    subelemento.SubItems(2) = 608
    subelemento.SubItems(3) = 1

    If UBound(QuestList(QuestIndex).RewardOBJ) > 0 Then
                    
        For I = 1 To UBound(QuestList(QuestIndex).RewardOBJ)
                                                                   
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(QuestList(QuestIndex).RewardOBJ(I).OBJIndex).Name)
                           
            subelemento.SubItems(1) = QuestList(QuestIndex).RewardOBJ(I).Amount
            subelemento.SubItems(2) = QuestList(QuestIndex).RewardOBJ(I).OBJIndex
            subelemento.SubItems(3) = 1
               
        Next I
    
    End If
    
    Exit Sub

lstQuests_Click_Err:

    Resume Next
    
End Sub


Sub DibujarBody(ByVal MyBody As Integer, Optional ByVal Heading As Byte = 3)
    
    On Error GoTo DibujarBody_Err
    

    

    Dim Grh As Grh

    Grh = BodyData(NpcData(MyBody).Body).Walk(3)

    Dim X    As Long

    Dim Y    As Long

    Dim grhH As Grh

    grhH = HeadData(NpcData(MyBody).Head).Head(3)

   X = (PlayerView.ScaleWidth - GrhData(Grh.GrhIndex).PixelWidth) / 2
    Y = (PlayerView.ScaleHeight - GrhData(Grh.GrhIndex).PixelHeight) / 2
     Call Grh_Render_To_Hdc(PlayerView, GrhData(Grh.GrhIndex).Frames(1), X, Y + 15, False, RGB(11, 11, 11))
    

    If NpcData(MyBody).Head <> 0 Then
        X = (PlayerView.ScaleWidth - GrhData(grhH.GrhIndex).PixelWidth) / 2
        Y = (PlayerView.ScaleHeight - GrhData(grhH.GrhIndex).PixelHeight) / 2 + 8 + BodyData(NpcData(MyBody).Body).HeadOffset.Y
        PlayerView.BackColor = RGB(11, 11, 11)
        Call Grh_Render_To_HdcSinBorrar(PlayerView, GrhData(grhH.GrhIndex).Frames(1), 42, Y - 8, False)


    End If

    
    Exit Sub

DibujarBody_Err:
    
    Resume Next
    
End Sub

