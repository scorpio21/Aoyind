VERSION 5.00
Begin VB.Form frmPanelTorneo 
   BackColor       =   &H00000000&
   Caption         =   "TORNEOS"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form4"
   ScaleHeight     =   6540
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Información de los eventos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6480
      Left            =   6000
      TabIndex        =   30
      Top             =   0
      Width           =   4890
      Begin VB.Frame Frame3 
         Caption         =   "Información del evento seleccionado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   390
         TabIndex        =   35
         Top             =   2400
         Width           =   4110
         Begin VB.CommandButton Command2 
            Caption         =   "DESCALIFICAR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1080
            TabIndex        =   41
            Top             =   3360
            Width           =   1965
         End
         Begin VB.ListBox lstUsers 
            Height          =   840
            Left            =   195
            TabIndex        =   40
            Top             =   2340
            Width           =   3720
         End
         Begin VB.Label lblCanjeCurso 
            BackStyle       =   0  'Transparent
            Caption         =   "Canje- Poso acumulado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label lblUsers 
            Caption         =   "Usuarios inscriptos y DISPONIBLES:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   42
            Top             =   1920
            Width           =   3330
         End
         Begin VB.Label lblDspCurso 
            Caption         =   "DSP- Poso acumulado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   39
            Top             =   1200
            Width           =   3330
         End
         Begin VB.Label lblOroCurso 
            Caption         =   "ORO- Poso acumulado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   38
            Top             =   960
            Width           =   3330
         End
         Begin VB.Label lblNivelCurso 
            Caption         =   "Nivel mínimo/máximo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   37
            Top             =   600
            Width           =   3330
         End
         Begin VB.Label lblQuotasCurso 
            Caption         =   "Cupos:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   36
            Top             =   240
            Width           =   3330
         End
      End
      Begin VB.ComboBox cmbModalityCurso 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2145
         TabIndex        =   33
         Text            =   "Vacio"
         Top             =   1695
         Width           =   2355
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Solicitar eventos en curso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   585
         TabIndex        =   32
         Top             =   330
         Width           =   3915
      End
      Begin VB.Label lblClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   4550
         TabIndex        =   43
         Top             =   1680
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label11 
         Caption         =   "Una vez que aparezca la lista de los eventos que hay disponibles (en curso) al seleccionar uno se actualizarán sus datos."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   600
         Left            =   195
         TabIndex        =   34
         Top             =   975
         Width           =   4305
      End
      Begin VB.Label Label10 
         Caption         =   "Eventos disponibles:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   31
         Top             =   1755
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmbCrear 
      Caption         =   "Crear nuevo torneo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   5760
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo torneo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6060
      Begin VB.CheckBox chkClass 
         Caption         =   "Ladrón"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   4440
         TabIndex        =   50
         Top             =   3960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Ladrón"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4440
         TabIndex        =   49
         Top             =   3600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbCanje 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   47
         Text            =   "Vacio"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox cmbTeam 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   45
         Text            =   "Vacio"
         Top             =   4440
         Width           =   1185
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Pirata"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   4440
         TabIndex        =   29
         Top             =   4440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   4440
         TabIndex        =   28
         Top             =   4680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Cazador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   3360
         TabIndex        =   27
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Paladin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   3360
         TabIndex        =   26
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Druida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   3360
         TabIndex        =   25
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Bardo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   3360
         TabIndex        =   24
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Ladrón"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   4440
         TabIndex        =   23
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Asesino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   3360
         TabIndex        =   22
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Guerrero"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   3360
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Clerigo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3360
         TabIndex        =   20
         Top             =   840
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkClass 
         Caption         =   "Mago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3360
         TabIndex        =   19
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.ComboBox cmbInit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   17
         Text            =   "Vacio"
         Top             =   3840
         Width           =   1185
      End
      Begin VB.ComboBox cmbCancel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         Text            =   "Vacio"
         Top             =   3360
         Width           =   1185
      End
      Begin VB.ComboBox cmbDsp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   15
         Text            =   "Vacio"
         Top             =   2280
         Width           =   1185
      End
      Begin VB.ComboBox cmbOro 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   14
         Text            =   "Vacio"
         Top             =   1880
         Width           =   1185
      End
      Begin VB.ComboBox cmbLvlMax 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   13
         Text            =   "Vacio"
         Top             =   1500
         Width           =   1185
      End
      Begin VB.ComboBox cmbLvlMin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   12
         Text            =   "Vacio"
         Top             =   1120
         Width           =   1185
      End
      Begin VB.ComboBox cmbQuotas 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   11
         Text            =   "Vacio"
         Top             =   730
         Width           =   1185
      End
      Begin VB.ComboBox cmbModality 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   10
         Text            =   "Vacio"
         Top             =   350
         Width           =   1185
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Canje inscripción:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   246
         TabIndex        =   46
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Cantidad por team:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   44
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Clases permitidas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3360
         TabIndex        =   18
         Top             =   390
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   3315
         X2              =   3315
         Y1              =   390
         Y2              =   3705
      End
      Begin VB.Label Label8 
         Caption         =   "Las inscripciones comienzan en:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   8
         Top             =   3720
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "Tiempo para cancelar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   "Dsp inscripción:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   6
         Top             =   2340
         Width           =   1380
      End
      Begin VB.Label Label5 
         Caption         =   "Oro inscripción:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   5
         Top             =   1950
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Nivel máximo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   4
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Nivel mínimo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   3
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Cupos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   2
         Top             =   780
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Modalidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   1
         Top             =   390
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmPanelTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCrear_Click()
    
    Dim AllowedClasses(1 To NUMCLASES) As Byte

    Dim LoopC                          As Byte
    
    If CheckForm(AllowedClasses) Then

        For LoopC = 1 To NUMCLASES
            AllowedClasses(LoopC) = Val(chkClass(LoopC).Value)
        Next LoopC
        
        Debug.Print "Modality: " & Val(cmbModality.ListIndex + 1)
        
        WriteNewEvent Val(cmbModality.ListIndex + 1), Val(cmbQuotas.List(cmbQuotas.ListIndex)), Val(cmbLvlMin.List(cmbLvlMin.ListIndex)), Val(cmbLvlMax.List(cmbLvlMax.ListIndex)), Val(cmbOro.List(cmbOro.ListIndex)), Val(cmbDsp.List(cmbDsp.ListIndex)), Val(cmbCanje.List(cmbCanje.ListIndex)), Val(cmbInit.List(cmbInit.ListIndex)) * 60, Val(cmbCancel.List(cmbCancel.ListIndex)) * 60, Val(cmbTeam.List(cmbTeam.ListIndex)), AllowedClasses()

    End If
    
End Sub

Private Function CheckForm(ByRef AllowedClasses() As Byte) As Boolean
    CheckForm = False
    
    If cmbTeam.Text = "Vacio" Then
        MsgBox "Seleccione la cantidad de pjs por team"
        Exit Function

    End If
    
    If cmbModality.Text = "Vacio" Then
        MsgBox "Seleccione la modalidad del evento"
        Exit Function

    End If
    
    If cmbQuotas.Text = "Vacio" Then
        MsgBox "Seleccione la cantidad de cupos que tendra"
        Exit Function

    End If
    
    If cmbLvlMin.Text = "Vacio" Then
        MsgBox "Seleccione el nivel minimo que requerirá el evento"
        Exit Function

    End If
    
    If cmbLvlMax.Text = "Vacio" Then
        MsgBox "Seleccione el nivel máximo que requerirá el evento"
        Exit Function

    End If
    
    If cmbOro.Text = "Vacio" Then
        MsgBox "Seleccione la inscripción por ORO"
        Exit Function

    End If
    
    If cmbDsp.Text = "Vacio" Then
        MsgBox "Seleccione la inscripción por DSP"
        Exit Function

    End If
    
    If cmbCanje.Text = "Vacio" Then
        MsgBox "Seleccione la inscripción por CANJES"
        Exit Function

    End If
    
    If cmbInit.Text = "Vacio" Then
        MsgBox "Seleccione el tiempo que tardará en iniciar las incripciones"
        Exit Function

    End If
    
    If cmbCancel.Text = "Vacio" Then
        MsgBox "Seleccione el tiempo que tendrá para cancelarse el evento si no se completan cupos"
        Exit Function

    End If
    
    Dim LoopC As Integer, Puede As Boolean
    
    For LoopC = 1 To NUMCLASES

        If AllowedClasses(LoopC) = 1 Then
            Puede = True
            Exit For

        End If

    Next LoopC
        
    CheckForm = True

End Function

Private Sub cmbModalityCurso_Click()

    If cmbModalityCurso.ListIndex = -1 Then Exit Sub
    
    lblClose.Visible = IIf(cmbModalityCurso.List(cmbModalityCurso.ListIndex) <> "Vacio", True, False)
    
    If cmbModalityCurso.List(cmbModalityCurso.ListIndex) = "Vacio" Then
        MsgBox "No se puede ver la información del evento que seleccionaste."
        Exit Sub

    End If
    
    Protocol.WriteRequiredDataEvent cmbModalityCurso.ListIndex + 1

End Sub

Private Sub Command1_Click()
    Protocol.WriteRequiredEvents

End Sub

Private Sub Form_Load()
    
    Dim LoopC As Integer
    
    cmbModality.AddItem "Castle Mode"
    cmbModality.AddItem "DagaRusa"
    cmbModality.AddItem "DeathMatch"
    cmbModality.AddItem "Enfrentamientos"
    
    For LoopC = 2 To 64
        cmbQuotas.AddItem LoopC
    Next LoopC
    
    For LoopC = 0 To 47
        cmbLvlMin.AddItem LoopC
        cmbLvlMax.AddItem LoopC
    Next LoopC
    
    cmbLvlMin.ListIndex = 0
    cmbLvlMax.ListIndex = 0
    
    cmbOro.AddItem "0"
    cmbOro.AddItem "25000"
    cmbOro.AddItem "50000"
    cmbOro.AddItem "100000"
    cmbOro.AddItem "200000"
    cmbOro.AddItem "300000"
    cmbOro.AddItem "400000"
    cmbOro.AddItem "500000"
    cmbOro.AddItem "600000"
    cmbOro.AddItem "700000"
    cmbOro.AddItem "800000"
    cmbOro.AddItem "900000"
    cmbOro.AddItem "1000000"
    cmbOro.ListIndex = 0
    
    cmbDsp.AddItem "0"
    cmbDsp.AddItem "1"
    cmbDsp.AddItem "2"
    cmbDsp.AddItem "5"
    cmbDsp.AddItem "10"
    cmbDsp.AddItem "15"
    cmbDsp.AddItem "20"
    cmbDsp.AddItem "25"
    cmbDsp.AddItem "30"
    cmbDsp.AddItem "35"
    cmbDsp.AddItem "40"
    cmbDsp.AddItem "45"
    cmbDsp.AddItem "50"
    
    cmbCanje.AddItem "0"
    cmbCanje.AddItem "2"
    cmbCanje.AddItem "4"
    cmbCanje.AddItem "8"
    cmbCanje.AddItem "16"
    cmbCanje.AddItem "32"
    cmbCanje.AddItem "64"
    cmbCanje.AddItem "128"
    cmbCanje.AddItem "256"
    cmbCanje.ListIndex = 0
    
    cmbTeam.AddItem 1
    cmbTeam.AddItem 2
    cmbTeam.AddItem 3
    cmbTeam.AddItem 4
    cmbTeam.AddItem 5
    cmbTeam.AddItem 6
    cmbTeam.AddItem 7
    cmbTeam.AddItem 8
    cmbTeam.AddItem 9
    cmbTeam.AddItem 10
    
    cmbDsp.ListIndex = 0
    
    For LoopC = 1 To 10
        cmbCancel.AddItem LoopC
        cmbInit.AddItem LoopC
    Next LoopC
    
    cmbCancel.ListIndex = 7
    cmbInit.ListIndex = 0
    
End Sub

Private Sub lblClose_Click()
    Protocol.WriteCloseEvent cmbModalityCurso.ListIndex + 1

End Sub

