VERSION 5.00
Begin VB.Form frmCerrar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Opcion 
      Height          =   405
      Index           =   2
      Left            =   630
      Top             =   1740
      Width           =   1980
   End
   Begin VB.Image Opcion 
      Height          =   405
      Index           =   1
      Left            =   630
      Top             =   1185
      Width           =   1980
   End
   Begin VB.Image Opcion 
      Height          =   405
      Index           =   0
      Left            =   630
      Top             =   615
      Width           =   1980
   End
End
Attribute VB_Name = "frmCerrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
'Declaraci�n del Api SetLayeredWindowAttributes que establece _
 la transparencia al form
  
Private Declare Function SetLayeredWindowAttributes _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal crKey As Long, _
                              ByVal bAlpha As Byte, _
                              ByVal dwFlags As Long) As Long
  
'Recupera el estilo de la ventana
Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long
  
'Declaraci�n del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes
  
Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long
  
Private Const GWL_EXSTYLE = (-20)

Private Const LWA_ALPHA = &H2

Private Const WS_EX_LAYERED = &H80000

'Funci�n para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuesti�n
 
Public bmoving      As Boolean

Public dX           As Integer

Public dy           As Integer

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Long) As Long

Private RealizoCambios As String

Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1

Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Sub moverForm()
    
    On Error GoTo moverForm_Err

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
    
    Exit Sub

moverForm_Err:
    Call LogError(err.Number, err.Description, "frmCerrar.moverForm", Erl)

    Resume Next
    
End Sub

Public Function Is_Transparent(ByVal hwnd As Long) As Boolean
    
    On Error GoTo Is_Transparent_Err

    Dim msg As Long

    msg = GetWindowLong(hwnd, GWL_EXSTYLE)

    If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
        Is_Transparent = True
    Else
        Is_Transparent = False

    End If

    If err Then
        Is_Transparent = False

    End If
    
    Exit Function

Is_Transparent_Err:
    Call LogError(err.Number, err.Description, "frmCerrar.Is_Transparent", Erl)

    Resume Next
    
End Function
  
'Funci�n que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hwnd As Long, Valor As Integer) As Long
    
    On Error GoTo Aplicar_Transparencia_Err

    Dim msg As Long

    If Valor < 0 Or Valor > 255 Then
        Aplicar_Transparencia = 1
    Else
        msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        msg = msg Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, msg
        'Establece la transparencia
        SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA
        Aplicar_Transparencia = 0

    End If
  
    If err Then
        Aplicar_Transparencia = 2

    End If
    
    Exit Function

Aplicar_Transparencia_Err:
    Call LogError(err.Number, err.Description, "frmCerrar.Aplicar_Transparencia", Erl)

    Resume Next
    
End Function

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    'Call FormParser.Parse_Form(Me)
    Call Aplicar_Transparencia(Me.hwnd, 220)
    Me.Picture = LoadPictureEX("desconectar.bmp")
    
    Exit Sub

Form_Load_Err:
    Call LogError(err.Number, err.Description, "frmCerrar.Form_Load", Erl)

    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Opcion(0).Tag = "0"
    Opcion(0).Picture = Nothing
    Opcion(1).Tag = "0"
    Opcion(1).Picture = Nothing
    Opcion(2).Tag = "0"
    Opcion(2).Picture = Nothing
    
    Exit Sub

Form_MouseMove_Err:
    Call LogError(err.Number, err.Description, "frmCerrar.Form_MouseMove", Erl)

    Resume Next
    
End Sub

Private Sub Opcion_Click(Index As Integer)

    On Error GoTo Opcion_Click_Err

    'Ladder 30/10/2020
    Select Case Index

    Case 0    ' Menu principal
        If Dir(App.path & "\INIT\" & UserName & ".rtf", vbArchive) <> "" Then

            Kill (App.path & "\INIT\" & UserName & ".rtf")

        End If
        Call WriteQuit

        Unload Me

    Case 1  'Cerrar juego
        If Dir(App.path & "\INIT\" & UserName & ".rtf", vbArchive) <> "" Then

            Kill (App.path & "\INIT\" & UserName & ".rtf")

        End If
        Call CloseClient


    Case 2    'Cancelar
        Unload Me

    End Select

    Exit Sub

Opcion_Click_Err:
    Call LogError(err.Number, err.Description, "frmCerrar.Opcion_Click", Erl)

    Resume Next

End Sub

Private Sub Opcion_MouseDown(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    
    On Error GoTo Opcion_MouseDown_Err

    'Ladder 30/10/2020
    Select Case Index

        Case 0 ' Menu principal
            Opcion(Index).Picture = LoadPictureEX("boton-menu-principal-ES-off.bmp")
            Opcion(Index).Tag = "1"

        Case 1  'Cerrar juego
            Opcion(Index).Picture = LoadPictureEX("boton-salir-ES-off.bmp")
            Opcion(Index).Tag = "1"

        Case 2 'Cancelar
            Opcion(Index).Picture = LoadPictureEX("boton-cancelar-ES-off.bmp")
            Opcion(Index).Tag = "1"

    End Select
    
    Exit Sub

Opcion_MouseDown_Err:
    Call LogError(err.Number, err.Description, "frmCerrar.Opcion_MouseDown", Erl)

    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err

    If (KeyAscii = 27) Then
        Unload Me

    End If
    
    Exit Sub

Form_KeyPress_Err:
    Call LogError(err.Number, err.Description, "frmCerrar.Form_KeyPress", Erl)

    Resume Next
    
End Sub

Private Sub Opcion_MouseMove(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    
    On Error GoTo Opcion_MouseMove_Err

    'Ladder 30/10/2020
    Select Case Index

        Case 0 ' Menu principal

            If Opcion(Index).Tag = "0" Then
                Opcion(Index).Picture = LoadPictureEX("boton-menu-principal-ES-over.bmp")
                Opcion(Index).Tag = "1"

            End If

        Case 1  'Cerrar juego

            If Opcion(Index).Tag = "0" Then
                Opcion(Index).Picture = LoadPictureEX("boton-salir-ES-over.bmp")
                Opcion(Index).Tag = "1"

            End If

        Case 2 'Cancelar

            If Opcion(Index).Tag = "0" Then
                Opcion(Index).Picture = LoadPictureEX("boton-cancelar-ES-over.bmp")
                Opcion(Index).Tag = "1"

            End If

    End Select
    
    Exit Sub

Opcion_MouseMove_Err:
    Call LogError(err.Number, err.Description, "frmCerrar.Opcion_MouseMove", Erl)

    Resume Next
    
End Sub
