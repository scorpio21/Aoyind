VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMensajes 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   4545
   ClientTop       =   7440
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox mensajes 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7011
      _Version        =   393217
      BackColor       =   -2147483641
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"FrmMensajes.frx":0000
   End
End
Attribute VB_Name = "FrmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_GETLINECOUNT = &HBA
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000
Private Sub Form_Load()
    On Error GoTo err
    'Me.Picture = LoadPictureEX("mensages" & RandomNumber(1, 3) & ".jpg")

    Skin Me, vbMagenta
    Call SetWindowLong(mensajes.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    mensajes.LoadFile (App.path & "/INIT/" & UserName & ".rtf")
    mensajes.Locked = True

    mensajes.SelLength = Len(mensajes)
    'mensajes.SelColor = RGB(255, 255, 255)
    'mensajes.SelBold = True

    mensajes.SelLength = 0
err:
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MoverVentana (Me.hwnd)
End Sub

Private Sub mensajes_GotFocus()
mensajes.SelStart = Len(mensajes.Text)
End Sub

Private Sub mensajes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If GetCurrentLine(mensajes) = 1 Then KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
        If GetCurrentLine(mensajes) = GetLineCount(mensajes) Then KeyCode = 0
    ElseIf KeyCode = 39 Then
        If GetCurrentLine(mensajes) = GetLineCount(mensajes) Then KeyCode = 0
    ElseIf KeyCode = 37 Then
        If GetCurrentLine(mensajes) = GetLineCount(mensajes) Then KeyCode = 0
    End If
End Sub

Private Sub mensajes_KeyUp(KeyCode As Integer, Shift As Integer)
mensajes.SelStart = Len(mensajes.Text)
End Sub

Private Sub mensajes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mensajes.SelStart = Len(mensajes.Text)
End Sub


Sub Skin(Frm As Form, Color As Long)

Frm.BackColor = Color
Dim ret As Long
ret = GetWindowLong(Frm.hwnd, G_E)
ret = ret Or W_E
SetWindowLong Frm.hwnd, G_E, ret
SetLayeredWindowAttributes Frm.hwnd, Color, 0, LW_KEY
End Sub

Private Sub Timer1_Timer()
'mensajes.LoadFile App.path & "/INIT/" & UserName & ".txt", rtfText
'mensajes.SelStart = Len(mensajes.Text)
'mensajes.Refresh
End Sub
Private Function GetCurrentLine(Txt As RichTextBox) As Long
    GetCurrentLine = SendMessage(Txt.hwnd, EM_LINEFROMCHAR, Txt.SelStart, 0&) + 1
End Function

Private Function GetLineCount(Txt As RichTextBox) As Long
    GetLineCount = SendMessage(Txt.hwnd, EM_GETLINECOUNT, 0&, 0&)
End Function
