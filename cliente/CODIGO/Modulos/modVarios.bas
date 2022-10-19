Attribute VB_Name = "modVarios"
Option Explicit

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Private Declare Function SendMessage _
                Lib "user32.dll" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function SetLayeredWindowAttributes _
                Lib "user32.dll" (ByVal hwnd As Long, _
                                  ByVal crKey As Long, _
                                  ByVal bAlpha As Byte, _
                                  ByVal dwFlags As Long) As Long

Const LW_KEY = &H1

Const G_E = (-20)

Const W_E = &H80000

Private Const WM_NCLBUTTONDOWN = &HA1

Private Const HTCAPTION = 2

Public Sub Skin(Frm As Form, Color As Long)
    Frm.BackColor = Color

    Dim Ret As Long

    Ret = GetWindowLong(Frm.hwnd, G_E)
    Ret = Ret Or W_E
    SetWindowLong Frm.hwnd, G_E, Ret
    SetLayeredWindowAttributes Frm.hwnd, Color, 0, LW_KEY

End Sub

Public Sub Auto_Drag(ByVal hwnd As Long)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)

End Sub

