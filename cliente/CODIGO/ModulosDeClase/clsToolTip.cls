VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private X       As Integer

Private Y       As Integer

Private OffX    As Single

Private OffY    As Single

Private Mensaje As String

Public Alpha    As Single

Private R       As Byte, G As Byte, b As Byte

Public Sub Render(TileX As Integer, _
                  TileY As Integer, _
                  ScreenOffX As Single, _
                  ScreenOffY As Single)

 On Error GoTo err:
    OffY = OffY - timerElapsedTime * 0.05
    Alpha = Alpha - timerElapsedTime * 0.2 '0.3

    If Alpha <= 0 Then
        Alpha = 0

    End If

    If X > 0 And Y > 0 Then
        Call RenderTextCentered((X - TileX + 12) * 32 + ScreenOffX + OffX - 16, (Y - TileY + 8) * 32 + ScreenOffY + OffY + 16, Mensaje, D3DColorRGBA(R, G, b, Alpha))

    End If
err:
End Sub

Public Sub Init(ByVal mX As Integer, _
                ByVal mY As Integer, _
                ByVal mMensaje As String, _
                ByVal mR As Byte, _
                ByVal mG As Byte, _
                ByVal mB As Byte)
    X = mX
    Y = mY
    Mensaje = mMensaje
    R = mR
    G = mG
    b = mB
    Alpha = 255

End Sub
