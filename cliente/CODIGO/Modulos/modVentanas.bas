Attribute VB_Name = "modVentanas"
Option Explicit

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal color As Long, _
                              ByVal bAlpha As Byte, _
                              ByVal Alpha As Long) As Boolean

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function CreateStreamOnHGlobal _
                Lib "ole32" (ByVal hGlobal As Long, _
                             ByVal fDeleteOnRelease As Long, _
                             ppstm As Any) As Long

Private Declare Function OleLoadPicture _
                Lib "olepro32" (pStream As Any, _
                                ByVal lSize As Long, _
                                ByVal fRunmode As Long, _
                                riid As Any, _
                                ppvObj As Any) As Long

Private Declare Function CLSIDFromString _
                Lib "ole32" (ByVal lpsz As Any, _
                             pclsid As Any) As Long

Private Declare Function GlobalAlloc _
                Lib "kernel32" (ByVal uFlags As Long, _
                                ByVal dwBytes As Long) As Long

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Sub MoveMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (pDest As Any, _
                                       pSource As Any, _
                                       ByVal dwLength As Long)

Private Const GWL_EXSTYLE = (-20)

Private Const WS_EX_LAYERED As Long = &H80000

Private Const LWA_ALPHA     As Long = &H2

Public Const NTRANS_GENERAL As Integer = 200

Public Sub SetTranslucent(ThehWnd As Long, nTrans As Integer)

    On Error GoTo ErrorRtn

    Dim attrib As Long

    'put current GWL_EXSTYLE in attrib
    attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)

    'change GWL_EXSTYLE to WS_EX_LAYERED - makes a window layered
    SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED

    'Make transparent (RGB value does not have any effect at this
    'time, will in Part 2 of this article)
    SetLayeredWindowAttributes ThehWnd, RGB(0, 0, 0), nTrans, LWA_ALPHA
    Exit Sub

ErrorRtn:
    MsgBox err.Description & " Source : " & err.Source

End Sub

Public Sub MoverVentana(hwnd As Long)
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, 0&

End Sub

Public Sub MessageBox(ByVal Message As String, _
                      Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                      Optional ByVal Title As String = "")
    frmMensaje.msg.Caption = Message
    frmMensaje.Show 1

End Sub

Public Function LoadPictureEX(ByVal FileName As String) As IPicture

    If FileName = "" Then
        Set LoadPictureEX = Nothing
    Else

        Dim b() As Byte

        'Call Get_File_Data(DirRecursos & "Interface.ao" & "Interface", FileName, b) 'neo carga de interface.ao
        Call Get_File_Data("Interface", FileName, b)
        Set LoadPictureEX = PictureFromByteStream(b)

    End If

End Function

Public Function PictureFromByteStream(ByRef b() As Byte) As IPicture

    Dim LowerBound As Long

    Dim ByteCount  As Long

    Dim hMem       As Long

    Dim lpMem      As Long

    Dim IID_IPicture(15)

    Dim istm As stdole.IUnknown

    On Error GoTo Err_Init

    If UBound(b, 1) < 0 Then
        Exit Function

    End If
    
    LowerBound = LBound(b)
    ByteCount = (UBound(b) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)

    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)

        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, b(LowerBound), ByteCount
            Call GlobalUnlock(hMem)

            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                    Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), PictureFromByteStream)

                End If

            End If

        End If

    End If
    
    Exit Function
    
Err_Init:

    MsgBox err.Number & " - " & err.Description

End Function
