Attribute VB_Name = "modTutorial"
Option Explicit

Private Type tTutorial

    Evento As String
    Linea1 As String
    Linea2 As String
    Linea3 As String
    Funcion As Integer

End Type

Public CantTutoriales As Integer

Public Tutoriales()   As tTutorial

Public Sub CargarTutorial()

    On Error Resume Next

    Dim archivoC As String
    
    archivoC = PathInit & "\tutorial.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
        'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar el tutorial. Falta el archivo tutorial.dat, reinstale el juego.", vbCritical + vbOKOnly)
        Exit Sub

    End If
    
    Dim I As Integer
    
    CantTutoriales = GetVar(archivoC, "Config", "Cantidad")
    
    ReDim Tutoriales(1 To CantTutoriales)
    
    For I = 1 To CantTutoriales
        Tutoriales(I).Evento = CByte(GetVar(archivoC, "Tutorial" & CStr(I), "Evento"))
        Tutoriales(I).Linea1 = GetVar(archivoC, "Tutorial" & CStr(I), "Linea1")
        Tutoriales(I).Linea2 = GetVar(archivoC, "Tutorial" & CStr(I), "Linea2")
        Tutoriales(I).Linea3 = GetVar(archivoC, "Tutorial" & CStr(I), "Linea3")
        Tutoriales(I).Funcion = Val(GetVar(archivoC, "Tutorial" & CStr(I), "Funcion"))
    Next I

End Sub

Public Sub DrawTextPergamino(ByVal Texto As String, _
                             ByVal Head As Integer, _
                             ByVal Opciones As Byte)

    'PergaminoDireccion = IIf(PergaminoDireccion = 1, 2, 1)
    'If PergaminoDireccion = 1 Then
    '    Audio.PlayWave ("213")
    'Else
    '    Audio.PlayWave ("214")
    'End If

    If PergaminoDireccion <> 1 Then
        PergaminoDireccion = 1
        Audio.PlayWave ("213")

    End If

    TiempoAbierto = Now

    D3DDevice.SetRenderTarget pRenderSurface, Nothing, ByVal 0
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    Call D3DX.DrawText(FontCartel, D3DColorRGBA(40, 35, 20, 255), Texto, DDRect(IIf(Head = 0, 5, 70), 0, 480, 80), DT_LEFT)

    Dim Color As Long

    Dim Grh   As GrhData
    
    If Head > 0 Then
        Grh = GrhData(HeadData(Head).Head(south).GrhIndex)
        Color = D3DColorRGBA(255, 255, 255, 210)
    
        Call Engine_Render_Rectangle(256 + 5, 256 + 5, 56, 56, 56, 0, 56, 56, , , , 14687, Color, Color, Color, Color)
        Call Engine_Render_Rectangle(256 + 18, 256 + 19, 32, 32, Grh.sX, Grh.sY, Grh.PixelWidth, Grh.PixelHeight, , , , Grh.FileNum, Color, Color, Color, Color)

    End If
    
    D3DDevice.EndScene
    D3DDevice.SetRenderTarget pBackbuffer, Nothing, ByVal 0

End Sub

