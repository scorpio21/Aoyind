Attribute VB_Name = "modPresentacion"
Option Explicit

'Presentacion
Public GTCPres            As Long, GTCInicial As Long, GTCChars As Long

Public Conectar           As Boolean

Public MostrarEntrar      As Long

Public MostrarCrearCuenta As Boolean

Private MouseOn           As Integer

Public logUser            As String

Public logPass            As String

Public logEmail           As String

Dim CallDoConectar        As Boolean

Dim IntroChars()          As Char

Dim CantPjs               As Integer

Dim MousePj               As Integer

Dim AperturaX             As Integer

Dim AperturaY             As Integer

Dim AperturaPj            As Integer

Dim AperturaTick          As Long

Dim fXGrh                 As Grh

'Dim fXGrhTrueno1 As Grh
Dim WithfX                As Byte

Dim AperturaPjLast        As Integer

Dim rPosX                 As Integer

Dim rPosY                 As Integer

Public UseMotionBlur      As Byte    'If motion blur is enabled or not

Public BlurIntensity      As Single

Public BlurTexture        As Direct3DTexture8

Public BlurSurf           As Direct3DSurface8

Public BlurStencil        As Direct3DSurface8

Public DeviceStencil      As Direct3DSurface8

Public DeviceBuffer       As Direct3DSurface8

Public BlurTA(0 To 3)     As TLVERTEX
 
'Zoom level - 0 = No Zoom, > 0 = Zoomed
Public ZoomLevel          As Single

Public Const MaxZoomLevel As Single = 0.183
 
Public Const ScreenWidth  As Long = 1024 'anchonho

Public Const ScreenHeight As Long = 768 'altongo d su render


Public Sub CerrarCrearCuenta()

    If Conectar Then
        If GTCPres < 10000 Then
            GTCInicial = GTCInicial - (10000 - GTCPres)
            Call Audio.PlayMp3("10.mp3")
        ElseIf MostrarEntrar > 0 Then

            If GTCPres - MostrarEntrar > 1000 Then
                MostrarEntrar = -GTCPres
                frmMain.tUser.Visible = False
                frmMain.tPass.Visible = False
                frmMain.tEmail.Visible = False
                frmMain.tRePass.Visible = False
                frmMain.TCod.Visible = False
                Call Audio.PlayWave(SND_CADENAS)

            End If
            
        Else
            prgRun = False
            Audio.StopMidi

        End If

    End If

End Sub

Public Sub MouseAction(X As Single, Y As Single, Button As Integer)

    If GTCPres > 7000 Then    'Me aseguro que este todo cargado

'Debug.Print X & "-" & Y
        If MostrarCrearCuenta = True Then
            If X >= 334 And X <= 703 + 112 And Y >= 514 And Y <= 730 Then   ' CUADRO CREAR CUENTA

                'BOTON SALIR
                If X >= 426 And X <= 498 And Y >= 704 And Y <= 734 Then
                    If Button = 1 Then
                        Call Audio.PlayWave(SND_CLICKNEW)
                        Call CerrarCrearCuenta
                    Else

                        If MouseOn <> 2 Then
                            MouseOn = 2
                            Call Audio.PlayWave(SND_MOUSEOVER)

                        End If

                    End If

                    'BOTON CREAR CUENTA
                ElseIf X >= 530 And X <= 605 And Y >= 704 And Y <= 734 Then

                    If Button = 1 Then
                        If ClickCrearCuenta Then Call CerrarCrearCuenta
                        Call Audio.PlayWave(SND_CLICKNEW)
                    Else

                        If MouseOn <> 2 Then
                            MouseOn = 2
                            Call Audio.PlayWave(SND_MOUSEOVER)

                        End If

                    End If


                    'BOTON ENVIAR MAIL
                ElseIf X >= 602 And X <= 670 And Y >= 678 And Y <= 703 Then

                    If Button = 1 Then
                        

                        ' If ClickCrearCuenta Then Call CerrarCrearCuenta
                        Call Audio.PlayWave(SND_CLICKNEW)
                        frmMain2.EnviarCorreoVal

                    Else

                        If MouseOn <> 2 Then
                            MouseOn = 2
                            Call Audio.PlayWave(SND_MOUSEOVER)

                        End If

                    End If

                Else
                    MouseOn = 0

                End If

            Else    'Si el mouse esta fuera del cuadro entrar

                If Button = 1 Then
                    Call Audio.PlayWave(SND_CLICKOFF)

                End If

                MouseOn = 0

            End If

        ElseIf MostrarEntrar > 0 Then    'Si esta abierto el cuadro entrar

            If X >= 145 + 112 And X <= 660 + 112 And Y >= 345 + 168 And Y <= 555 + 168 Then    ' Cuadro entrar
                If X >= 500 + 112 And X <= 545 + 112 And Y >= 432 + 168 And Y <= 484 + 168 Then
                    If Button = 1 Then
                        ClickAbrirCuenta
                        Call Audio.PlayWave(SND_CLICKNEW)
                    Else

                        If MouseOn <> 2 Then
                            MouseOn = 2
                            Call Audio.PlayWave(SND_MOUSEOVER)

                        End If

                    End If

                    'RECORDAR BUTTON
                ElseIf X >= 616 And X <= 646 And Y >= 676 And Y <= 706 Then

                    If Button = 1 Then
                        mOpciones.Recordar = Not mOpciones.Recordar
                        Call Audio.PlayWave(SND_CLICKNEW)
                    Else

                        If MouseOn <> 2 Then
                            MouseOn = 2
                            Call Audio.PlayWave(SND_MOUSEOVER)

                        End If

                    End If

                Else
                    MouseOn = 0

                End If

            Else    'Si el mouse esta fuera del cuadro entrar

                If Button = 1 Then
                    Call Audio.PlayWave(SND_CLICKOFF)

                End If

                MouseOn = 0

            End If

        ElseIf MostrarEntrar = 0 Then  'Si no esta abierto el cuadro entrar

            If AperturaPj > 0 And Button = 1 Then
                CloseSock
                AperturaPj = -AperturaPj
                AperturaTick = (GetTickCount() And &H7FFFFFFF)
            ElseIf X >= 355 + 112 And X <= 450 + 112 And Y >= 130 And Y <= 160 Then    'Boton Entrar

                If Button = 1 Then
                    MostrarEntrar = GTCPres
                    Call Audio.PlayWave(SND_CLICKNEW)
                    Call Audio.PlayWave(SND_CADENAS)
                Else

                    If MouseOn <> 1 Then
                        MouseOn = 1
                        Call Audio.PlayWave(SND_MOUSEOVER)

                    End If

                End If

            ElseIf X >= 15 + 112 And X <= 105 + 112 And Y >= 50 And Y <= 75 Then    'Boton crear

                If Button = 1 Then
                    MostrarEntrar = GTCPres
                    MostrarCrearCuenta = True
                    Call Audio.PlayWave(SND_CLICKNEW)
                    Call Audio.PlayWave(SND_CADENAS)
                Else

                    If MouseOn <> 1 Then
                        MouseOn = 1
                        Call Audio.PlayWave(SND_MOUSEOVER)

                    End If

                End If

            ElseIf X >= 121 + 112 And X <= 229 + 112 And Y >= 50 And Y <= 75 Then    'Boton recuperar

                If Button = 1 Then
                    frmNavegador.TIPO = Recuperar
                    frmNavegador.Show vbModal
                    Call Audio.PlayWave(SND_CLICKNEW)
                Else

                    If MouseOn <> 1 Then
                        MouseOn = 1
                        Call Audio.PlayWave(SND_MOUSEOVER)

                    End If

                End If

            ElseIf X >= 576 + 112 And X <= 668 + 112 And Y >= 50 And Y <= 75 Then    'Boton borrar

                If Button = 1 Then
                    frmBorrar.Show vbModeless
                    Call Audio.PlayWave(SND_CLICKNEW)
                Else

                    If MouseOn <> 1 Then
                        MouseOn = 1
                        Call Audio.PlayWave(SND_MOUSEOVER)

                    End If

                End If

            ElseIf X >= 693 + 112 And X <= 783 + 112 And Y >= 50 And Y <= 75 Then    'Boton salir

                If Button = 1 Then
                    prgRun = False
                    Call Audio.PlayWave(SND_CLICKNEW)
                Else

                    If MouseOn <> 1 Then
                        MouseOn = 1
                        Call Audio.PlayWave(SND_MOUSEOVER)

                    End If

                End If

            ElseIf X >= 105 + 112 And X <= 200 + 112 And Y >= 130 And Y <= 160 Then

                If Button = 1 Then

                Else

                    If MouseOn <> 1 Then
                        MouseOn = 1
                        Call Audio.PlayWave(SND_MOUSEOVER)

                    End If

                End If

            ElseIf CantPjs > 0 Then

                Dim I As Integer

                Dim Angulo As Single

                MousePj = 0

                For I = 1 To CantPjs
                    Angulo = (-40 * CantPjs + I * 80 - 48) / 180 - 1.57

                    If Abs(X - (512 + Cos(Angulo) * 320 + 16)) < 32 And Abs(Y - (450 + Sin(Angulo) * 160)) < 54 Then
                        MousePj = I

                    End If

                Next I

                'If x >= 400 - CantPjs * 40 And x <= 400 + CantPjs * 40 And y >= 250 And y <= 350 Then
                '    MousePj = (x - 400 + CantPjs * 40 - 48) / 80 + 1
                'Else
                '    MousePj = 0
                'End If
                If MousePj > 0 And Button = 1 Then
                    If IntroChars(MousePj).priv = 9 Then
                        frmCrearPersonaje.Show , frmMain
                        Call Audio.PlayWave(SND_CLICKNEW)
                    Else
                        UserName = IntroChars(MousePj).nombre
                        AperturaPj = MousePj
                        AperturaTick = (GetTickCount() And &H7FFFFFFF)
                        EstadoLogin = Normal
                        'Login

                        iServer = 0
                        iCliente = 0
                        DummyCode = StrConv(StrReverse("conectar") & "CuEnTa", vbFromUnicode)
                        DoEvents

                        If Not ClientSetup.WinSock Then
                            frmMain.Client.CloseSck
                            frmMain.Client.Connect IpServidor, PuertoServidor
                        Else
                            frmMain.WSock.Close
                            frmMain.WSock.Connect IpServidor, PuertoServidor

                        End If

                    End If

                End If

            Else
                MouseOn = 0

            End If

        End If

    End If

End Sub

Public Sub ClickAbrirCuenta()
    logUser = frmMain.tUser.Text
    logPass = frmMain.tPass.Text

    If mOpciones.Recordar = True Then
        mOpciones.RecordarUsuario = logUser
        mOpciones.RecordarPassword = logPass
    Else
        mOpciones.RecordarUsuario = ""
        mOpciones.RecordarPassword = ""

    End If

    Call SaveConfig

    If frmMain.tEmail.Visible = True Then
        logEmail = frmMain.tEmail.Text

    End If

    DoConectar

End Sub

Public Function ClickCrearCuenta() As Boolean
    
    ClickCrearCuenta = False
    
    If frmMain.tUser.Text = "" Then
        MessageBox "Escriba un usuario."
        Exit Function

    End If
    
    If frmMain.tPass.Text = "" Then
        MessageBox "Escriba un password."
        Exit Function

    End If
    
    If frmMain.tEmail.Text = "" Then
        MessageBox "Escriba un e-mail."
        Exit Function

    End If
    
    If frmMain.tPass.Text <> frmMain.tRePass.Text Then
        MessageBox "Las contraseñas no coinciden."
        Exit Function

    End If
    
     If frmMain.TCod.Text = "" Then
        MessageBox "Ingrese el codigo de Validacion.. En caso de no tenerlo presione en enviar."
        Exit Function

    End If
    
    If frmMain.TCod.Text <> CodVerificacion Then
     MessageBox "El codigo de validacion ingresado es invalido verifique sus e-mail"
        Exit Function

    End If
    
    UserAccount = frmMain.tUser.Text
    UserPassword = MD5(frmMain.tPass.Text)
    UserEmail = frmMain.tEmail.Text
    
    If right$(frmMain.tUser.Text, 1) = " " Then
        UserAccount = RTrim$(UserAccount)
        MessageBox "Nombre invalido, se han removido los espacios al final del nombre"

    End If

    If Len(frmMain.tUser.Text) > 20 Then
        MessageBox "El nombre es demasiado largo, debe tener como máximo 20 letras."
        Exit Function

    End If
    
    If Not ClientSetup.WinSock Then
        frmMain.Client.CloseSck
                
        EstadoLogin = E_MODO.CrearCuenta
                
        frmMain.Client.Connect IpServidor, PuertoServidor
                
        If Not frmMain.Client.State <> SockState.sckConnected Then
            MessageBox "Error de conexión."
        Else
            Call Login

        End If

    Else
        frmMain.WSock.Close
                
        EstadoLogin = E_MODO.CrearCuenta
                
        frmMain.WSock.Connect IpServidor, PuertoServidor
                
        If Not frmMain.WSock.State <> SockState.sckConnected Then
            MessageBox "Error de conexión."
        Else
            Call Login

        End If

    End If
    CodVerificacion = "hhg97YyEssGk56aahH"
    ClickCrearCuenta = True

End Function

Public Sub DoConectar()
    CloseSock
    DoEvents
    'update user info
    UserName = logUser

    Dim aux As String

    aux = logPass
    UserPassword = MD5(aux)
    UserAccount = logUser
    iCliente = 0
    iServer = 0
    DummyCode = StrConv(StrReverse("conectar") & "CuEnTa", vbFromUnicode)

    If CheckUserData(False) = True Then
        EstadoLogin = Cuentas
    
        If Not ClientSetup.WinSock Then
            frmMain.Client.Connect IpServidor, PuertoServidor
        Else
            frmMain.WSock.Connect IpServidor, PuertoServidor

        End If

    End If

End Sub

Sub RenderConectar()

    Static Ang As Single

    Dim Color As Long

    GTCPres = Abs((GetTickCount() And &H7FFFFFFF) - GTCInicial)

    If GTCPres < 4000 Then
        Color = D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 0, 220, 15))
        Call Engine_Render_Rectangle(412, 284, 200, 200, 0, 0, 400, 400, , , 0, 14705, Color, Color, Color, Color)
        Call Engine_Render_Rectangle(390, 340, 262, 402, 0, 0, 262, 402, , , 0, 14785, Color, Color, Color, Color)

    End If

    Dim a As Single

    Dim T As Single

    Dim T2 As Single

    Dim Mueve As Single

    Dim tmpColor As Long

    Dim I As Integer

    Dim Angulo As Single

    Dim Separacion As Single

    Dim X As Single

    Dim Y As Single

    Static AngSel As Single

    a = 20
    T = (GTCPres - 4000) / 1000

    If GTCPres >= 4000 Then
        If MostrarEntrar > 0 Then
            tmpColor = 255 - CalcAlpha(GTCPres, MostrarEntrar, 170, 5)
        ElseIf MostrarEntrar < 0 Then
            tmpColor = 85 + CalcAlpha(GTCPres, -MostrarEntrar, 170, 5)
        Else
            tmpColor = 255

        End If

        Call Engine_Render_D3DXSprite(0, 0, 1024, 768, 0, 0, D3DColorRGBA(tmpColor, tmpColor, tmpColor, CalcAlpha(GTCPres, 4000, 255, 15)), 14704, 0)

        tmpColor = D3DColorRGBA(255, 255, 255, 220 - CalcAlpha(GTCPres, 4000, 220, 15))
        Call Engine_Render_Rectangle(412, 284 - a * (T ^ 2), 200, 200, 0, 0, 400, 400, , , 0, 14705, tmpColor, tmpColor, tmpColor, tmpColor)

        Mueve = (T * 20) Mod 512

        If CantPjs > 0 Then

            For I = 1 To CantPjs
                T2 = Abs((GetTickCount() And &H7FFFFFFF) - GTCChars) / 1000
                Separacion = 120 * T2 - 40 * (T2 ^ 2)

                If T2 > 1 Then Separacion = 80
                If (AperturaPj <= 0 And MousePj = I) Or AperturaPj = I Then
                    If IntroChars(I).Alpha < 255 Then
                        If RandomNumber(1, 3) = 1 Then
                            IntroChars(I).Alpha = IntroChars(I).Alpha + 1

                        End If

                    End If

                    If IntroChars(I).Alpha >= 250 Then IntroChars(I).Alpha = 255
                Else
                    IntroChars(I).Alpha = 85

                End If

                'Call DrawFont(CStr(Separacion), 323, 423, D3DColorRGBA(255, 255, 255, 160))
                'Call CharRender(IntroChars(i), -1, 255 + 400 - Separacion / 2 * CantPjs + i * Separacion - 48 * Separacion / 80, 255 + 300)

                Angulo = (-Separacion / 2 * CantPjs + I * Separacion - 48 * Separacion / 80) / 180 - 1.57

                T2 = Abs((GetTickCount() And &H7FFFFFFF) - AperturaTick) / 1000

                If AperturaPj > 0 And AperturaX < 660 Then
                    AperturaX = 320 + (T2 ^ 2) * 550
                    AperturaY = 160 + (T2 ^ 2) * 412.5

                    If AperturaX >= 660 Then
                        AperturaX = 660
                        AperturaY = 441

                    End If

                ElseIf AperturaPj < 0 And AperturaX > 320 Then
                    AperturaX = 660 - (T2 ^ 2) * 550
                    AperturaY = 441 - (T2 ^ 2) * 412.5

                    If AperturaX <= 320 Then
                        AperturaX = 320
                        AperturaY = 160

                    End If

                End If

                If I = AperturaPj Or I = -AperturaPj Then
                    X = 497 + Cos(Angulo) * (497 - AperturaX / 2)
                    Y = 412 + Sin(Angulo) * (147 - AperturaY / 3) - (AperturaY - 110) / 6
                Else
                    X = 512 + Cos(Angulo) * AperturaX * 1.5
                    Y = 450 + Sin(Angulo) * AperturaY * 1.5

                End If

                If IntroChars(I).logged Then
                    IntroChars(I).Alpha = 70

                End If

                Call CharRender(IntroChars(I), -1, X, Y)
                a = a + 1

                If (AperturaPj <= 0 And MousePj = I) Or AperturaPj = I Then
                    rPosX = X
                    rPosY = Y

                End If

            Next I

        End If

        'NIEBLA
        tmpColor = D3DColorRGBA(255, 200, 200, CalcAlpha(GTCPres, 4000, 165, 15))

        Call Engine_Render_D3DXSprite(0, 0, 512 - Mueve, 512, Mueve, 0, tmpColor, 14706, 0)
        Call Engine_Render_D3DXSprite(0, 512, 512 - Mueve, 256, Mueve, 0, tmpColor, 14706, 0)

        Call Engine_Render_D3DXSprite(512 - Mueve, 0, 512, 512, 0, 0, tmpColor, 14706, 0)
        Call Engine_Render_D3DXSprite(512 - Mueve, 512, 512, 256, 0, 0, tmpColor, 14706, 0)

        Call Engine_Render_D3DXSprite(1024 - Mueve, 0, Mueve, 512, 0, 0, tmpColor, 14706, 0)
        Call Engine_Render_D3DXSprite(1024 - Mueve, 512, Mueve, 256, 0, 0, tmpColor, 14706, 0)

        If MousePj > 0 Then
            If UBound(IntroChars) > 0 Then
                Call CharRender(IntroChars(MousePj), -1, rPosX, rPosY)

                D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

                'AngSel = AngSel + 0.05
                'Call Engine_Render_Rectangle(X - 47, Y - 52, 128, 128, 224, 0, 128, 128, , , AngSel, 14332)
                If WithfX = 0 Then
                    Call InitGrh(fXGrh, FxData(50).Animacion)
                    fXGrh.Speed = 900
                    fXGrh.Loops = 1
                    WithfX = 1

                End If

                Call DrawGrh(fXGrh, rPosX, rPosY + 40, 1, 1)

                'Call DrawGrh(fXGrhTrueno1, X + 45, Y + 30, 1, 1)
                'If fXGrh.Started = 0 Then .fX = 0

                If AperturaPjLast = MousePj Then
                    WithfX = 1
                Else
                    WithfX = 0
                    AperturaPjLast = MousePj

                End If

                If Conectar Then
                    Color = D3DColorRGBA(122, 122, 122, 122)

                    If IntroChars(MousePj).Clase > 0 Then
                        Call RenderTextCentered(rPosX + 20, rPosY + 65, ListaClases(IntroChars(MousePj).Clase) & " (" & IntroChars(MousePj).Elv & ")", Color)

                        Color = D3DColorRGBA(70, 70, 0, 100)    'Color para el GLD

                        Dim strOro As String

                        If IntroChars(MousePj).Gld <> 0 Then
                            strOro = Format$(IntroChars(MousePj).Gld, "##,##")
                        Else
                            strOro = IntroChars(MousePj).Gld

                        End If

                        Call RenderTextCentered(rPosX + 20, rPosY + 80, "$" & strOro, Color)

                        If IntroChars(MousePj).logged Then
                            Color = D3DColorRGBA(10, 200, 10, 200)
                            Call RenderTextCentered(rPosX + 20, rPosY + 105, "(Online)", Color)

                        End If

                    End If

                End If

                D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

            End If

        End If

        Mueve = (T * 20) Mod 512

        tmpColor = D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 150, 15))

        If MostrarEntrar > 0 Then
            T2 = (GTCPres - MostrarEntrar) / 1000

            If MostrarCrearCuenta = True Then
                If T2 < 1 Then
                    Call Engine_Render_D3DXSprite(314, 768 - 388.5 * T2 + 263 / 2 * (T2 ^ 2), 400, 263, 0, 0, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, MostrarEntrar, 200, 4)), 14769, 0)
                Else
                    Call Engine_Render_D3DXSprite(314, 768 - 272, 400, 263, 0, 0, D3DColorRGBA(255, 255, 255, 200), 14769, 0)

                    If frmMain.tUser.Visible = False Then
                        frmMain.tUser.Text = ""
                        frmMain.tEmail.Text = ""
                        frmMain.tPass.Text = ""
                        frmMain.tRePass.Text = ""
                        frmMain.TCod.Text = ""

                        frmMain.tUser.Visible = True
                        frmMain.tPass.Visible = True
                        frmMain.tEmail.Visible = True
                        frmMain.tRePass.Visible = True
                        frmMain.TCod.Visible = True
                        frmMain.tUser.SetFocus

                        frmMain.tUser.top = 603 - 13 - 8
                        frmMain.tEmail.top = 626 - 10 - 8
                        frmMain.tPass.top = 648 - 5 - 8
                        frmMain.tRePass.top = 665 + 2 - 8
                        frmMain.TCod.top = 687 + 2 - 8

                        frmMain.tPass.left = 446
                        frmMain.tRePass.left = 446
                        frmMain.tUser.left = 446
                        frmMain.tEmail.left = 446
                        frmMain.TCod.left = 446

                        frmMain.tUser.Width = 153
                        frmMain.tEmail.Width = 153
                        frmMain.tPass.Width = 153
                        frmMain.tRePass.Width = 153
                        frmMain.TCod.Width = 153

                        frmMain.tUser.Height = 15
                        frmMain.tEmail.Height = 15
                        frmMain.tPass.Height = 15
                        frmMain.tRePass.Height = 15
                        frmMain.TCod.Height = 15
                    End If

                End If

            Else

                If T2 < 1 Then
                    Call Engine_Render_D3DXSprite(0, 768 + 14 - 388.5 * T2 + 259 / 2 * (T2 ^ 2), 1024, 259, 0, 177, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, MostrarEntrar, 200, 4)), 14703, 0)
                Else
                    Call Engine_Render_D3DXSprite(0, 768 + 14 - 259, 1024, 259, 0, 177, D3DColorRGBA(255, 255, 255, 200), 14703, 0)

                    If mOpciones.Recordar = True Then
                        Call Engine_Render_D3DXSprite(614, 664 + 14, 37, 41, 0, 0, D3DColorRGBA(255, 255, 255, 200), 14771, 0)
                    Else
                        Call Engine_Render_D3DXSprite(614, 664 + 14, 37, 41, 0, 0, D3DColorRGBA(255, 255, 255, 200), 14770, 0)

                    End If

                    If frmMain.tUser.Visible = False Then

                        frmMain.tUser.top = 612 + 5
                        frmMain.tPass.top = 637 + 5
                        frmMain.tPass.left = 428
                        frmMain.tUser.left = 428

                        frmMain.tUser.Width = 180
                        frmMain.tPass.Width = 180

                        frmMain.tUser.Height = 19
                        frmMain.tPass.Height = 19

                        frmMain.tEmail.Visible = False
                        frmMain.tUser.Visible = True
                        frmMain.tPass.Visible = True
                        frmMain.tRePass.Visible = False
                        frmMain.TCod.Visible = False
                        frmMain.tUser.SetFocus

                    End If

                End If

            End If

        ElseIf MostrarEntrar < 0 Then
            T2 = (GTCPres + MostrarEntrar) / 1000

            If MostrarCrearCuenta = True Then
                If T2 < 1 Then
                    Call Engine_Render_D3DXSprite(314, 768 - 263 + 388.5 * T2 - 263 / 2 * (T2 ^ 2), 400, 263, 0, 0, D3DColorRGBA(255, 255, 255, 200 - CalcAlpha(GTCPres, -MostrarEntrar, 200, 4)), 14769, 0)
                Else
                    MostrarCrearCuenta = False
                    MostrarEntrar = 0

                    If CallDoConectar Then
                        CallDoConectar = False
                        MostrarEntrar = -GTCPres
                        frmMain.tUser.Visible = False
                        frmMain.tPass.Visible = False
                        frmMain.tEmail.Visible = False
                        frmMain.tRePass.Visible = False
                        frmMain.TCod.Visible = False
                        Call Audio.PlayWave(SND_CADENAS)

                    End If

                End If

            Else

                If T2 < 1 Then
                    Call Engine_Render_D3DXSprite(0, 768 - 259 + 388.5 * T2 - 259 / 2 * (T2 ^ 2), 800, 259, 0, 177, D3DColorRGBA(255, 255, 255, 200 - CalcAlpha(GTCPres, -MostrarEntrar, 200, 4)), 14703, 0)
                Else
                    MostrarEntrar = 0

                    If CallDoConectar Then
                        CallDoConectar = False
                        MostrarEntrar = -GTCPres
                        frmMain.tUser.Visible = False
                        frmMain.tPass.Visible = False
                        frmMain.tEmail.Visible = False
                        frmMain.tRePass.Visible = False
                         frmMain.TCod.Visible = False
                        Call Audio.PlayWave(SND_CADENAS)

                    End If

                End If

            End If

        End If

        If T <= 4 Then
            Call Engine_Render_D3DXSprite(0, -177 + Int(88.5 * T - 22.125 / 2 * (T ^ 2)), 1024, 177, 0, 0, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 255, 15)), 14703, 0)
            Call Engine_Render_D3DXSprite(0, 768 - Int(23 * T - 5.75 / 2 * (T ^ 2)), 1024, 47, 0, 436, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 255, 15)), 14703, 0)
        Else
            Call Engine_Render_D3DXSprite(0, 0, 1024, 177, 0, 0, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 255, 15)), 14703, 0)
            Call Engine_Render_D3DXSprite(0, 768 - 34, 1024, 47, 0, 436, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 255, 15)), 14703, 0)

        End If

        'If T Mod 2 = 0 Then If Not MP3P.IsItPlaying Then Call EscucharMp3(10)
    End If

    Ang = Ang + 0.001

    'Call TestPart.ReLocate(Cos(Ang * 3) * 40 + 300, Sin(Ang * 2) * 40 + 300)

    'TestPart.Update
    'TestPart.Render

    'Call Engine_Render_D3DXTexture(255 + 209, 255 + 200, 103, 132, 0, 0, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 0, 220, 15)), ImgBruma, 0)

    'Particle_Group_Render MapData(150, 800, 1).particle_group_index, 400, 400

End Sub

Sub ShowConnect()
    Call Audio.PlayMp3("10.mp3")
    frmMain.SetRender (True)
    GTCInicial = (GetTickCount() And &H7FFFFFFF) - 10000
    GTCPres = (GetTickCount() And &H7FFFFFFF)
    MouseOn = 0
    MostrarEntrar = 0
    Conectar = True
    CantPjs = 0
    ReDim IntroChars(0)
    frmMain.tUser.Visible = False
    frmMain.tPass.Visible = False
    frmMain.tEmail.Visible = False
    frmMain.tRePass.Visible = False
     frmMain.TCod.Visible = False

End Sub

Public Sub HandleOpenAccount()

    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub

    End If

    'Remove packet ID
    Call incomingData.ReadByte

    Dim I As Integer

    Dim Arma As Integer, Escudo As Integer, Casco As Integer

    CantPjs = incomingData.ReadInteger() + 1

    If CantPjs > 0 Then
        ReDim IntroChars(1 To CantPjs)

        For I = 1 To CantPjs - 1

            With IntroChars(I)
                .ACTIVE = 1

                .nombre = incomingData.ReadASCIIString()
                .iHead = incomingData.ReadInteger()
                .iBody = incomingData.ReadInteger()

                Arma = incomingData.ReadInteger()
                Escudo = incomingData.ReadInteger()
                Casco = incomingData.ReadInteger()

                If Arma = 0 Then Arma = 2
                If Escudo = 0 Then Escudo = 2
                If Casco = 0 Then Casco = 2

                .Head = HeadData(.iHead)
                .Body = BodyData(.iBody)
                .Arma = WeaponAnimData(Arma)
                .Escudo = ShieldAnimData(Escudo)
                .Casco = CascoAnimData(Casco)

                If .iBody = FRAGATA_FANTASMAL Then
                    .Head = HeadData(2)

                End If

                .Heading = south

                .Alpha = 0

                .logged = incomingData.ReadBoolean()
                .Criminal = incomingData.ReadBoolean()
                .muerto = .iHead = CASPER_HEAD Or .iHead = CASPER_HEAD_CRIMI Or .iBody = FRAGATA_FANTASMAL
                .Elv = incomingData.ReadByte()
                .Clase = incomingData.ReadByte()
                .Gld = incomingData.ReadLong()
                .priv = incomingData.ReadByte()
                .priv = Log(.priv) / Log(2)

            End With

        Next I

        If CantPjs <= 10 Then    'Si tiene 10 pjs no le deja crear mas

            With IntroChars(CantPjs)
                .ACTIVE = 1
                .nombre = "CREAR PJ"
                .iHead = 10
                .iBody = 21
                .priv = 9
                .Elv = 1
                .Head = HeadData(.iHead)
                .Body = BodyData(.iBody)
                .Arma = WeaponAnimData(2)
                .Escudo = ShieldAnimData(2)
                .Casco = CascoAnimData(2)

                .Heading = south

                .Alpha = 0

            End With

        Else
            CantPjs = CantPjs - 1

        End If

    End If

    MousePj = 0
    MostrarEntrar = -GTCPres
    frmMain.tUser.Visible = False
    frmMain.tPass.Visible = False
    frmMain.tEmail.Visible = False
    frmMain.tRePass.Visible = False
    frmMain.TCod.Visible = False
    GTCChars = (GetTickCount() And &H7FFFFFFF)

    AperturaX = 220
    AperturaY = 110
    AperturaPj = 0
    AperturaTick = 0

End Sub

Public Function RandomLetrasMayusculas(ByVal cantidad As Integer) As String
    Dim I&, menor&, mayor&, X&, R$      'Declaraciones
    menor = 65                      'Caracter Ascii menor
    mayor = 90                      'Caracter Ascii mayor
    Randomize                           'Inicializar el generador de numeros aleatorios
    For I = 1 To cantidad             'Desde uno hasta la longitud del texto
        X = Int((mayor - menor + 1) * Rnd + menor)      'Escoje un valor aleatorio tipo integer entre el Ascii menor y el Ascii mayor y lo asigna a x
        If X > mayor Then X = mayor         'Si el valor de x es mayor al Ascii mayor entonces el valor de x es igual al Ascii mayor
        R = R & Chr$(X)                     'El texto final es el texto final mas el caracter que simboliza el codigo Ascii en x
    Next I                              'Termina el contador
    RandomLetrasMayusculas = R                      'La funcion es igual al texto generado
End Function



Public Function RandomLetrasMinusculas(ByVal cantidad As Integer) As String
    Dim I&, menor&, mayor&, X&, R$      'Declaraciones
    menor = 97                      'Caracter Ascii menor
    mayor = 122                        'Caracter Ascii mayor
    Randomize                           'Inicializar el generador de numeros aleatorios
    For I = 1 To cantidad             'Desde uno hasta la longitud del texto
        X = Int((mayor - menor + 1) * Rnd + menor)      'Escoje un valor aleatorio tipo integer entre el Ascii menor y el Ascii mayor y lo asigna a x
        If X > mayor Then X = mayor         'Si el valor de x es mayor al Ascii mayor entonces el valor de x es igual al Ascii mayor
        R = R & Chr$(X)                     'El texto final es el texto final mas el caracter que simboliza el codigo Ascii en x
    Next I                              'Termina el contador
    RandomLetrasMinusculas = R                      'La funcion es igual al texto generado
End Function
