Attribute VB_Name = "modEngine"
Option Explicit

Public dX                       As DirectX8

Public D3D                      As Direct3D8

Public D3DX                     As D3DX8

Public D3DDevice                As Direct3DDevice8

Public ParticleTexture(1 To 13) As Direct3DTexture8

Public ShadowColor              As Long

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE 'D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1 ' D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Public Const PI             As Single = 3.14159275180032   'can be worked out using (4*atn(1))

Public Const ANSI_FIXED_FONT As Long = 11

Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Type TLVERTEX

    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    tu As Single
    tv As Single

End Type

Public Type TLVERTEXDos

    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    tu As Single
    tv As Single
    Specular As Long
End Type

Public Type TexInfo

    X As Integer
    Y As Integer

End Type

Private Type PosC

    X As Long
    Y As Long
    x2 As Long
    y2 As Long

End Type

Public SurfaceSize()        As TexInfo

'The size of a FVF vertex
Private Const FVF_Size      As Long = 28

Public MainFont             As D3DXFont

Public MainFontDesc         As IFont

Public fnt                  As New StdFont

Public FontCartel           As D3DXFont

Public FontCartelDesc       As IFont

Public fntCartel            As New StdFont

Public MainFontBig          As D3DXFont

Public MainFontBigDesc      As IFont

Public fnt2                 As New StdFont

Public FontRender           As D3DXFont

Public FontRenderDesc       As IFont

Public fntRender            As New StdFont

Public pRenderTexture       As Direct3DTexture8

Public pRenderSurface       As Direct3DSurface8

Public pBackbuffer          As Direct3DSurface8

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180

Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

Private LastTexture         As Integer

Public Textura              As Direct3DTexture8

Private Caracteres(255)     As PosC

Public Sprite               As D3DXSprite

Private SpriteScaleVector   As D3DVECTOR2

Public RectJuego            As D3DRECT

Dim End_Time                As Currency

Dim timer_freq              As Currency

Public Function DDRect(X, Y, x1, y1) As RECT
    DDRect.bottom = y1
    DDRect.top = Y
    DDRect.left = X
    DDRect.right = x1

End Function

Public Function IniciarD3D() As Boolean

    On Error Resume Next

    Set dX = New DirectX8

    If err Then
        MessageBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function

    End If

    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8

    If Not IniciarDevice(D3DCREATE_PUREDEVICE) Then
        If Not IniciarDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not IniciarDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not IniciarDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    MessageBox "No se pudo iniciar el D3DDevice. Saliendo...", vbCritical
                    LiberarObjetosDX
                    End

                End If

            End If

        End If

    End If

    Call SetDevice(D3DDevice)

    Set Sprite = D3DX.CreateSprite(D3DDevice)
     
    'Set the scaling to default aspect ratio
    SpriteScaleVector.X = 1
    SpriteScaleVector.Y = 1

    Call setup_ambient

    IluRGB.R = 255
    IluRGB.G = 255
    IluRGB.b = 255

    Iluminacion = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)
    ColorTecho = Iluminacion

    bAlpha = 255
    nAlpha = 100

    Set pRenderTexture = D3DX.CreateTexture(D3DDevice, 480, 80, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    Set pRenderSurface = pRenderTexture.GetSurfaceLevel(0)
    Set pBackbuffer = D3DDevice.GetRenderTarget

    IniciarD3D = True

End Function

Public Sub SetDevice(D3DD As Direct3DDevice8)

    With D3DD
        .SetVertexShader FVF

        'Set the render states
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, True
        .SetRenderState D3DRS_ZWRITEENABLE, True
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

        'Particle engine settings
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        'Set the texture stage stats (filters)
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT

    End With

End Sub

Public Sub CargarFont()

    Dim I As Integer

    Open (PathInit & "\Font.ind") For Binary As #1

    For I = 1 To 255
        Get #1, , Caracteres(I)
    Next I

    Close #1

    Dim hFont As Long

    fntCartel.Name = "Augusta"
    fntCartel.Size = 14
    fntCartel.bold = False
    Set FontCartelDesc = fntCartel

    fnt.Name = "Augusta"
    fnt.Size = 48
    fnt.bold = False
    Set MainFontDesc = fnt

    fnt2.Name = "Augusta"
    fnt2.Size = 72
    fnt2.bold = False
    Set MainFontBigDesc = fnt2
    
    fntRender.Name = "Arial"
    fntRender.Size = 22
    fntRender.bold = False
    Set FontRenderDesc = fntRender
    'hFont = GetStockObject(ANSI_FIXED_FONT)
    
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
    Set MainFontBig = D3DX.CreateFont(D3DDevice, MainFontBigDesc.hFont)
    Set FontCartel = D3DX.CreateFont(D3DDevice, FontCartelDesc.hFont)
Set FontRender = D3DX.CreateFont(D3DDevice, FontRenderDesc.hFont)
End Sub

Public Sub DrawFont(Texto As String, _
                    ByVal X As Long, _
                    ByVal Y As Long, _
                    ByVal color As Long, _
                    Optional Centrado As Boolean = False)

    Dim I     As Integer

    Dim SumaX As Integer

    Dim SumaL As Integer

    Dim CharC As Byte

    If Centrado Then

        For I = 1 To Len(Texto)
            CharC = Asc(mid$(Texto, I, 1))
            SumaL = SumaL + Caracteres(CharC).x2 - 2
        Next I

        SumaL = SumaL / 2

    End If

    For I = 1 To Len(Texto)
        CharC = Asc(mid$(Texto, I, 1))
        'Call Engine_Render_D3DXSprite(X - SumaL + SumaX, Y, Caracteres(CharC).X2 + 2, Caracteres(CharC).Y2 + 2, Caracteres(CharC).X + 1, Caracteres(CharC).Y + 1, Color, 14324, 0)
        Call Engine_Render_Rectangle(X - SumaL + SumaX, Y, Caracteres(CharC).x2 + 2, Caracteres(CharC).y2 + 2, Caracteres(CharC).X + 1, Caracteres(CharC).Y + 1, Caracteres(CharC).x2 + 2, Caracteres(CharC).y2 + 2, , , , 14324, color, color, color, color)
    
        SumaX = SumaX + Caracteres(CharC).x2 - 2
    Next I

End Sub

Public Function IniciarDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean

    On Error GoTo ErrOut

    Dim DispMode  As D3DDISPLAYMODE

    Dim D3DWindow As D3DPRESENT_PARAMETERS

    UseMotionBlur = 1

    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    D3DWindow.Windowed = 1
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.BackBufferFormat = DispMode.Format
    D3DWindow.EnableAutoDepthStencil = 0
    D3DWindow.AutoDepthStencilFormat = D3DFMT_A8R8G8B8

    'If UseMotionBlur Then
    '    D3DWindow.EnableAutoDepthStencil = 1
    '    D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
    'End If
    frmMain.SetRender (True)

    RectJuego.x1 = 0
    RectJuego.y1 = 0
    #If RenderFull = 0 Then
        RectJuego.x2 = 1024
        RectJuego.y2 = 782

        If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
        Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.pRender.hwnd, D3DCREATEFLAGS, D3DWindow)
    #Else
        RectJuego.x1 = 0
        RectJuego.y1 = 0
        RectJuego.x2 = 800
        RectJuego.y2 = 608

        If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
        Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.pRender.hwnd, D3DCREATEFLAGS, D3DWindow)

    #End If

    If UseMotionBlur Then
        'Set DeviceBuffer = D3DDevice.GetRenderTarget
        'Set DeviceStencil = D3DDevice.GetDepthStencilSurface
        'Set BlurStencil = D3DDevice.CreateDepthStencilSurface(800, 600, D3DFMT_D16, D3DMULTISAMPLE_NONE)
        'Set BlurTexture = D3DX.CreateTexture(D3DDevice, 1024, 1024, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
        #If RenderFull = 0 Then
            Set BlurTexture = D3DX.CreateTexture(D3DDevice, 1024, 782, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
        #Else
            Set BlurTexture = D3DX.CreateTexture(D3DDevice, 800, 608, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
        #End If

        Set BlurSurf = BlurTexture.GetSurfaceLevel(0)

        Dim T As Long

        'Create the motion-blur vertex array
        For T = 0 To 3
            BlurTA(T).color = D3DColorXRGB(255, 255, 255)
            BlurTA(T).rhw = 1
        Next T

        BlurTA(1).X = ScreenWidth
        BlurTA(2).Y = ScreenHeight
        BlurTA(3).X = ScreenWidth
        BlurTA(3).Y = ScreenHeight

    End If

    'Set the blur to off
    BlurIntensity = 255

    IniciarDevice = True
    Exit Function

ErrHandler:
    MessageBox "Su placa de video no es combatible. Este al tanto en la página web para parches que puedan solucionar este incomveniente.", vbCritical
    IniciarDevice = False

    Exit Function

ErrOut:

    'Destroy the D3DDevice so it can be remade
    Set D3DDevice = Nothing

    'Return a failure
    IniciarDevice = False

End Function

Public Sub LiberarObjetosDX()
    err.Clear

    On Error GoTo fin:

    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set dX = Nothing
    Exit Sub
fin:                                                                                                 MsgBox "Error producido en Public Sub LiberarObjetosDX()"

End Sub

Public Sub Engine_ReadyTexture(ByVal TextureNum As Integer)

    'Set the texture
    If TextureNum > 0 Then
        If LastTexture <> TextureNum Then
            Set Textura = SurfaceDB.Surface(TextureNum)
            D3DDevice.SetTexture 0, Textura
            LastTexture = TextureNum

        End If

    End If

    LastTexture = TextureNum

End Sub

Public Sub Engine_Render_D3DXSprite(ByVal X As Single, _
                                    ByVal Y As Single, _
                                    ByVal Width As Single, _
                                    ByVal Height As Single, _
                                    ByVal srcX As Single, _
                                    ByVal srcY As Single, _
                                    ByVal Light As Long, _
                                    ByVal TextureNum As Long, _
                                    ByVal Degrees As Single)

    Dim srcRect As RECT

    Dim v2      As D3DVECTOR2

    Dim v3      As D3DVECTOR2
    
    'Ready the texture
    Engine_ReadyTexture TextureNum
    
    'Create the source rectangle
    With srcRect
        .left = srcX
        .top = srcY
        .right = .left + Width
        .bottom = .top + Height

    End With
    
    'Create the rotation point
    If Degrees Then
        Degrees = ((Degrees + 180) * DegreeToRadian)

        If Degrees > 360 Then Degrees = Degrees - 360

        With v2
            .X = (Width * 0.5)
            .Y = (Height * 0.5)

        End With

    End If
    
    'Set the translation (location on the screen)
    v3.X = X
    v3.Y = Y

    'Draw the sprite
    If TextureNum > 0 Then
        Sprite.Draw Textura, srcRect, SpriteScaleVector, v2, Degrees, v3, Light
    Else

        'Sprite.Draw Nothing, SrcRect, SpriteScaleVector, v2, 0, v3, Light
    End If
    
End Sub

Public Sub Engine_Render_D3DXTexture(ByVal X As Single, _
                                     ByVal Y As Single, _
                                     ByVal Width As Single, _
                                     ByVal Height As Single, _
                                     ByVal srcX As Single, _
                                     ByVal srcY As Single, _
                                     ByVal Light As Long, _
                                     ByVal texture As Direct3DTexture8, _
                                     ByVal Degrees As Single)

    Dim srcRect As RECT

    Dim v2      As D3DVECTOR2

    Dim v3      As D3DVECTOR2
    
    'Ready the texture
    'D3DDevice.SetTexture 0, Texture
    LastTexture = 0
    
    'Create the source rectangle
    With srcRect
        .left = srcX
        .top = srcY
        .right = .left + Width
        .bottom = .top + Height

    End With
    
    'Create the rotation point
    If Degrees Then
        Degrees = ((Degrees + 180) * DegreeToRadian)

        If Degrees > 360 Then Degrees = Degrees - 360

        With v2
            .X = (Width * 0.5)
            .Y = (Height * 0.5)

        End With

    End If
    
    'Set the translation (location on the screen)
    v3.X = X
    v3.Y = Y

    'Draw the sprite
    Sprite.Draw texture, srcRect, SpriteScaleVector, v2, Degrees, v3, Light
    
End Sub

Sub Engine_Render_Rectangle(ByVal X As Single, _
                            ByVal Y As Single, _
                            ByVal Width As Single, _
                            ByVal Height As Single, _
                            ByVal srcX As Single, _
                            ByVal srcY As Single, _
                            ByVal SrcWidth As Single, _
                            ByVal SrcHeight As Single, _
                            Optional ByVal SrcBitmapWidth As Long = -1, _
                            Optional ByVal SrcBitmapHeight As Long = -1, _
                            Optional ByVal Degrees As Single = 0, _
                            Optional ByVal TextureNum As Long, _
                            Optional ByVal Color0 As Long = -1, _
                            Optional ByVal Color1 As Long = -1, _
                            Optional ByVal Color2 As Long = -1, _
                            Optional ByVal Color3 As Long = -1, _
                            Optional ByVal Shadow As Byte = 0, _
                            Optional ByVal InBoundsCheck As Boolean = True)

    '************************************************************
    'Render a square/rectangle based on the specified values then rotate it if needed
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_Rectangle
    '************************************************************
    Dim VertexArray(0 To 3) As TLVERTEX

    Dim RadAngle            As Single 'The angle in Radians

    Dim CenterX             As Single

    Dim CenterY             As Single

    Dim Index               As Integer

    Dim NewX                As Single

    Dim NewY                As Single

    Dim SinRad              As Single

    Dim CosRad              As Single

    Dim ShadowAdd           As Single

    Dim l                   As Single
       
    Width = Width
    Height = Height

    'Perform in-bounds check if needed
    If InBoundsCheck Then
        If X + SrcWidth <= 0 Then Exit Sub
        If Y + SrcHeight <= 0 Then Exit Sub
        If X >= frmMain.pRender.Width Then Exit Sub
        If Y >= frmMain.pRender.Height Then Exit Sub

    End If

    'Ready the texture
    Engine_ReadyTexture TextureNum

    'Set the bitmap dimensions if needed
    If SrcBitmapWidth = -1 Then SrcBitmapWidth = SurfaceSize(TextureNum).X
    If SrcBitmapHeight = -1 Then SrcBitmapHeight = SurfaceSize(TextureNum).Y
    
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).color = Color0
    VertexArray(1).color = Color1
    VertexArray(2).color = Color2
    VertexArray(3).color = Color3

    If Shadow Then

        'To make things easy, we just do a completely separate calculation the top two points
        ' with an uncropped tU / tV algorithm
        VertexArray(0).X = X + (Width * 0.5)
        VertexArray(0).Y = Y - (Height * 0.5)
        VertexArray(0).tu = (srcX / SrcBitmapWidth)
        VertexArray(0).tv = (srcY / SrcBitmapHeight)
        
        VertexArray(1).X = VertexArray(0).X + Width
        VertexArray(1).tu = ((srcX + Width) / SrcBitmapWidth)

        VertexArray(2).X = X
        VertexArray(2).tu = (srcX / SrcBitmapWidth)

        VertexArray(3).X = X + Width
        VertexArray(3).tu = (srcX + SrcWidth + ShadowAdd) / SrcBitmapWidth

    Else
        
        '------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------
        'If the image is partially outside of the screen, it is trimmed so only that which is in the screen is drawn
        'This provides for quite a decent FPS boost if you have lots of tiles that stretch outside of the view area
        'Important: Something about this doesn't seem to be functioning correctly. It is supposed to crop down the
        'image and only draw that which is going to be in the screen, but it doesn't work right and I have no
        'idea why. Uncomment the lines to see what happens. I have given up on this since the FPS boost really isn't
        'significant for me to put any more work into it, but if someone could fix it, it would definitely be
        'added back into the engine.
        '------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------
        'If X < 0 Then
        '    SrcX = SrcX - X
        '    SrcWidth = SrcWidth + X
        '    Width = Width + X
        '    X = 0
        'End If
        'If Y < 0 Then
        '    SrcY = SrcY - Y
        '    SrcHeight = SrcHeight + Y
        '    Height = Height + Y
        '    Y = 0
        'End If
        'If X + Width > ScreenWidth Then
        '    L = X + Width - ScreenWidth
        '    Width = Width - L
        '    SrcWidth = SrcWidth - L
        'End If
        'If Y + Height > ScreenHeight Then
        '    L = Y + Height - ScreenHeight
        '    Height = Height - L
        '    SrcHeight = SrcHeight - L
        'End If
        '------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------
        
        'If we are NOT using shadows, then we add +1 to the width/height (trust me, just do it)
        ShadowAdd = 1

        'Find the left side of the rectangle
        VertexArray(0).X = X

        If SrcBitmapWidth = 0 Then Exit Sub
        VertexArray(0).tu = (srcX / SrcBitmapWidth)

        'Find the top side of the rectangle
        VertexArray(0).Y = Y
        VertexArray(0).tv = (srcY / SrcBitmapHeight)
    
        'Find the right side of the rectangle
        VertexArray(1).X = X + Width
        VertexArray(1).tu = (srcX + SrcWidth + ShadowAdd) / SrcBitmapWidth
 
        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X

    End If
    
    'Find the bottom of the rectangle
    VertexArray(2).Y = Y + Height
    VertexArray(2).tv = (srcY + SrcHeight + ShadowAdd) / SrcBitmapHeight

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).Y = VertexArray(0).Y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).Y = VertexArray(2).Y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv
    
    'Check if a rotation is required
    If Degrees Mod 360 <> 0 Then

        'Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        'Set the CenterX and CenterY values
        CenterX = X + (Width * 0.5)
        CenterY = Y + (Height * 0.5)

        'Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        'Loops through the passed vertex buffer
        For Index = 0 To 3

            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (VertexArray(Index).X - CenterX) * CosRad - (VertexArray(Index).Y - CenterY) * SinRad
            NewY = CenterY + (VertexArray(Index).Y - CenterY) * CosRad + (VertexArray(Index).X - CenterX) * SinRad

            'Applies the new co-ordinates to the buffer
            VertexArray(Index).X = NewX
            VertexArray(Index).Y = NewY

        Next Index

    End If

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub

Public Function EngineSpeed() As Single

    EngineSpeed = 0.0176

    If UserEquitando Then
        EngineSpeed = EngineSpeed + 0.002

    End If
    
    If UserCongelado Then
        EngineSpeed = EngineSpeed - 0.005

    End If
    
    If UserChiquito Then
        EngineSpeed = EngineSpeed + 0.001

    End If
    
End Function

Public Function DamePasos(ByVal Heading As Byte) As Integer

    If UserNavegando Then Exit Function
    'If UserEstado = 1 Then Exit Function

    Select Case Heading

        Case E_Heading.north
            DamePasos = 26646

            '62
        Case E_Heading.east
            DamePasos = 26644

            '63
        Case E_Heading.west
            DamePasos = 26645

            '64
        Case E_Heading.south
            DamePasos = 26647

            '61
    End Select
    
End Function

Public Sub RenderCuentaRegresiva()

    Dim color        As Long

    Static last_tick As Long
    
    If AlphaCuenta > 0 Then
        If (GetTickCount And &H7FFFFFFF) - last_tick >= 18 Then
            AlphaCuenta = AlphaCuenta - 5
            last_tick = (GetTickCount And &H7FFFFFFF)

        End If

    End If
    
    color = D3DColorRGBA(255, 255, 255, AlphaCuenta)

    Call Engine_Render_D3DXSprite(525, 410, 256, 256, 0, 0, color, 4000 + CUENTA, 0)

End Sub

Public Sub RenderUserDieBlood()

    Dim color        As Long

    Static last_tick As Long
    
    If AlphaBloodUserDie > 0 Then
        If (GetTickCount And &H7FFFFFFF) - last_tick >= 18 Then
            AlphaBloodUserDie = AlphaBloodUserDie - 3
            last_tick = (GetTickCount And &H7FFFFFFF)

        End If

    End If
    
    color = D3DColorRGBA(255, 255, 255, AlphaBloodUserDie)
    'helios Grafico Muerto 06/06/2021
    Call Engine_Render_D3DXSprite(250, 70, 1024, 768, 0, 0, color, 4011, 0)

End Sub

Public Sub RenderBlood()

    Dim color        As Long

    Static last_tick As Long
    
    If AlphaBlood > 0 Then
        If (GetTickCount And &H7FFFFFFF) - last_tick >= 18 Then
            AlphaBlood = AlphaBlood - 5
            last_tick = (GetTickCount And &H7FFFFFFF)

        End If

    End If
    
    color = D3DColorRGBA(255, 255, 255, AlphaBlood)

    Call Engine_Render_D3DXSprite(250, 0, 1024, 768, 0, 0, color, 3999, 0)

End Sub

Public Sub RenderCeguera()

    Dim color        As Long

    Static last_tick As Long
    
    If AlphaCeguera > 10 Then
        If (GetTickCount And &H7FFFFFFF) - last_tick >= 18 Then
            If AlphaCeguera > 249 Then
                AlphaCeguera = AlphaCeguera - 0.2
            Else
                AlphaCeguera = AlphaCeguera - 10

            End If

            last_tick = (GetTickCount And &H7FFFFFFF)

        End If

    Else
        AlphaCeguera = 255

    End If
    
    color = D3DColorRGBA(0, 0, 0, AlphaCeguera)

    Call Engine_Render_D3DXSprite(255, 255, 1024, 768, 0, 0, color, 14706, 0)
    'Call Engine_Render_D3DXSprite(400, 325, 1024, 768, 0, 0, color, 4011, 0)

End Sub

Public Sub RenderTextKills()

    Dim color        As Long

    Static last_tick As Long
    
    If AlphaTextKills > 0 Then
        If (GetTickCount And &H7FFFFFFF) - last_tick >= 18 Then
            AlphaTextKills = AlphaTextKills - 5
            last_tick = (GetTickCount And &H7FFFFFFF)

        End If

    End If
    
    color = D3DColorRGBA(255, 255, 255, AlphaTextKills)
    
    Call Engine_Render_D3DXSprite(525, 450, 257, 58, 0, 0, color, 4130 + TextKillsType, 0)

End Sub

Public Sub ReproducirSonidosDeAmbiente()

    If UserCharIndex <= 0 Then Exit Sub

    If Zonas(ZonaActual).CantSonidos > 0 Then

        Static last_tick As Long
    
        If (GetTickCount And &H7FFFFFFF) - last_tick >= 12000 Then

            With charlist(UserCharIndex)
                Call Audio.PlayWave(Zonas(ZonaActual).Sonido(RandomNumber(1, Zonas(ZonaActual).CantSonidos)), RandomNumber(.Pos.X - 10, .Pos.X + 10), RandomNumber(.Pos.Y - 9, .Pos.Y + 9))
                last_tick = (GetTickCount And &H7FFFFFFF)

            End With

        End If

    End If

End Sub

Public Sub RenderRelampago()

    'Dim color As Long
    Static last_tick As Long

    Dim color        As Long
    
    If AlphaRelampago > 0 Then
        If (GetTickCount And &H7FFFFFFF) - last_tick >= 18 Then
            AlphaRelampago = AlphaRelampago - RandomNumber(15, 25)
            last_tick = (GetTickCount And &H7FFFFFFF)

        End If

    Else

        If HayRelampago = True Then
            Hora = OrigHora
            Call SetDayLight(False)
            HayRelampago = False

        End If

    End If
  
    'color = D3DColorRGBA(255, 255, 255, AlphaRelampago)
    'Call Engine_Render_D3DXSprite(randomRelampagoX2, randomRelampagoY2, 400, 313, 0, 0, color, 14783, 0)
End Sub

Public Sub RenderSaliendo()

    Dim color        As Long

    Static last_tick As Long
    
    If AlphaSalir > 0 And AlphaSalir < 255 Then
        If (GetTickCount And &H7FFFFFFF) - last_tick >= 32 Then
            AlphaSalir = AlphaSalir + 1
            last_tick = (GetTickCount And &H7FFFFFFF)

        End If

    Else
        AlphaSalir = 255

    End If
    
    color = D3DColorRGBA(0, 0, 0, AlphaSalir)

    Call Engine_Render_D3DXSprite(0, 0, 1024, 768, 0, 0, color, 14706, 0)
    
End Sub

