Attribute VB_Name = "modIconMouse"
Option Explicit

Public Enum ModosDeStretch

    BlackOnWhite = 1
    WhiteOnBlack = 2
    ColorOnColor = 3
    Halftone = 4
    Desconocida = 5

End Enum

Private Type PICTDESC

    cbSizeOfStruct As Long
    PicType As Long
    hgdiObj As Long
    hPalOrXYExt As Long

End Type

Private Type IID

    data1 As Long
    data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte

End Type

Private Declare Sub OleCreatePictureIndirect _
                Lib "oleaut32.dll" (lpPictDesc As PICTDESC, _
                                    riid As IID, _
                                    ByVal fOwn As Boolean, _
                                    lplpvObj As Object)
    
Private Type pvICONINFO

    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long

End Type

Private Type pvRECT

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As pvICONINFO) As Long

Private Declare Function CreateBitmap _
                Lib "gdi32" (ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal nPlanes As Long, _
                             ByVal nBitCount As Long, _
                             lpBits As Any) As Long

Private Declare Function SetStretchBltMode _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nStretchMode As Long) As Long

Private Declare Function StretchBlt _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal nSrcWidth As Long, _
                             ByVal nSrcHeight As Long, _
                             ByVal dwRop As Long) As Long

Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal dwRop As Long) As Long

Private Declare Function CreateCompatibleBitmap _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function ReleaseDC _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hDC As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function FillRect _
                Lib "user32" (ByVal hDC As Long, _
                              lppvRECT As pvRECT, _
                              ByVal hBrush As Long) As Long

Private Declare Function SetBkColor _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal Color As Long) As Long

Private Declare Function SetTextColor _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal Color As Long) As Long

Const DI_MASK = &H1

Const DI_IMAGE = &H2

Public Sub pvBMPaICO(srchDC As Long, _
                     ByVal x As Integer, _
                     ByVal y As Integer, _
                     ByVal hImagen As Long, _
                     ByVal hMask As Long, _
                     ByVal hDC As Long, _
                     ByVal ScrDC As Long, _
                     ByVal MaskColor As Long, _
                     ByVal ModoDeStretch As Byte)

    Dim R       As pvRECT, hBr As Long

    Dim hOldPal As Long, hDC_Copia As Long

    Dim TmpBMP  As Long

    R.Bottom = 32
    R.Right = 32
    
    SetStretchBltMode hDC, ModoDeStretch
            
    ' Dibujo la mascara
    
    If MaskColor = -1 Then    ' No hay transparencia
        
        ' Selecciono la mascara...
        SelectObject hDC, hMask
        
        ' ... y la lleno con negro (opaco)
        hBr = CreateSolidBrush(&H0)
        FillRect hDC, R, hBr
        DeleteObject hBr
    
        ' Selecciono la imagen
        SelectObject hDC, hImagen
    
        StretchBlt hDC, 0, 0, 32, 32, srchDC, x, y, 32, 32, vbSrcCopy
               
    Else

        ' Creo un DC y un bitmap para
        ' copiar la imagen. Esto lo
        ' debo hacer porque si el bitmap
        ' es DIB no pasa a B&N usando
        ' los colores de fondo y texto.
        
        hDC_Copia = CreateCompatibleDC(ScrDC)
        
        SetStretchBltMode hDC_Copia, ModoDeStretch
        
        TmpBMP = CreateCompatibleBitmap(ScrDC, 32, 32)
                
        ' Hago la copia del bitmap
        SelectObject hDC_Copia, TmpBMP
        
        StretchBlt hDC_Copia, 0, 0, 32, 32, srchDC, x, y, 32, 32, vbSrcCopy
              
        ' De ahora en mas utilizo la copia
        ' de la que ya a sido modificado su
        ' tama~o
              
        ' ---- Creo la mascara -----
        
        ' Selecciono la mascara en el DC
        SelectObject hDC, hMask
        
        ' Seteo el color de fondo con
        ' el color de mascara.
        SetBkColor hDC_Copia, MaskColor
        SetTextColor hDC_Copia, vbWhite
        
        ' Al copiar windows transforma en blanco
        ' todos los pixel con el color de fondo
        ' y en negro el resto
        BitBlt hDC, 0, 0, 32, 32, hDC_Copia, 0, 0, vbSrcCopy
          
        SelectObject hDC, hImagen
        SelectObject hDC_Copia, hMask

        hBr = CreateSolidBrush(&H0)
        FillRect hDC, R, hBr
        DeleteObject hBr
        
        ' Copio la mascara y luego la imagen
        BitBlt hDC, 0, 0, 32, 32, hDC_Copia, 0, 0, vbNotSrcCopy
        BitBlt hDC, 0, 0, 32, 32, srchDC, 0, 0, vbSrcAnd
            
        DeleteDC hDC_Copia
        DeleteObject TmpBMP
        
    End If

End Sub

Public Function HandleToPicture(ByVal hGDIHandle As Long, _
                                ByVal ObjectType As PictureTypeConstants, _
                                Optional ByVal hPal As Long = 0) As StdPicture

    Dim iPic As IPicture, picdes As PICTDESC, iidIPicture As IID
    
    ' Fill picture description
    picdes.cbSizeOfStruct = Len(picdes)
    picdes.PicType = ObjectType
    picdes.hgdiObj = hGDIHandle
    picdes.hPalOrXYExt = hPal
    
    ' IPictureDisp {7BF80981-BF32-101A-8BBB-00AA00300CAB}
    iidIPicture.data1 = &H7BF80981

    iidIPicture.data2 = &HBF32

    iidIPicture.Data3 = &H101A

    iidIPicture.Data4(0) = &H8B

    iidIPicture.Data4(1) = &HBB

    iidIPicture.Data4(2) = &H0

    iidIPicture.Data4(3) = &HAA

    iidIPicture.Data4(4) = &H0

    iidIPicture.Data4(5) = &H30

    iidIPicture.Data4(6) = &HC

    iidIPicture.Data4(7) = &HAB
    
    ' Crea el objeto con el handle
    OleCreatePictureIndirect picdes, iidIPicture, True, iPic
    
    Set HandleToPicture = iPic
        
End Function

Public Function GetIcon(ByVal srchDC As Long, _
                        ByVal srcX As Integer, _
                        ByVal srcY As Integer, _
                        Optional ModoDeStretch As ModosDeStretch = Halftone, _
                        Optional CrearCursor As Boolean = False, _
                        Optional MaskColor As Long = -1) As StdPicture

    Dim hIcon    As Long, IconPict As StdPicture

    Dim ScreenDC As Long, BitmapDC As Long

    Dim hMask    As Long, hImagen As Long

    Dim hIcn     As Long, II As pvICONINFO

    On Error Resume Next
   
    ScreenDC = GetWindowDC(0&)
    BitmapDC = CreateCompatibleDC(ScreenDC)
    
    hImagen = CreateCompatibleBitmap(ScreenDC, 32, 32)
    hMask = CreateBitmap(32, 32, 1, 1, ByVal 0&)
      
    pvBMPaICO srchDC, srcX, srcY, hImagen, hMask, BitmapDC, ScreenDC, MaskColor, ModoDeStretch
    
    DeleteDC BitmapDC
    ReleaseDC 0&, ScreenDC
    
    II.fIcon = CrearCursor
    II.hbmColor = hImagen
    II.hbmMask = hMask
    
    hIcon = CreateIconIndirect(II)
    
    Set IconPict = HandleToPicture(hIcon, vbPicTypeIcon)
     
    If IconPict Is Nothing Then
        
        DeleteObject hIcn
        Set GetIcon = Nothing
        
    Else
        
        Set GetIcon = IconPict
        
    End If
     
End Function

