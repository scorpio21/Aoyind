Attribute VB_Name = "modCompression"
Option Explicit

Public Const GRH_RESOURCE_FILE As String = "Graphics.AO"

'This structure will describe our binary file's
'size, number and version of contained files
Public Type FILEHEADER

    lngNumFiles As Long                 'How many files are inside?
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngFileVersion As Long              'The resource version (Used to patch)

End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER

    lngFileSize As Long             'How big is this chunk of stored data?
    lngFileStart As Long            'Where does the chunk start?
    lngRnd As Long
    strFileName As String * 40      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
    
    #If SeguridadAlkon Then
        lngCheckSum As Long
    #End If

End Type

Private Enum PatchInstruction

    Delete_File
    Create_File
    Modify_File

End Enum

Private Declare Function compress _
                Lib "zlib.dll" (dest As Any, _
                                destlen As Any, _
                                src As Any, _
                                ByVal srclen As Long) As Long

Private Declare Function uncompress _
                Lib "zlib.dll" (dest As Any, _
                                destlen As Any, _
                                src As Any, _
                                ByVal srclen As Long) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef dest As Any, _
                                       ByRef Source As Any, _
                                       ByVal ByteCount As Long)

'BitMaps Strucures
Public Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

Public Type BITMAPINFOHEADER

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

Public Type RGBQUAD

    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte

End Type

Public Type BITMAPINFO

    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD

End Type

Private Const BI_RGB       As Long = 0

Private Const BI_RLE8      As Long = 1

Private Const BI_RLE4      As Long = 2

Private Const BI_BITFIELDS As Long = 3

Private Const BI_JPG       As Long = 4

Private Const BI_PNG       As Long = 5

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace _
                Lib "kernel32" _
                Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
                                             FreeBytesToCaller As Currency, _
                                             bytesTotal As Currency, _
                                             FreeBytesTotal As Currency) As Long

Private Const GDIP_OK As Long = 0

Private Type GdiplusStartupInput

    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long

End Type

Private Declare Function GdiplusStartup _
                Lib "gdiplus" (ByRef Token As Long, _
                               ByRef inputbuf As GdiplusStartupInput, _
                               ByVal pOutputBuf As Long) As Long

Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long

Private Declare Function GdipCreateHBITMAPFromBitmap _
                Lib "gdiplus" (ByVal BITMAP As Long, _
                               ByRef hbmReturn As Long, _
                               ByVal background As Long) As Long

Private Declare Function GdipCreateBitmapFromFile _
                Lib "gdiplus" (ByVal pFileName As Long, _
                               ByRef BITMAP As Long) As Long

Private Declare Function GdipCreateBitmapFromStream _
                Lib "gdiplus" (ByVal Stream As IUnknown, _
                               ByRef BITMAP As Long) As Long

Private Declare Function GdipGetImageDimension _
                Lib "gdiplus" (ByVal Image As Long, _
                               ByRef Width As Single, _
                               ByRef Height As Single) As Long

Public GdipToken       As Long

Public GdipInitialized As Boolean

Public GdipClients     As Long

Private Declare Function GlobalAlloc _
                Lib "kernel32" (ByVal uFlags As Long, _
                                ByVal dwBytes As Long) As Long
    
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    
Private Const S_OK As Long = 0

Private Declare Function CreateStreamOnHGlobal _
                Lib "ole32" (ByVal hGlobal As Long, _
                             ByVal fDeleteOnRelease As Long, _
                             ByRef Stream As Object) As Long

Private Declare Function OleTranslateColor _
                Lib "oleaut32" (ByVal clr As OLE_COLOR, _
                                ByVal hPal As Long, _
                                ByRef colorref As Long) As Long

Private Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 6/07/2004
    '
    '**************************************************************
    Dim retval As Long

    Dim FB     As Currency

    Dim BT     As Currency

    Dim FBT    As Currency
    
    retval = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes

End Function

''
' Sorts the info headers by their file name. Uses QuickSort.
'
' @param    InfoHead() The array of headers to be ordered.
' @param    first The first index in the list.
' @param    last The last index in the list.

Private Sub Sort_Info_Headers(ByRef InfoHead() As INFOHEADER, _
                              ByVal First As Long, _
                              ByVal Last As Long)

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/20/2007
    'Sorts the info headers by their file name using QuickSort.
    '*****************************************************************
    Dim aux  As INFOHEADER

    Dim min  As Long

    Dim max  As Long

    Dim comp As String
    
    min = First
    max = Last
    
    comp = InfoHead((min + max) \ 2).strFileName
    
    Do While min <= max
        Do While InfoHead(min).strFileName < comp And min < Last
            min = min + 1
        Loop

        Do While InfoHead(max).strFileName > comp And max > First
            max = max - 1
        Loop

        If min <= max Then
            aux = InfoHead(min)
            InfoHead(min) = InfoHead(max)
            InfoHead(max) = aux
            min = min + 1
            max = max - 1

        End If

    Loop
    
    If First < max Then Call Sort_Info_Headers(InfoHead, First, max)
    If min < Last Then Call Sort_Info_Headers(InfoHead, min, Last)

End Sub

''
' Searches for the specified InfoHeader.
'
' @param    ResourceFile A handler to the data file.
' @param    InfoHead The header searched.
' @param    FirstHead The first head to look.
' @param    LastHead The last head to look.
' @param    FileHeaderSize The bytes size of a FileHeader.
' @param    InfoHeaderSize The bytes size of a InfoHeader.
'
' @return   True if found.
'
' @remark   File must be already open.
' @remark   InfoHead must have set its file name to perform the search.

Private Function BinarySearch(ByRef ResourceFile As Integer, _
                              ByRef InfoHead As INFOHEADER, _
                              ByVal FirstHead As Long, _
                              ByVal LastHead As Long, _
                              ByVal FileHeaderSize As Long, _
                              ByVal InfoHeaderSize As Long) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/21/2007
    'Searches for the specified InfoHeader
    '*****************************************************************
    Dim ReadingHead  As Long

    Dim ReadInfoHead As INFOHEADER
    
    Do Until FirstHead > LastHead
        ReadingHead = (FirstHead + LastHead) \ 2

        Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + 1, ReadInfoHead

        If InfoHead.strFileName = ReadInfoHead.strFileName Then
            InfoHead = ReadInfoHead
            BinarySearch = True
            Exit Function
        Else

            If InfoHead.strFileName < ReadInfoHead.strFileName Then
                LastHead = ReadingHead - 1
            Else
                FirstHead = ReadingHead + 1

            End If

        End If

    Loop

End Function

''
' Retrieves the InfoHead of the specified graphic file.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    InfoHead The InfoHead where data is returned.
'
' @return   True if found.

Private Function Get_InfoHeader(ByRef ResourcePath As String, _
                                ByRef FileName As String, _
                                ByRef InfoHead As INFOHEADER) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/21/2007
    'Retrieves the InfoHead of the specified graphic file
    '*****************************************************************
    Dim ResourceFile     As Integer

    Dim ResourceFilePath As String

    Dim FileHead         As FILEHEADER
    
    On Local Error GoTo ErrHandler

    ResourceFilePath = ResourcePath
    
    'Set InfoHeader we are looking for
    InfoHead.strFileName = UCase$(FileName)

    Call Secure_Info_Header(InfoHead)
    
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    'Extract the FILEHEADER
    Get ResourceFile, 1, FileHead

    Call Secure_File_Header(FileHead)
        
    'Check the file for validity
    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        MsgBox "Archivo de recursos dañado. " & ResourceFilePath, , "Error"
        Close ResourceFile
        Exit Function

    End If
        
    'Search for it!
    If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, Len(FileHead), Len(InfoHead)) Then

        Call Secure_Info_Header(InfoHead)
            
        Get_InfoHeader = True

    End If
        
    Close ResourceFile
    Exit Function

ErrHandler:
    Close ResourceFile
    
    Call MsgBox("Error al intentar leer el archivo " & ResourceFilePath & ". Razón: " & err.Number & " : " & err.Description, vbOKOnly, "Error")

End Function

''
' Compresses binary data avoiding data loses.
'
' @param    data() The data array.

Private Sub Compress_Data(ByRef Data() As Byte)

    '*****************************************************************
    'Author: Juan Martín Dotuyo Dodero
    'Last Modify Date: 10/13/2004
    'Compresses binary data avoiding data loses
    '*****************************************************************
    Dim Dimensions As Long

    Dim DimBuffer  As Long

    Dim BufTemp()  As Byte

    Dim LoopC      As Long
    
    Dimensions = UBound(Data) + 1
    
    ' The worst case scenario, compressed info is 1.06 times the original - see zlib's doc for more info.
    DimBuffer = Dimensions * 1.06
    
    ReDim BufTemp(DimBuffer)
    
    Call compress(BufTemp(0), DimBuffer, Data(0), Dimensions)
    
    Erase Data
    
    ReDim Data(DimBuffer - 1)
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    Data = BufTemp
    
    Erase BufTemp

    Call Secure_Compressed_Data(Data)

End Sub

''
' Decompresses binary data.
'
' @param    data() The data array.
' @param    OrigSize The original data size.

Private Sub Decompress_Data(ByRef Data() As Byte, ByVal OrigSize As Long)

    '*****************************************************************
    'Author: Juan Martín Dotuyo Dodero
    'Last Modify Date: 10/13/2004
    'Decompresses binary data
    '*****************************************************************
    Dim BufTemp() As Byte
    
    ReDim BufTemp(OrigSize - 1)

    Call Secure_Compressed_Data(Data)
    
    Call uncompress(BufTemp(0), OrigSize, Data(0), UBound(Data) + 1)
    
    ReDim Data(OrigSize - 1)
    
    Data = BufTemp
    
    Erase BufTemp

End Sub

''
' Retrieves a byte array with the compressed data from the specified file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   InfoHead must not be encrypted.
' @remark   Data is not desencrypted.

Public Function Get_File_RawData(ByRef ResourcePath As String, _
                                 ByRef InfoHead As INFOHEADER, _
                                 ByRef Data() As Byte) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/24/2007
    'Retrieves a byte array with the compressed data from the specified file
    '*****************************************************************
    Dim ResourceFilePath As String

    Dim ResourceFile     As Integer
    
    On Local Error GoTo ErrHandler
    ResourceFilePath = ResourcePath
    
    'Size the Data array
    ReDim Data(InfoHead.lngFileSize - 1)
    
    'Open the binary file
    ResourceFile = FreeFile
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    'Get the data
    Get ResourceFile, InfoHead.lngFileStart, Data
    'Close the binary file
    Close ResourceFile
    
    Get_File_RawData = True
    Exit Function

ErrHandler:
    Close ResourceFile

End Function

''
' Extract the specific file from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Extract_File(ByRef ResourcePath As String, _
                             ByRef InfoHead As INFOHEADER, _
                             ByRef Data() As Byte) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/20/2007
    'Extract the specific file from a resource file
    '*****************************************************************
    On Local Error GoTo ErrHandler
    
    If Get_File_RawData(ResourcePath, InfoHead, Data) Then

        'Decompress all data
        If InfoHead.lngFileSize < InfoHead.lngFileSizeUncompressed Then
            Call Decompress_Data(Data, InfoHead.lngFileSizeUncompressed)

        End If
        
        Extract_File = True

    End If

    Exit Function

ErrHandler:
    Call MsgBox("Error al intentar decodificar recursos. Razon: " & err.Number & " : " & err.Description, vbOKOnly, "Error")

End Function

''
' Retrieves a byte array with the specified file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Get_File_Data(ByRef TIPO As String, _
                              ByRef FileName As String, _
                              ByRef Data() As Byte) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/21/2007
    'Retrieves a byte array with the specified file data
    '*****************************************************************
    #If Debugging Then

        Dim ResourceFile As Integer

        ResourceFile = FreeFile

        Dim path As String
    
        If TIPO = "Graficos" Then
            path = PathGraficos
        ElseIf TIPO = "Wav" Then
            path = PathWav
        ElseIf TIPO = "Interface" Then
            path = PathInterface

        End If
        
        Open path & "\" & FileName For Binary Access Read Lock Write As ResourceFile
        'Get the data
        ReDim Data(LOF(ResourceFile) - 1)
        Get ResourceFile, , Data
        'Close the binary file
        Close ResourceFile
        Get_File_Data = True
    #Else

        Dim InfoHead As INFOHEADER
    
        If Get_InfoHeader(ResourcePath, UCase$(FileName), InfoHead) Then
            'Extract!
            Get_File_Data = Extract_File(ResourcePath, InfoHead, Data)
        Else
            Call MsgBox("No se se encontro el recurso " & FileName)

        End If

    #End If

End Function

Private Function LoadPictureBytes(ByRef SourceBytes() As Byte, _
                                  ByRef hBitmap As Long) As Boolean

    'Returns True upon success.
    Dim Size       As Long

    Dim hMem       As Long

    Dim lpMem      As Long

    Dim Stream     As Object

    Dim GdipSI     As GdiplusStartupInput

    Dim GdipToken  As Long

    Dim GdipBitmap As Long

    Dim GdipWidth  As Single

    Dim GdipHeight As Single
    
    If Not GdipInitialized Then
        GdipSI.GdiplusVersion = 1
        GdipInitialized = GdiplusStartup(GdipToken, GdipSI, 0) = GDIP_OK

    End If

    Size = UBound(SourceBytes) - LBound(SourceBytes) + 1

    If Size < 1 Then Exit Function
    
    hMem = GlobalAlloc(&H2&, Size)

    If hMem = 0 Then Exit Function
    
    lpMem = GlobalLock(hMem)

    If lpMem = 0 Then GoTo KernelCleanup
    
    CopyMemory ByVal lpMem, SourceBytes(LBound(SourceBytes)), Size
    GlobalUnlock hMem

    If CreateStreamOnHGlobal(hMem, Not 0, Stream) <> S_OK Then GoTo KernelCleanup
    
    If GdipCreateBitmapFromStream(Stream, GdipBitmap) <> GDIP_OK Then GoTo KernelCleanup
    
    If GdipGetImageDimension(GdipBitmap, GdipWidth, GdipHeight) = GDIP_OK Then
        'mWidth = GdipWidth
        'mHeight = GdipHeight
        LoadPictureBytes = GdipCreateHBITMAPFromBitmap(GdipBitmap, hBitmap, 0) = GDIP_OK

    End If

    GdipDisposeImage GdipBitmap

KernelCleanup:
    GlobalFree hMem

End Function

''
' Retrieves bitmap file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.

Public Function Get_Bitmap(ByRef ResourcePath As String, _
                           ByRef FileName As String, _
                           ByRef bmpInfo As BITMAPINFO, _
                           ByRef hDC As Long) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 11/30/2007
    'Retrieves bitmap file data
    '*****************************************************************
    Dim InfoHead   As INFOHEADER

    Dim rawData()  As Byte

    Dim offBits    As Long

    Dim bitmapSize As Long

    Dim colorCount As Long

    Dim Archivo    As Long

    #If Debugging = 1 Then

        Archivo = FreeFile()
        
        Open PathGraficos & "\" & FileName For Binary As Archivo
        ReDim rawData(LOF(Archivo))
        Get #Archivo, , rawData
        Close Archivo
        
        ' Dim hBitmap As Long
        
        Call LoadPictureBytes(rawData, hDC)
     
        'Set img = LoadPicEx(PathGraficos & "\" & FileName)
     
        'Set img = LoadGraphicEX(Filename)
    
        'bmpInfo.bmiHeader.biWidth = img.Width
        'bmpInfo.bmiHeader.biHeight = img.Height
    
        'hdc = img.handle

        'Call CopyMemory(offBits, rawData(10), 4)
        'Call CopyMemory(bmpInfo.bmiHeader, rawData(14), 40)
    
        'With bmpInfo.bmiHeader
        '    bitmapSize = AlignScan(.biWidth, .biBitCount) * Abs(.biHeight)
        '
        '    If .biBitCount < 24 Or .biCompression = BI_BITFIELDS Or (.biCompression <> BI_RGB And .biBitCount = 32) Then
        '        If .biClrUsed < 1 Then
        '            colorCount = 2 ^ .biBitCount
        '        Else
        '            colorCount = .biClrUsed
        '        End If
        '
        '        ' When using bitfields on 16 or 32 bits images, bmiColors has a 3-longs mask.
        '        If .biBitCount >= 16 And .biCompression = BI_BITFIELDS Then colorCount = 3
        '
        '        Call CopyMemory(bmpInfo.bmiColors(0), rawData(54), colorCount * 4)
        '    End If
        'End With
    
        'ReDim Data(bitmapSize - 1) As Byte
        'Call CopyMemory(Data(0), rawData(offBits), bitmapSize)
    
        Get_Bitmap = True
    #Else

        If Get_InfoHeader(ResourcePath, FileName, InfoHead) Then

            'Extract the file and create the bitmap data from it.
            If Extract_File(ResourcePath, InfoHead, rawData) Then
                Call CopyMemory(offBits, rawData(10), 4)
                Call CopyMemory(bmpInfo.bmiHeader, rawData(14), 40)
            
                With bmpInfo.bmiHeader
                    bitmapSize = AlignScan(.biWidth, .biBitCount) * Abs(.biHeight)
                
                    If .biBitCount < 24 Or .biCompression = BI_BITFIELDS Or (.biCompression <> BI_RGB And .biBitCount = 32) Then
                        If .biClrUsed < 1 Then
                            colorCount = 2 ^ .biBitCount
                        Else
                            colorCount = .biClrUsed

                        End If
                    
                        ' When using bitfields on 16 or 32 bits images, bmiColors has a 3-longs mask.
                        If .biBitCount >= 16 And .biCompression = BI_BITFIELDS Then colorCount = 3
                    
                        Call CopyMemory(bmpInfo.bmiColors(0), rawData(54), colorCount * 4)

                    End If

                End With
            
                ReDim Data(bitmapSize - 1) As Byte
                Call CopyMemory(Data(0), rawData(offBits), bitmapSize)
            
                Get_Bitmap = True

            End If

        Else
            Call MsgBox("No se encontro el recurso " & FileName)

        End If

    #End If

End Function

''
' Compare two byte arrays to detect any difference.
'
' @param    data1() Byte array.
' @param    data2() Byte array.
'
' @return   True if are equals.

Private Function Compare_Datas(ByRef data1() As Byte, ByRef data2() As Byte) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 02/11/2007
    'Compare two byte arrays to detect any difference
    '*****************************************************************
    Dim Length As Long

    Dim act    As Long
    
    Length = UBound(data1) + 1
    
    If (UBound(data2) + 1) = Length Then

        While act < Length

            If data1(act) Xor data2(act) Then Exit Function
            
            act = act + 1
        Wend
        
        Compare_Datas = True

    End If

End Function

''
' Retrieves the next InfoHeader.
'
' @param    ResourceFile A handler to the resource file.
' @param    FileHead The reource file header.
' @param    InfoHead The returned header.
' @param    ReadFiles The number of headers that have already been read.
'
' @return   False if there are no more headers tu read.
'
' @remark   File must be already open.
' @remark   Used to walk through the resource file info headers.
' @remark   The number of read files will increase although there is nothing else to read.
' @remark   InfoHead is encrypted.

Private Function ReadNext_InfoHead(ByRef ResourceFile As Integer, _
                                   ByRef FileHead As FILEHEADER, _
                                   ByRef InfoHead As INFOHEADER, _
                                   ByRef ReadFiles As Long) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/24/2007
    'Reads the next InfoHeader
    '*****************************************************************

    If ReadFiles < FileHead.lngNumFiles Then
        'Read header
        Get ResourceFile, Len(FileHead) + Len(InfoHead) * ReadFiles + 1, InfoHead
        
        'Update
        ReadNext_InfoHead = True

    End If
    
    ReadFiles = ReadFiles + 1

End Function

''
' Retrieves the next bitmap.
'
' @param    ResourcePath The resource file folder.
' @param    ReadFiles The number of bitmaps that have already been read.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   False if there are no more bitmaps tu get.
'
' @remark   Used to walk through the resource file bitmaps.

Public Function GetNext_Bitmap(ByRef ResourcePath As String, _
                               ByRef ReadFiles As Long, _
                               ByRef bmpInfo As BITMAPINFO, _
                               ByRef Data() As Byte, _
                               ByRef fileIndex As Long) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 12/02/2007
    'Reads the next InfoHeader
    '*****************************************************************
    On Error Resume Next

    Dim ResourceFile As Integer

    Dim FileHead     As FILEHEADER

    Dim InfoHead     As INFOHEADER

    Dim FileName     As String
    
    ResourceFile = FreeFile
    Open ResourcePath For Binary Access Read Lock Write As ResourceFile
    Get ResourceFile, 1, FileHead

    Call Secure_File_Header(FileHead)
    
    If ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ReadFiles) Then

        Call Secure_Info_Header(InfoHead)
        
        ' Call Get_Bitmap(ResourcePath, InfoHead.strFileName, bmpInfo, Data())
        FileName = Trim$(InfoHead.strFileName)
        fileIndex = CLng(Left$(FileName, Len(FileName) - 4))
        
        GetNext_Bitmap = True

    End If
    
    Close ResourceFile

End Function

Public Function GetNext_File(ByRef ResourcePath As String, _
                             ByRef ReadFiles As Long, _
                             ByRef Data() As Byte, _
                             ByRef fileIndex As Long) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 12/02/2007
    'Reads the next InfoHeader
    '*****************************************************************
    On Error Resume Next

    Dim ResourceFile As Integer

    Dim FileHead     As FILEHEADER

    Dim InfoHead     As INFOHEADER

    Dim FileName     As String
    
    ResourceFile = FreeFile
    Open ResourcePath For Binary Access Read Lock Write As ResourceFile
    Get ResourceFile, 1, FileHead

    Call Secure_File_Header(FileHead)
    
    If ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ReadFiles) Then

        Call Secure_Info_Header(InfoHead)
        
        Call Get_File_Data(ResourcePath, InfoHead.strFileName, Data())
        FileName = Trim$(InfoHead.strFileName)
        fileIndex = CLng(Left$(FileName, Len(FileName) - 4))
        
        GetNext_File = True

    End If
    
    Close ResourceFile

End Function

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    '*****************************************************************
    'Author: Unknown
    'Last Modify Date: Unknown
    '*****************************************************************
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8

End Function

''
' Retrieves the version number of a given resource file.
'
' @param    ResourceFilePath The resource file complete path.
'
' @return   The version number of the given file.

Public Function GetVersion(ByVal ResourceFilePath As String) As Long

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/23/2008
    '
    '*****************************************************************
    Dim ResourceFile As Integer

    Dim FileHead     As FILEHEADER
    
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    'Extract the FILEHEADER
    Get ResourceFile, 1, FileHead

    Call Secure_File_Header(FileHead)

    Close ResourceFile
    
    GetVersion = FileHead.lngFileVersion

End Function

Private Sub Secure_Compressed_Data(ByRef Data() As Byte)

    Dim I As Long

    For I = 0 To UBound(Data)
        Data(I) = Data(I) Xor 245 Xor 9
    Next I

End Sub

Private Sub Secure_Info_Header(ByRef Header As INFOHEADER)
    Header.lngFileSize = Header.lngFileSize Xor 6709
    Header.lngFileSizeUncompressed = Header.lngFileSizeUncompressed Xor 2147
    Header.lngFileStart = Header.lngFileStart Xor 4451
    Header.lngRnd = CLng(Rnd * 2147215225)

End Sub

Private Sub Secure_File_Header(ByRef Header As FILEHEADER)
    Header.lngFileSize = Header.lngFileSize Xor 6631
    Header.lngFileVersion = Header.lngFileVersion Xor 7782
    Header.lngNumFiles = Header.lngNumFiles Xor 9361

End Sub
