VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If Win64 Then

  'For bitmap save function
  Private Declare PtrSafe Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiCreateCompatibleDC Lib "gdi32" Alias "CreateCompatibleDC" (ByVal hdc As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiCreateCompatibleBitmap Lib "gdi32" Alias "CreateCompatibleBitmap" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
  Private Declare PtrSafe Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiBitBlt Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare PtrSafe Function apiGetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
  Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As LongPtr
  End Type
  Private Type BITMAPFILEHEADER
    bfType As Integer  'as Integer = misses 2 bytes alignment
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
  End Type
  Private Type BITMAPINFOHEADER '40 bytes
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
  Private Declare PtrSafe Function apiGetDIBits Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
  Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
  Private Declare PtrSafe Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long

  Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As LongPtr
  End Type
  Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
  End Type
  Private Type BITMAPINFOHEADER '40 bytes
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

#Else

  Private Type BITMAP '24 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As LongPtr
  End Type
  Private Type BITMAPFILEHEADER '16 bytes, not 14 bytes, padding of 2 bytes after bfType in 32 bits
    bfType As Integer           '2 bytes padding needed after here
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
  End Type
  Private Type BITMAPINFOHEADER '40 bytes
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

  Private Declare PtrSafe Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiCreateCompatibleDC Lib "gdi32" Alias "CreateCompatibleDC" (ByVal hdc As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiCreateCompatibleBitmap Lib "gdi32" Alias "CreateCompatibleBitmap" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
  Private Declare PtrSafe Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiBitBlt Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare PtrSafe Function apiGetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
  Private Declare PtrSafe Function apiGetDIBits Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
  Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
  Private Declare PtrSafe Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
  Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
  Private Const BI_RGB = 0&
  ' Global Memory Flags
  Private Const GMEM_MOVEABLE = &H2
  Private Const GMEM_ZEROINIT = &H40
  Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
  Private Declare PtrSafe Function apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As Long, lpRect As RECT) As Long
  
  Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus.dll" (ByVal psFilename As String, ByRef lpRetData As Long) As Long
  Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus.dll" (ByVal lpImageData As Long) As Long
  
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare PtrSafe Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As Long, ByVal lpBI As Long, ByVal wUsage As Long) As Long
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  
#End If

Private mabBitmap()     As Byte
Private mabiHeader()    As Byte
Private mtBitmapInfo    As BITMAPINFOHEADER

Private mlBitmapDataPtr As LongPtr

' IClassError implementation
Implements IClassError
Private mlErrNo   As Long
Private msErrCtx  As String
Private msErrDesc As String

Private Sub ClearErr()
  mlErrNo = 0&
  msErrCtx = ""
  msErrDesc = ""
End Sub

Private Sub SetErr(ByVal psErrCtx As String, ByVal plErrNum As Long, ByVal psErrDesc As String)
  mlErrNo = plErrNum
  msErrCtx = psErrCtx
  msErrDesc = psErrDesc
End Sub

Public Property Get LastErr() As Long
  LastErr = mlErrNo
End Property

Public Property Get LastErrDesc() As String
  LastErrDesc = msErrDesc
End Property

Public Property Get IIClassError() As IClassError
  Set IIClassError = Me
End Property

Private Property Get IClassError_LastErr() As Long
  IClassError_LastErr = mlErrNo
End Property

Private Property Get IClassError_LastErrCtx() As String
  IClassError_LastErrCtx = msErrCtx
End Property

Private Property Get IClassError_LastErrDesc() As String
  IClassError_LastErrDesc = msErrDesc
End Property

'--------------------------------------
Private Sub Class_Initialize()
  '/**/
End Sub

Private Sub Class_Terminate()
  '/**/
End Sub

Public Function LoadFromFile(ByVal psFilename As String) As Boolean
  Const LOCAL_ERR_CTX As String = "LoadFromFile"
  On Error GoTo LoadFromFile_Err
  ClearErr
  
  Dim lRet As LongPtr
  
  lRet = GdipCreateBitmapFromFile(psFilename, mlBitmapDataPtr)
  
  Stop '/**/
  
LoadFromFile_Exit:
  If mlBitmapDataPtr Then GdipDisposeImage mlBitmapDataPtr
  
  Exit Function
  
LoadFromFile_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LoadFromFile_Exit
End Function

'This is essentially a port of the MSDN sample:
'https://docs.microsoft.com/fr-fr/windows/win32/gdi/capturing-an-image
Public Function SaveConsoleAsBitmap(ByRef poConsole As CConsoul, ByVal psFilename As String, ByVal piStartLine As Integer, ByVal piEndLine As Integer) As Boolean
  Const LOCAL_ERR_CTX As String = "SaveConsoleAsBitmap"
  On Error GoTo SaveConsoleAsBitmap_Err
  ClearErr
  
  Dim hMemDC      As Long
  Dim hConsoleDC  As Long
  'Dim rcConsole   As RECT
  Dim lWidth      As Long
  Dim lHeight     As Long
  Dim hMemBmp     As Long
  Dim iDrawnCt    As Integer
  Dim bmpConsole  As BITMAP
  
  hConsoleDC = apiGetDC(poConsole.hWnd)
  hMemDC = apiCreateCompatibleDC(hConsoleDC)
  If hMemDC = 0 Then
    SetErr LOCAL_ERR_CTX, -1&, "CreateCompatibleDC has failed"
    GoTo SaveConsoleAsBitmap_Exit
  End If
  
  'Call apiGetClientRect(poConsole.hwnd, rcConsole)
  lWidth = poConsole.GetLongestLineWidth() 'rcConsole.Right - rcConsole.Left
  lHeight = (piEndLine - piStartLine + 1) * poConsole.CharHeight
  
  If lWidth = 0 Then
    SetErr LOCAL_ERR_CTX, -1&, "Only empty lines in console"
    GoTo SaveConsoleAsBitmap_Exit
  End If
  
  hMemBmp = apiCreateCompatibleBitmap(hConsoleDC, lWidth, lHeight)
  If hMemBmp = 0 Then
    SetErr LOCAL_ERR_CTX, -1&, "CreateCompatibleDC has failed"
    GoTo SaveConsoleAsBitmap_Exit
  End If
  
  Call apiSelectObject(hMemDC, hMemBmp)
  
  iDrawnCt = poConsole.PaintOnDC(hMemDC, piStartLine, piEndLine, lWidth, lHeight)
  
  apiGetObject hMemBmp, Len(bmpConsole), bmpConsole
  
  Dim bmfHeader As BITMAPFILEHEADER
  Dim bi        As BITMAPINFOHEADER
  Dim bytes_per_scanLine  As Integer
  Dim pad_per_scanLine    As Integer
  
  bi.biSize = Len(bi)
  bi.biWidth = bmpConsole.bmWidth
  bi.biHeight = bmpConsole.bmHeight
  bi.biPlanes = 1
  bi.biBitCount = 32
  bi.biCompression = BI_RGB
  
  'Computations source: http://www.vb-helper.com/howto_make_gray.html
  bytes_per_scanLine = ((((bi.biWidth * bi.biBitCount) + 31) \ 32) * 4)
  pad_per_scanLine = bytes_per_scanLine - (((bi.biWidth * bi.biBitCount) + 7) \ 8)
  bi.biSizeImage = bytes_per_scanLine * Abs(bi.biHeight)
  'bi.biSizeImage = 0
  bi.biXPelsPerMeter = 0
  bi.biYPelsPerMeter = 0
  bi.biClrUsed = 0
  bi.biClrImportant = 0
  
  Dim dwBmpSize As Long
  dwBmpSize = bi.biSizeImage 'BUG: don't do /32 but \32, this is not correct : ((bmpConsole.bmWidth * bi.biBitCount + 31) / 32) * 4 * bmpConsole.bmHeight
  
  Dim hDIB As Long
  hDIB = GlobalAlloc(GHND, dwBmpSize)
  
  Dim lpBitmap As Long
  Dim lAPIRet  As Long
  
  lpBitmap = GlobalLock(hDIB)
  
  'The return value of GetDIBits has to be checked, see
  'https://docs.microsoft.com/fr-fr/windows/win32/api/wingdi/nf-wingdi-getdibits?redirectedfrom=MSDN
  lAPIRet = GetDIBits(hMemDC, hMemBmp, 0, _
                 bmpConsole.bmHeight, _
                 lpBitmap, _
                 VarPtr(bi), DIB_RGB_COLORS)
  If lAPIRet = 0 Then
    SetErr LOCAL_ERR_CTX, -1&, LastApiError()
    GoTo SaveConsoleAsBitmap_Exit
  End If
  
  Dim dwSizeofDIB As Long
  dwSizeofDIB = dwBmpSize + Len(bmfHeader) + Len(bi)
  bmfHeader.bfOffBits = Len(bmfHeader) + Len(bi)
  bmfHeader.bfSize = dwSizeofDIB
  bmfHeader.bfType = &H4D42
  
  'write to file
  Dim hFile As Long
  Dim dwBytesWritten  As Long
  Dim lFileErr As Long
  Dim sFileErr As String
  
  'hFile = Win32OpenFileRaw(psFilename, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0&, sFileErr)
  Dim fh As Integer
  fh = FreeFile
  Open psFilename For Binary Access Write Lock Read Write As #fh
  Close fh
  
  Dim fOK As Boolean
  fOK = True
  hFile = Win32OpenFile(psFilename, False, False, lFileErr, sFileErr)
  If hFile <> INVALID_HANDLE_VALUE Then
    'The bitmap file header is aligned in memory on a word boundary and so it takes up 16 not 14 bytes
    'So we copy into a byte array to write the correct bytes in the file.
    'Source: https://www.tek-tips.com/viewthread.cfm?qid=1666866
    Dim abHeader(0 To Len(bmfHeader) - 1) As Byte
    CopyMemory abHeader(0), bmfHeader, 2
    CopyMemory abHeader(2), bmfHeader.bfSize, Len(bmfHeader) - 2
    Call Win32WriteBytes(hFile, abHeader(), dwBytesWritten, lFileErr, sFileErr)
    If lFileErr = 0 Then Call Win32WriteData(hFile, VarPtr(bi), Len(bi), dwBytesWritten, lFileErr, sFileErr)
    If lFileErr = 0 Then Call Win32WriteData(hFile, lpBitmap, dwBmpSize, dwBytesWritten, lFileErr, sFileErr)
    If lFileErr <> 0 Then
      SetErr LOCAL_ERR_CTX, lFileErr, sFileErr
      fOK = False
    End If
    Call Win32CloseFile(hFile, lFileErr, sFileErr)
  Else
    SetErr LOCAL_ERR_CTX, lFileErr, sFileErr
    fOK = False
  End If
  
  Call GlobalUnlock(hDIB)
  Call GlobalFree(hDIB)
  
  SaveConsoleAsBitmap = fOK
  
SaveConsoleAsBitmap_Exit:
  Call apiDeleteObject(hMemBmp)
  Call apiDeleteObject(hMemDC)
  Call apiReleaseDC(poConsole.hWnd, hConsoleDC)
  Exit Function
  
SaveConsoleAsBitmap_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume SaveConsoleAsBitmap_Exit
End Function

