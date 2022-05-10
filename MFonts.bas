Attribute VB_Name = "MFonts"
Option Compare Database
Option Explicit

'Win32 API

' Font Families
Public Const FF_DONTCARE = 0    '  Don't care or don't know.
Public Const FF_ROMAN = 16      '  Variable stroke width, serifed.
' Times Roman, Century Schoolbook, etc.
Public Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.
' Helvetica, Swiss, etc.
Public Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.
' Pica, Elite, Courier, etc.
Public Const FF_SCRIPT = 64     '  Cursive, etc.
Public Const FF_DECORATIVE = 80 '  Old English, etc.

Private Const LF_FACESIZE = 32
 
 'types expected by the Windows callback
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Private Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

Declare PtrSafe Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As LongPtr, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As LongPtr, ByVal lParam As LongPtr) As Long
Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long

Private mfFontsLoaded       As Boolean
Private mlFontCount         As Long
Private masFontName()       As String
Private miFontFamilyFilter  As Integer
Private msSelectedFontName  As String

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

Public Function FontLastErr() As Long
  FontLastErr = mlErrNo
End Function

Public Function FontLastErrDesc() As String
  FontLastErrDesc = msErrDesc
End Function

Public Function FontLastErrCtx() As String
  FontLastErrCtx = msErrCtx
End Function

Public Function LoadFontNames(Optional ByVal piFontFamilyFilter As Integer = 0, Optional ByVal hWnd As LongPtr = 0&) As Boolean
  On Error GoTo LoadFontNames_Err
  Dim hdc As LongPtr
  
  ClearErr
  mfFontsLoaded = False
  If mlFontCount Then Erase masFontName
  mlFontCount = 0
  
  If hWnd = 0& Then hWnd = GetFocus()
  hdc = GetDC(hWnd)
  miFontFamilyFilter = piFontFamilyFilter
   'this line requests Windows to call the 'EnumFontFamProc' function for each installed font
  EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, ByVal 0&
  If mlFontCount > 0 Then
    QuickSort masFontName, 1&, mlFontCount
  End If
  mfFontsLoaded = True
  LoadFontNames = True

LoadFontNames_Exit:
  On Error Resume Next
  ReleaseDC hWnd, hdc
  Exit Function

LoadFontNames_Err:
  SetErr "LoadFontNames", Err.Number, Err.Description
  Resume LoadFontNames_Exit
End Function

Public Function GetFontCount() As Long
  If Not mfFontsLoaded Then
    If Not LoadFontNames() Then Exit Function
  End If
  GetFontCount = mlFontCount
End Function

Public Function GetFontName(ByVal plIndex As Long) As String
  If Not mfFontsLoaded Then
    If Not LoadFontNames() Then Exit Function
  End If
  GetFontName = masFontName(plIndex)
End Function

Public Function FontExists(ByVal psFontName As String) As Boolean
  Dim fOK       As Boolean
  Dim i         As Long
  
  If Not mfFontsLoaded Then
    fOK = LoadFontNames()
    If Not fOK Then Exit Function
  End If
  
  For i = 1& To mlFontCount
    If StrComp(psFontName, masFontName(i), vbTextCompare) = 0 Then
      FontExists = True
      Exit Function
    End If
  Next i
End Function

Public Function FontGetFamilyFontFilter() As Integer
  FontGetFamilyFontFilter = miFontFamilyFilter
End Function

Public Function GetFontsComboSource(ByVal piFontFamilyFilter As Integer) As String
  Dim i               As Long
  Dim sSource         As String
  Dim fOK             As Boolean
  
  If (Not mfFontsLoaded) Or (piFontFamilyFilter <> miFontFamilyFilter) Then
    miFontFamilyFilter = piFontFamilyFilter
    fOK = LoadFontNames(miFontFamilyFilter)
    If Not fOK Then Exit Function
  End If
  
  For i = 1& To mlFontCount
    If i > 1& Then
      sSource = sSource & ";"
    End If
    sSource = sSource & masFontName(i)
  Next i
  
  GetFontsComboSource = sSource
End Function

Public Sub FontSetSelectedFont(ByVal psFontName As String)
  msSelectedFontName = psFontName
End Sub

Public Function FontGetSelectedFont() As String
  FontGetSelectedFont = msSelectedFontName
End Function

Private Function EnumFontFamProc(lpNLF As LOGFONT, _
                                 lpNTM As NEWTEXTMETRIC, _
                                 ByVal FontType As Long, _
                                 lParam As Long) As Long
  On Error GoTo errorcode
  Dim sFaceName As String
  Dim fFiltered As Boolean
  
  sFaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
  
  If left$(sFaceName, 1) = "@" Then  'Don't know what this means in a font name, but filter
    fFiltered = True
  End If
  If miFontFamilyFilter <> 0 Then
    'Bits 4 to 7 are Family
    If (lpNLF.lfPitchAndFamily And &H78) <> miFontFamilyFilter Then
      fFiltered = True
    End If
  End If
  
  If Not fFiltered Then
    mlFontCount = mlFontCount + 1&
    ReDim Preserve masFontName(1& To mlFontCount) As String
    masFontName(mlFontCount) = left$(sFaceName, InStr(sFaceName, vbNullChar) - 1)
  End If
  
  EnumFontFamProc = 1
  Exit Function
errorcode:
  EnumFontFamProc = 1
End Function

Private Sub QuickSort(ByRef pasArray() As String, ByVal iLBound As Long, ByVal iUBound As Long)
  Dim sPivot      As String
  Dim sTemp       As String
  Dim iLBoundTemp As Long
  Dim iUBoundTemp As Long
  iLBoundTemp = iLBound
  iUBoundTemp = iUBound
  sPivot = pasArray((iLBound + iUBound) \ 2)
  While (iLBoundTemp <= iUBoundTemp)
    While (pasArray(iLBoundTemp) < sPivot And iLBoundTemp < iUBound)
      iLBoundTemp = iLBoundTemp + 1
    Wend
    While (sPivot < pasArray(iUBoundTemp) And iUBoundTemp > iLBound)
      iUBoundTemp = iUBoundTemp - 1
    Wend
    If iLBoundTemp < iUBoundTemp Then
      sTemp = pasArray(iLBoundTemp)
      pasArray(iLBoundTemp) = pasArray(iUBoundTemp)
      pasArray(iUBoundTemp) = sTemp
    End If
    If iLBoundTemp <= iUBoundTemp Then
      iLBoundTemp = iLBoundTemp + 1&
      iUBoundTemp = iUBoundTemp - 1&
    End If
  Wend
  'the function calls itself until everything is in good order
  If (iLBound < iUBoundTemp) Then QuickSort pasArray, iLBound, iUBoundTemp
  If (iLBoundTemp < iUBound) Then QuickSort pasArray, iLBoundTemp, iUBound
End Sub


