Attribute VB_Name = "MSupport"
Option Compare Database
Option Explicit

'For the color picker dialog
Private Const LOGPIXELSX = 88 ' ditto
Private Const LOGPIXELSY = 90 ' ditto
Private Const TwipsPerInch = 1440
Private Declare PtrSafe Function GetDC Lib "user32" ( _
  ByVal hWnd As LongPtr) As LongPtr ' returns a HDC - Handle to a Device Context

Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
  ByVal hdc As LongPtr, ByVal nIndex As Long) As Long ' returns a C/C++ int

Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
  ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long ' also returns an int

Public ConsoulEventDispatcher As New CConsoulEventDispatcher
Public MessageManager As New CMessageManager

Public Const MSGTOPIC_LOCKUI              As String = "lockui"            'No params
Public Const MSGTOPIC_UNLOCKUI            As String = "unlockui"          'No params
Public Const MSGTOPIC_CANUNLOAD           As String = "canunload"         'param = message box title, but return value <> 0 if cannot unload
Public Const MSGTOPIC_UNLOADNOW           As String = "unloadnow"         'No params
Public Const MSGTOPIC_CHARSELECTED        As String = "charselected"      'param: char code as Long
Public Const MSGTOPIC_CANVASRESIZED       As String = "canvasresized"     'rowParams("WindowLeft","WindowTop","WindowWidth","WindowHeight")
Public Const MSGTOPIC_CHARMAPMOVED        As String = "charmapmoved"      'rowParams("WindowLeft","WindowTop","WindowWidth","WindowHeight")
Public Const MSGTOPIC_GETSELCOLOR         As String = "getselcolor"       'rowParams("color")=Null, returns selected color in rowParams("color")
Public Const MSGTOPIC_ADDNSELCOLOR        As String = "addnselcolor"      'rowParams("color")
Public Const MSGTOPIC_FONTNAMECHANGED     As String = "fontnamechanged"   'param: new font name
Public Const MSGTOPIC_FONTFAMLCHANGED     As String = "fontfamlchanged"   'No params
Public Const MSGTOPIC_FINDCHAR            As String = "findchar"          'param: char code

'Attributes bit values for CConsoleGrid, made public
Public Const ATTRIB_BOLDON      As Integer = 1
Public Const ATTRIB_ITALICON    As Integer = 2
Public Const ATTRIB_UNDLON      As Integer = 4
Public Const ATTRIB_STRIKEON    As Integer = 8
Public Const ATTRIB_INVERSEON   As Integer = 16
Public Const ATTRIB_BOLDOFF     As Integer = 32
Public Const ATTRIB_ITALICOFF   As Integer = 64
Public Const ATTRIB_UNDLOFF     As Integer = 128
Public Const ATTRIB_STRIKEOFF   As Integer = 256
Public Const ATTRIB_INVERSEOFF  As Integer = 512
Public Const ATTRIB_RESET       As Integer = 1024

'Special folders
Public Enum esfSpecialFolder
  CSIDL_DESKTOP = &H0
  CSIDL_PROGRAMS = &H2
  CSIDL_CONTROLS = &H3
  CSIDL_PRINTERS = &H4
  CSIDL_PERSONAL = &H5
  CSIDL_FAVORITES = &H6
  CSIDL_STARTUP = &H7
  CSIDL_RECENT = &H8
  CSIDL_SENDTO = &H9
  CSIDL_BITBUCKET = &HA
  CSIDL_STARTMENU = &HB
  CSIDL_DESKTOPDIRECTORY = &H10
  CSIDL_DRIVES = &H11
  CSIDL_NETWORK = &H12
  CSIDL_NETHOOD = &H13
  CSIDL_FONTS = &H14
  CSIDL_TEMPLATES = &H15
  CSIDL_COMMON_STARTMENU = &H16
  CSIDL_COMMON_PROGRAMS = &H17
  CSIDL_COMMON_STARTUP = &H18
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19
  CSIDL_APPDATA = &H1A
  CSIDL_PRINTHOOD = &H1B
End Enum

Private Const FORMAT_MESSAGE_FROM_SYSTEM& = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS& = &H200

#If Win64 Then
  Private Type SHITEMID
    cb As LongPtr
    abID As Byte
  End Type
  Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Any, ByVal pszPath As String) As LongPtr
  Private Declare PtrSafe Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As LongPtr, ByVal nFolder As Long, pidl As TITEMIDLIST) As LongPtr
  Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
  Private Declare PtrSafe Function PathCompactPath Lib "shlwapi.dll" Alias "PathCompactPathW" (ByVal hdc As LongPtr, ByVal pszPath As LongPtr, ByVal dx As Long) As Long
  Private Declare PtrSafe Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
  Private Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  Private Declare PtrSafe Function apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
  Private Declare PtrSafe Function apiMapWindowPoints Lib "user32" Alias "MapWindowPoints" (ByVal hwndFrom As LongPtr, ByVal hwndTo As LongPtr, lppt As RECT, ByVal cPoints As Long) As Long
  Private Declare PtrSafe Function apiCopyRect Lib "user32" Alias "CopyRect" (ByVal lpDestRect As LongPtr, ByVal lpSourceRect As LongPtr) As Long
  Private Declare PtrSafe Function apiRectangle Lib "gdi32" Alias "Rectangle" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
  Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (firstbyte As Byte) As Long

  Public Declare PtrSafe Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Public Declare PtrSafe Function apiGetStockObject Lib "gdi32" Alias "GetStockObject" (ByVal nIndex As Long) As LongPtr
  Public Declare PtrSafe Function apiCreatePen Lib "gdi32" Alias "CreatePen" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
  Public Declare PtrSafe Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As LongPtr) As Long

#Else
  Private Type SHITEMID
    cb As Long
    abID As Byte
  End Type
  Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
  Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As TITEMIDLIST) As Long
  Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
  Private Declare Function FormatMessage& Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long)
  Private Declare Function PathCompactPath Lib "shlwapi.dll" Alias "PathCompactPathW" (ByVal hdc As Long, ByVal pszPath As Long, ByVal dx As Long) As Long
  Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As RECT) As Long
  Private Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  Private Declare PtrSafe Function apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As Long, lpRect As RECT) As Long
  Private Declare PtrSafe Function apiMapWindowPoints Lib "user32" Alias "MapWindowPoints" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As RECT, ByVal cPoints As Long) As Long
  Private Declare PtrSafe Function apiCopyRect Lib "user32" Alias "CopyRect" (ByVal lpDestRect As Long, ByVal lpSourceRect As Long) As Long
  Private Declare PtrSafe Function apiRectangle Lib "gdi32" Alias "Rectangle" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
  Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (firstbyte As Byte) As Long
  
  Public Declare PtrSafe Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As Long, ByVal hObject As Long) As Long
  Public Declare PtrSafe Function apiGetStockObject Lib "gdi32" Alias "GetStockObject" (ByVal nIndex As Long) As Long
  Public Declare PtrSafe Function apiCreatePen Lib "gdi32" Alias "CreatePen" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
  Public Declare PtrSafe Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
  Public Declare PtrSafe Function apiMoveToEx Lib "gdi32" Alias "MoveToEx" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpPoint As LongPtr) As Long
  Public Declare PtrSafe Function apiLineTo Lib "gdi32" Alias "LineTo" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
#End If
' Stock Logical Objects
Public Const WHITE_BRUSH = 0
Public Const LTGRAY_BRUSH = 1
Public Const GRAY_BRUSH = 2
Public Const DKGRAY_BRUSH = 3
Public Const BLACK_BRUSH = 4
Public Const NULL_BRUSH = 5
Public Const HOLLOW_BRUSH = NULL_BRUSH
Public Const WHITE_PEN = 6
Public Const BLACK_PEN = 7
Public Const NULL_PEN = 8
Public Const OEM_FIXED_FONT = 10
Public Const ANSI_FIXED_FONT = 11
Public Const ANSI_VAR_FONT = 12
Public Const SYSTEM_FONT = 13
Public Const DEVICE_DEFAULT_FONT = 14
Public Const DEFAULT_PALETTE = 15
Public Const SYSTEM_FIXED_FONT = 16
Public Const STOCK_LAST = 16

'  Pen Styles
Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6
Public Const PS_USERSTYLE = 7
Public Const PS_ALTERNATE = 8
Public Const PS_STYLE_MASK = &HF

Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_ENDCAP_MASK = &HF00

Public Const PS_JOIN_ROUND = &H0
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_MASK = &HF000&

Public Const PS_COSMETIC = &H0
Public Const PS_GEOMETRIC = &H10000
Public Const PS_TYPE_MASK = &HF0000

Private Type TITEMIDLIST
  mkid As SHITEMID
End Type

Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11

'----------------- Consoul callbacks -----------
Public Function OnConsoulMouseButton(ByVal phWnd As Long, ByVal piEvtCode As Integer, ByVal pwParam As Integer, ByVal piZoneID As Integer, ByVal piRow As Integer, ByVal piCol As Integer, ByVal piPosX As Integer, ByVal piPosY As Integer) As Integer
  On Error Resume Next
  ConsoulEventDispatcher.BroadcastMouseEvent phWnd, piEvtCode, pwParam, piZoneID, piRow, piCol, piPosX, piPosY
End Function

Public Function OnConsoulVirtualLine(ByVal phWnd As Long, ByVal piLine As Long) As Integer
  On Error Resume Next
  ConsoulEventDispatcher.BroadcastVirtualLineEvent phWnd, piLine
End Function

Public Function OnConsoulWmPaint(ByVal phWnd As Long, ByVal pwCbkMode As Integer, ByVal phDC As Long, ByVal lprcLinePos As Long, ByVal lprcLineRect As Long, ByVal lprcPaint As Long) As Integer
  On Error Resume Next
  ConsoulEventDispatcher.BroadcastConsolePaint phWnd, pwCbkMode, phDC, lprcLinePos, lprcLineRect, lprcPaint
End Function

' ---------------- Global tools ----------------

'Originally from the book "http://vb.mvps.org/hardweb/hardbook.htm", but formula also
'explained on stackoverflow:
'https://stackoverflow.com/questions/22627708/generating-a-random-number-between-1-and-20
Public Function GetRandom(ByVal piLo As Long, ByVal piHi As Long) As Long
  GetRandom = Int(piLo + (Rnd * (piHi - piLo + 1)))
End Function

'Wait a certain amount of seconds, ...approximately
Public Sub ApproxWait(ByVal pfWaitSec As Double)
  Dim t As Single
  t = Timer
  While Timer - t < pfWaitSec
    DoEvents
  Wend
End Sub

' ---------------------- String/parsing utilites ----------------------------

'Get the part of a string before a chr$(0), which is the string terminator in C.
Public Function CtoVB(ByRef pszString As String) As String
  Dim i   As Long
  i = InStr(pszString, Chr$(0))
  If i Then
    CtoVB = left$(pszString, i - 1&)
  Else
    CtoVB = pszString
  End If
End Function

'Split a string into a new array.
'Returns the number of elements in the array.
Public Function SplitString(ByRef asRetItems() As String, _
  ByVal sToSplit As String, _
  Optional sSep As String = " ", _
  Optional lMaxItems As Long = 0&, _
  Optional eCompare As VbCompareMethod = vbBinaryCompare) _
  As Long

  Dim lPos        As Long
  Dim lDelimLen   As Long
  Dim lRetCount   As Long
  
  On Error Resume Next
  Erase asRetItems
  On Error GoTo SplitString_Err
  
  If Len(sToSplit) Then
    lDelimLen = Len(sSep)
    If lDelimLen Then
      lPos = InStr(1, sToSplit, sSep, eCompare)
      Do While lPos
        lRetCount = lRetCount + 1&
        ReDim Preserve asRetItems(1& To lRetCount)
        asRetItems(lRetCount) = left$(sToSplit, lPos - 1&)
        sToSplit = Mid$(sToSplit, lPos + lDelimLen)
        If lMaxItems Then
          If lRetCount = lMaxItems - 1& Then Exit Do
        End If
        lPos = InStr(1, sToSplit, sSep, eCompare)
      Loop
    End If
    lRetCount = lRetCount + 1&
    ReDim Preserve asRetItems(1& To lRetCount)
    asRetItems(lRetCount) = sToSplit
  End If
  SplitString = lRetCount
SplitString_Err:
End Function

Public Function StrBlock(ByVal psText As String, ByVal psPadChar As String, ByVal piMaxLen As Integer) As String
  Dim iLen      As Integer
  
  iLen = Len(psText)
  If iLen <= piMaxLen Then
    StrBlock = psText & String$(piMaxLen - iLen, psPadChar)
  Else
    If piMaxLen > 6 Then
      StrBlock = left$(psText, piMaxLen - 3) & "..."
    Else
      StrBlock = left$(psText, piMaxLen)
    End If
  End If
End Function

' ---------------------- Color picker dialog ----------------------------
'Modififed from From https://stackoverflow.com/questions/23042374/access-2010-vba-api-twips-pixel
Function TwipsToPixelsY(ByVal y As Long) As Integer
  Dim ScreenDC As LongPtr
  ScreenDC = GetDC(0)
  TwipsToPixelsY = y / TwipsPerInch * GetDeviceCaps(ScreenDC, LOGPIXELSY)
  ReleaseDC 0, ScreenDC
End Function

Function TwipsToPixelsX(ByVal x As Long) As Integer
  Dim ScreenDC As LongPtr
  ScreenDC = GetDC(0)
  TwipsToPixelsX = x / TwipsPerInch * GetDeviceCaps(ScreenDC, LOGPIXELSX)
  ReleaseDC 0, ScreenDC
End Function

Function PixelsToTwipsX(ByVal x As Integer) As Long
  Dim ScreenDC As LongPtr
  ScreenDC = GetDC(0)
  PixelsToTwipsX = CLng(CDbl(x) * TwipsPerInch / GetDeviceCaps(ScreenDC, LOGPIXELSX))
  ReleaseDC 0, ScreenDC
End Function

Function PixelsToTwipsY(ByVal y As Integer) As Long
  Dim ScreenDC As LongPtr
  ScreenDC = GetDC(0)
  PixelsToTwipsY = CLng(CDbl(y) * TwipsPerInch / GetDeviceCaps(ScreenDC, LOGPIXELSY))
  ReleaseDC 0, ScreenDC
End Function

Public Sub ShowUFError(ByVal psNiceErrorMessage As String, ByVal psTechErrorDesc As String)
  Dim sMsg        As String
  
  sMsg = psNiceErrorMessage & vbCrLf & vbCrLf & "Technical details:" & vbCrLf & vbCrLf & psTechErrorDesc
  MsgBox sMsg, vbCritical, "Application error"
End Sub

Public Function IsIntegerEditKeyCodeAllowed(ByVal piKeyCode As Integer) As Boolean
  If ((piKeyCode >= vbKey0) And (piKeyCode <= vbKey9)) Or _
     ((piKeyCode >= vbKeyNumpad0) And (piKeyCode <= vbKeyNumpad9)) Or _
     (piKeyCode = vbKeyDelete) Or (piKeyCode = vbKeyLeft) Or (piKeyCode = vbKeyRight) Or _
     (piKeyCode = vbKeyHome) Or (piKeyCode = vbKeyEnd) Or _
     (piKeyCode = vbKeyTab) Or (piKeyCode = vbKeyBack) Or (piKeyCode = vbKeyReturn) _
      Then
    IsIntegerEditKeyCodeAllowed = True
  End If
End Function

'------------------------- Help from The Colibri project ------------------------
Function CountStringParts(ByRef sSource As String, ByRef sSeps As String) As Long
  Dim n           As Long
  Dim p           As Long
  Dim sBreak      As String
  Dim iBreakLen   As Long
  
  sBreak = sSeps
  iBreakLen = Len(sBreak)

  'Remove any leading / trailing sBreak
  While left$(sSource, iBreakLen) = sBreak
      sSource = Right$(sSource, Len(sSource) - iBreakLen)
  Wend
  While Right$(sSource, iBreakLen) = sBreak
    sSource = left$(sSource, Len(sSource) - iBreakLen)
  Wend

  'Count
  p = InStr(sSource, sBreak)
  While p&
    n = n + 1
    p = InStr(p + iBreakLen, sSource, sBreak)
  Wend
  If n = 0 Then
    If sSource <> "" Then
      CountStringParts = 1&
    End If
  Else
    CountStringParts = n + 1&
  End If
End Function

Function GetStringPart(ByVal plIndex As Long, ByRef psSep As String, ByRef psObjects As String, Optional ByVal pfTrimSeps As Boolean = True) As String
  Dim i    As Long
  Dim p    As Long
  Dim pp   As Long
  Dim fBad As Boolean
  
  Dim strBreak    As String
  Dim strObjects  As String
  Dim lBreakLen   As Long
  
  strObjects = psObjects
  strBreak = psSep
  lBreakLen = Len(strBreak)
  
  If pfTrimSeps Then
    While left$(strObjects, lBreakLen) = strBreak
      strObjects = Right$(strObjects, Len(strObjects) - lBreakLen)
    Wend
    While Right$(strObjects, lBreakLen) = strBreak
      strObjects = left$(strObjects, Len(strObjects) - lBreakLen)
    Wend
  End If

  If plIndex > 1& Then
    For i = 1& To plIndex - 1&
      p = InStr(p + 1&, strObjects, strBreak)
      If p = 0& Then fBad = True: Exit For
    Next i
    If Not fBad Then
      pp = InStr(p + lBreakLen, strObjects, strBreak)
      If pp Then
        GetStringPart = Mid$(strObjects, p + lBreakLen, pp - p - lBreakLen)
      Else
        GetStringPart = Right$(strObjects, Len(strObjects) - p - lBreakLen + 1&)
      End If
    End If
  Else
    p = InStr(strObjects, strBreak)
    If p Then
      GetStringPart = left$(strObjects, p - 1&)
    Else
      GetStringPart = strObjects
    End If
  End If
End Function

'-------- color to rgb parts

'From http://www.excely.com/excel-vba/bit-shifting-function.shtml
Public Function shr(ByVal Value As Long, ByVal Shift As Byte) As Long
    Dim i As Byte
    shr = Value
    If Shift > 0 Then
        shr = Int(shr / (2 ^ Shift))
    End If
End Function

Public Function GetRed(ByVal lColor As Long) As Integer
  GetRed = shr(lColor, 16) And &HFF
End Function

Public Function GetBlue(ByVal lColor As Long) As Integer
  GetBlue = lColor And &HFF
End Function

Public Function GetGreen(ByVal lColor As Long) As Integer
  GetGreen = shr(lColor, 8) And &HFF
End Function

Public Function HexWord(ByVal pwWord As Integer) As String
  Dim sWord As String
  sWord = Hex$(pwWord)
  If Len(sWord) < 2 Then
    sWord = "0" & sWord
  End If
  HexWord = sWord
End Function

Public Function HexInt(ByVal piInt As Integer) As String
  Dim sInt As String
  sInt = Hex$(piInt)
  If Len(sInt) < 4 Then
    sInt = String$(4 - Len(sInt), "0") & sInt
  End If
  HexInt = sInt
End Function

Public Function ColorToHex(ByVal plColor As Long) As String
  Dim R As Integer
  Dim G As Integer
  Dim B As Integer
  R = GetRed(plColor)
  G = GetGreen(plColor)
  B = GetBlue(plColor)
  ColorToHex = HexWord(R) & HexWord(G) & HexWord(B)
End Function

Public Function ColorFromHex(ByVal psHexColor As String) As Long
  If Len(psHexColor) = 0 Then Exit Function
  psHexColor = UCase$(psHexColor)
  If left$(psHexColor, 2) = "&H" Then
    psHexColor = Right$(psHexColor, Len(psHexColor) - 2)
  End If
  If Len(psHexColor) > 6 Then Exit Function
  If Len(psHexColor) < 6 Then
    psHexColor = String$(6 - Len(psHexColor), "0") & psHexColor
  End If
  
  Dim R As Integer
  Dim G As Integer
  Dim B As Integer
  R = CInt(Val(left$(psHexColor, 2)))
  G = CInt(Val(Mid$(psHexColor, 3, 2)))
  B = CInt(Val(Right$(psHexColor, 2)))
  ColorFromHex = RGB(R, G, B)
End Function

' 3 different ways to display a color value (for labels that show the color code)
Public Function GetColorDispString(ByVal plColor As Long, ByVal piMode As Integer) As String
  Select Case piMode
  Case 0
    GetColorDispString = "RGB: " & GetRed(plColor) & "," & GetGreen(plColor) & "," & GetBlue(plColor)
  Case 1
    GetColorDispString = "HEX: " & ColorToHex(plColor)
  Case 2
    GetColorDispString = "DEC: " & plColor
  End Select
End Function

'--------------- from Colibri's MTools
'Note 9999 is the max we can get out!
Public Function IntChooseBox(ByVal sText As String, ByVal sTitle As String, ByVal sDefault As String, ByVal iMax As Integer, Optional ByVal iMin As Integer = 1) As Integer
  Dim sInput    As String
  Dim fValid    As Boolean
  Dim iRet      As Integer
  
  If iMin < 1 Then iMin = 1
  If iMax < iMin Then
    iMax = iMin 'sounds dummy...
  End If
  
  Do
    sInput = InputBox$(sText, sTitle, sDefault)
    If Len(sInput) Then
      If IsNumeric(sInput) Then
        If Len(sInput) <= 4 Then
          iRet = CInt(Val(sInput))
          If (iRet >= iMin) And (iRet <= iMax) Then
            fValid = True
          Else
            MsgBox "Please enter a number between " & iMin & " and " & iMax, vbCritical
          End If
        Else
          MsgBox "The text you typed is too long", vbCritical
        End If
      Else
        MsgBox "The text you typed is not a valid number", vbCritical
      End If
    End If
  Loop Until (sInput = "") Or fValid
  
  If fValid Then
    IntChooseBox = iRet
  End If
End Function

'max 9 milions (9 999 999)
Public Function LongChooseBox( _
    ByVal sText As String, _
    ByVal sTitle As String, _
    ByVal sDefault As String, _
    ByVal lMax As Long, _
    Optional ByVal lMin As Long = 1&, _
    Optional ByVal piMaxInputLen As Integer = 7 _
  ) As Long
  Dim sInput    As String
  Dim fValid    As Boolean
  Dim lRet      As Long
  
  If lMin < 1& Then lMin = 1&
  If lMax < lMin Then
    lMax = lMin 'sounds dummy...
  End If
  
  Do
    sInput = InputBox$(sText, sTitle, sDefault)
    If Len(sInput) Then
      If IsNumeric(sInput) Then
        If Len(sInput) <= piMaxInputLen Then
          On Error Resume Next
          lRet = CLng(Val(sInput))
          If Err.Number = 0 Then
            If (lRet >= lMin) And (lRet <= lMax) Then
              fValid = True
            Else
              MsgBox "Please enter a number between " & lMin & " and " & lMax, vbCritical
            End If
          Else
            MsgBox "Invalid color value", vbCritical
          End If
          On Error GoTo 0
        Else
          MsgBox "The text you typed is too long", vbCritical
        End If
      Else
        MsgBox "The text you typed is not a valid number", vbCritical
      End If
    End If
  Loop Until (sInput = "") Or fValid
  
  If fValid Then
    LongChooseBox = lRet
  End If
End Function

'Returns true is a valid color value is entered
Public Function InputColor( _
    ByRef plRetColor As Long, _
    ByVal sText As String, _
    ByVal sTitle As String, _
    ByVal sDefault As String _
  ) As Boolean
  Dim sInput    As String
  Dim fValid    As Boolean
  Dim lRet      As Long
  Dim sHex      As String
  
  Do
    fValid = False
    sInput = Trim$(InputBox$(sText, sTitle, sDefault))
    If Len(sInput) > 0 Then
      If IsNumeric(sInput) Then
        If Len(sInput) <= 10 Then
          On Error Resume Next
          If UCase$(left$(sInput, 2)) = "&H" Then
            sHex = Right$(sInput, Len(sInput) - 2)
            If Len(sHex) <> 6 Then
              MsgBox "Please specify 6 digits after '&H'", vbCritical
              GoTo ask_again
            End If
            lRet = Val(sInput)
          Else
            lRet = CLng(Val(sInput))
          End If
          lRet = RGB(GetRed(lRet), GetGreen(lRet), GetBlue(lRet))
          Debug.Print "InputColor --> " & lRet & " => &H" & Hex$(lRet)
          If Err.Number = 0 Then
            fValid = True
            plRetColor = lRet
          Else
            MsgBox "Invalid color value", vbCritical
          End If
          On Error GoTo 0
        Else
          MsgBox "The text you typed is too long", vbCritical
        End If
      Else
        MsgBox "The text you typed is not a valid number", vbCritical
      End If
    Else
      Exit Do
    End If
ask_again:
  Loop Until fValid
  
  InputColor = fValid
End Function

'
' File / Dirs tools
'

' Test the existence of a file
Public Function ExistFile(psSpec As String) As Boolean
  On Error Resume Next
  Call FileLen(psSpec)
  ExistFile = (Err.Number = 0&)
End Function

Public Function ExistDir(psSpec As String) As Boolean
  On Error Resume Next
  Dim lAttr As Long
  lAttr = GetAttr(psSpec)
  If (Err.Number = 0) Then
    ExistDir = CBool(lAttr And vbDirectory)
  End If
End Function

Public Function GetSpecialFolder(ByVal phWndRef As LongPtr, eSpecialFolderID As esfSpecialFolder) As String
  #If Win64 Then
  Dim lRet          As LongPtr
  Dim lTrans        As LongPtr
  #Else
  Dim lRet          As Long
  Dim lTrans        As Long
  #End If
  Dim spath         As String
  Dim TITEMIDLIST   As TITEMIDLIST
  
  Const klMaxLength As Long = 1024&

  lRet = SHGetSpecialFolderLocation(phWndRef, eSpecialFolderID, TITEMIDLIST)
  If lRet = 0 Then
    spath = String$(klMaxLength, Chr$(0))
    lTrans = TITEMIDLIST.mkid.cb
    lRet = SHGetPathFromIDList(ByVal lTrans, spath)
    If lRet <> 0 Then
      GetSpecialFolder = left$(spath, InStr(spath, Chr$(0)) - 1) & PATH_SEP
    End If
  End If
End Function

'Original source: http://www.vb-helper.com/howto_invert_color.html
Public Function InverseColor(ByVal plColor As Long) As Long
  Dim iR      As Integer
  Dim iG      As Integer
  Dim iB      As Integer
  Dim lColor  As Long
  iR = 255 - GetRed(plColor)
  iG = 255 - GetGreen(plColor)
  iB = 255 - GetBlue(plColor)
  lColor = RGB(iR, iG, iB)
  InverseColor = lColor
End Function

Public Sub FilePutUnicodeString(ByVal fh As Integer, ByVal psText As String, Optional ByVal pfOutputLength As Boolean = True)
  Dim j       As Integer
  Dim iLen    As Integer
  Dim iAscW   As Integer
  
  iLen = CInt(Len(psText))
  If pfOutputLength Then Put #fh, , iLen
  If iLen > 0 Then
    For j = 1 To iLen
      iAscW = AscW(Mid$(psText, j, 1))
      Put #fh, , iAscW
    Next j
  End If
End Sub

Public Function FileGetUnicodeString(ByVal fh As Integer, Optional ByVal piMaxChars As Integer = 0) As String
  Dim j       As Integer
  Dim iLen    As Integer
  Dim sText   As String
  Dim iAscW   As Integer
  
  Get #fh, , iLen
  
  If (piMaxChars = 0) Or (piMaxChars > iLen) Then
    piMaxChars = iLen
  End If
  sText = Space$(piMaxChars)
  
  For j = 1 To iLen
    Get #fh, , iAscW
    'sText = sText & ChrW$(iAscW)
    If j <= piMaxChars Then
      Mid$(sText, j, 1) = ChrW$(iAscW)
    End If
  Next j
  FileGetUnicodeString = sText
End Function

Public Function GetFileText(sFilename As String) As String
  Dim nFile As Integer, sText As String
  nFile = FreeFile
  If Not ExistFile(sFilename) Then Exit Function
  ' Let others read but not write
  Open sFilename For Binary Access Read Lock Write As nFile
  sText = String$(LOF(nFile), 0)
  Get nFile, 1, sText
  Close nFile
  GetFileText = sText
End Function

Function ApiError(ByVal e As Long) As String
  Dim s As String, c As Long
  s = String(256, 0)
  c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                    FORMAT_MESSAGE_IGNORE_INSERTS, _
                    ByVal 0&, e, 0&, s, Len(s), ByVal 0&)
  If c Then ApiError = left$(s, c)
End Function

Function LastApiError() As String
  LastApiError = ApiError(Err.LastDllError)
End Function

#If Win64 Then
Public Function CompactPath(ByVal phWnd As LongPtr, ByVal psPath As String, ByVal plPixelLen As Long) As String
  Dim hdc As LongPtr
#Else
Public Function CompactPath(ByVal phWnd As Long, ByVal psPath As String, ByVal plPixelLen As Long) As String
  Dim hdc As Long
#End If
  On Error Resume Next
  Dim sCompact    As String
  
  If Len(psPath) = 0 Then Exit Function
  If phWnd = 0 Then phWnd = apiGetDesktopWindow()
  psPath = psPath & Chr$(0)
  sCompact = psPath
  hdc = GetDC(phWnd)
  Call PathCompactPath(hdc, StrPtr(psPath), plPixelLen)
  Call ReleaseDC(phWnd, hdc)
  CompactPath = CtoVB(psPath)
End Function

#If Win64 Then
Public Function GetWindowRect(ByVal plHWnd As LongPtr, ByRef pRetRECT As RECT) As Long
#Else
Public Function GetWindowRect(ByVal plHWnd As Long, ByRef pRetRECT As RECT) As Long
#End If
  GetWindowRect = apiGetWindowRect(plHWnd, pRetRECT)
End Function

#If Win64 Then
Public Function GetClientRect(ByVal plHWnd As LongPtr, ByRef pRetRECT As RECT) As Long
#Else
Public Function GetClientRect(ByVal plHWnd As Long, ByRef pRetRECT As RECT) As Long
#End If
  GetClientRect = apiGetClientRect(plHWnd, pRetRECT)
End Function

#If Win64 Then
Public Function MoveWindow(plHWnd As LongPtr, ByVal px As Long, ByVal py As Long, ByVal pnWidth As Long, ByVal pnHeight As Long, ByVal pbRepaint As Long)
#Else
Public Function MoveWindow(plHWnd As Long, ByVal px As Long, ByVal py As Long, ByVal pnWidth As Long, ByVal pnHeight As Long, ByVal pbRepaint As Long)
#End If
  MoveWindow = apiMoveWindow(plHWnd, px, py, pnWidth, pnHeight, pbRepaint)
End Function

#If Win64 Then
Public Function MapWindowPoints(ByVal phWndFrom As LongPtr, ByVal phWndTo As LongPtr, ByRef pRetRECT As RECT)
  MapWindowPoints = apiMapWindowPoints(phWndFrom, phWndTo, pRetRECT, 2&)
End Function
#Else
Public Function MapWindowPoints(ByVal phWndFrom As Long, ByVal phWndTo As LongPtr, ByRef pRetRECT As RECT)
  MapWindowPoints = apiMapWindowPoints(phWndFrom, phWndTo, pRetRECT, 2&)
End Function
#End If

#If Win64 Then
Public Function CopyRectByAddress(ByVal lpDestRect As LongPtr, ByVal lpSrcRect As LongPtr) As Long
#Else
Public Function CopyRectByAddress(ByVal lpDestRect As Long, ByVal lpSrcRect As Long) As Long
#End If
  CopyRectByAddress = apiCopyRect(lpDestRect, lpSrcRect)
End Function

#If Win64 Then
Public Function WinRectangle(ByVal hdc As LongPtr, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
#Else
Public Function WinRectangle(ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
#End If
  WinRectangle = apiRectangle(hdc, X1, y1, x2, y2)
End Function

Public Function LPad(ByVal s As String, ByVal PadChar As String, ByVal iLen As Integer) As String
  If iLen Then
    If Len(s) < iLen Then
      LPad = String$(iLen - Len(s), PadChar) & s
    Else
      LPad = left$(s, iLen)
    End If
  End If
End Function

Public Function ColorHex(ByVal plColor As Long) As String
  Dim sColor    As String
  sColor = Hex$(plColor)
  sColor = LPad(sColor, "0", 6)
  ColorHex = sColor
End Function

'---------- Consoul Show Grid Feature (to call in paint callbacks) ----------
Public Function OnPaintConsoleGrid(ByRef poConsole As CConsoul, ByVal phWnd As Long, ByVal phDC As LongPtr, ByVal lprcLinePos As Long, ByVal lprcLineRect As Long, ByVal lprcPaint As Long) As Long
  Dim rcPaint     As RECT
  Dim rcLinePos   As RECT
  Dim rcLineRect  As RECT
  Dim hOldBrush   As LongPtr
  Dim hBrush      As LongPtr
  Dim hDotPen     As LongPtr
  Dim hOldPen     As LongPtr
  
  On Error Resume Next
  
  CopyRectByAddress VarPtr(rcPaint), lprcPaint
  CopyRectByAddress VarPtr(rcLinePos), lprcLinePos
  CopyRectByAddress VarPtr(rcLineRect), lprcLineRect
'  Debug.Print
'  Debug.Print "WM_PAINT: hwnd=" & phWnd & ", pwCbkMode=" & pwCbkMode & ", phDC=" & phDC & "lprcLinePos=" & lprcLinePos & ", lprcLineRect=" & lprcLineRect & ", lprcPaint=" & lprcPaint
'  Debug.Print "-------------------------------------------------------------"
'  Debug.Print "rcPaint[";
'  With rcPaint
'    Debug.Print .Left & ", " & .Top & ", " & .Right & ", " & .Bottom & "]"
'  End With
'  Debug.Print "rcLinePos[";
'  With rcLinePos
'    Debug.Print .Left & ", " & .Top & ", " & .Right & ", " & .Bottom & "]"
'  End With
'  Debug.Print "rcLineRect[";
'  With rcLineRect
'    Debug.Print .Left & ", " & .Top & ", " & .Right & ", " & .Bottom & "]"
'  End With
  
  hBrush = apiGetStockObject(NULL_BRUSH)
  hOldBrush = apiSelectObject(phDC, hBrush)
  hDotPen = apiCreatePen(PS_DOT, 1, RGB(200, 200, 200))
  
  hOldPen = apiSelectObject(phDC, hDotPen)
  
  Dim iCol As Integer
  Dim x    As Integer
  Dim iCharHeight As Integer
  Dim iCharWidth  As Integer
  
  iCharHeight = poConsole.LineHeight 'poConsole.CharHeight
  iCharWidth = poConsole.CharWidth
  
  'Debug.Print "Drawing cols[" & rcLinePos.Left & " To " & rcLinePos.Right & "]"
  'Debug.Print "X Positions : {";
  For iCol = 0 To rcLinePos.Right - rcLinePos.left
    x = rcLineRect.left + (iCol * iCharWidth)
  '  Debug.Print x;
    Call apiMoveToEx(phDC, x, rcPaint.Top, 0)
    Call apiLineTo(phDC, x, rcPaint.Bottom)
  Next iCol
  'Debug.Print "}"
  
  Dim iRow As Integer
  Dim y    As Integer
  Dim iCount As Integer
  If iCharHeight > 0 Then
    iCount = (rcLineRect.Bottom - rcLineRect.Top) \ iCharHeight
  End If
  For iRow = 0 To iCount
    y = rcLineRect.Top + (iRow * iCharHeight)
    Call apiMoveToEx(phDC, rcPaint.left, y, 0)
    Call apiLineTo(phDC, rcPaint.Right, y)
  Next iRow
  
  Call apiSelectObject(phDC, hOldBrush)
  Call apiSelectObject(phDC, hOldPen)
  Call apiDeleteObject(hDotPen)
  
  OnPaintConsoleGrid = 0
End Function


'
' GUID
'
' Create a new GUID
' The string format compresses the GUID into 24 characters, well
' within the 31 character limit of an Ole Structured Storage Name
Public Function CreateGUID() As String
  Dim guidbuffer(15) As Byte
  Dim guidstring As String
  Dim prefixstring As String
  Dim prefixbyte As Integer
  Dim tchar As Byte
  Dim x As Integer
  Dim group As Integer
  Dim bytecounter As Integer
  Dim newhex As String
  Const MAP_STRING = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_-"
  
  Call CoCreateGuid(guidbuffer(0))
  
  For group = 0 To 3
    prefixbyte = 0
    For x = 0 To 3
      tchar = guidbuffer(bytecounter)
      bytecounter = bytecounter + 1
      prefixbyte = prefixbyte * 4
      ' Take low 6 bits
      guidstring = guidstring & Mid$(MAP_STRING, (tchar And &H3F) + 1, 1)
      prefixbyte = prefixbyte + ((tchar \ &H40) And &H3)
    Next x
    newhex = Hex$(prefixbyte)
    If Len(newhex) = 1 Then newhex = "0" & newhex
    prefixstring = prefixstring & newhex
  Next group
  CreateGUID = prefixstring & guidstring
End Function

Public Function NewSelectFileFilterList() As CList
  Dim lstFilters As CList
  Set lstFilters = New CList
  lstFilters.ArrayDefine Array("name", "extensions"), Array(vbString, vbString)
  Set NewSelectFileFilterList = lstFilters
End Function

Public Function SelectLoadFile( _
    ByVal psIniKeyInitialDir As String, _
    ByVal psDialogTitle As String, _
    ByRef plstFilters As CList, _
    ByRef psRetFileName As String _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "SelectLoadFile"
  On Error GoTo SelectLoadFile_Err
  
  Dim sInitialDir As String
  Dim iChoice     As Integer
  Dim i           As Integer
  
  psRetFileName = ""
  AppIniFile.GetOption psIniKeyInitialDir, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath(GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), APP_NAME)
  End If
  If Not ExistDir(sInitialDir) Then
    If Not CreatePath(sInitialDir) Then
      sInitialDir = GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL)
    End If
  End If
  
  With Application.FileDialog(msoFileDialogFilePicker)
    .Title = psDialogTitle
    .InitialFileName = NormalizePath(sInitialDir)
    .Filters.Clear
    If Not plstFilters Is Nothing Then
      If plstFilters.Count > 0 Then
        For i = 1 To plstFilters.Count
          .Filters.Add plstFilters("name", i), plstFilters("extensions", i)
        Next i
        .FilterIndex = 1
      End If
    End If
    iChoice = .Show()
    If iChoice <> 0 Then
      psRetFileName = .SelectedItems(1)
      SelectLoadFile = True
    End If
  End With

SelectLoadFile_Exit:
  Exit Function
SelectLoadFile_Err:
  ShowUFError "Filed to select file for loading", Err.Description
  Resume SelectLoadFile_Exit
End Function

Public Function CreatePath(ByVal psPathToMake As String) As Boolean
  Dim sCurPathSegment As String
  Dim iOffset         As Integer
  Dim iAnchor         As Integer
  Dim sOldPath        As String

  On Error Resume Next

  'Add trailing backslash
  If Right$(psPathToMake, 1) <> PATH_SEP Then psPathToMake = psPathToMake & PATH_SEP
  sOldPath = CurDir$
  iAnchor = 0

  'Loop and make each subdir of the path separately.
  iOffset = InStr(iAnchor + 1, psPathToMake, PATH_SEP)
  iAnchor = iOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
  Do
    iOffset = InStr(iAnchor + 1, psPathToMake, PATH_SEP)
    iAnchor = iOffset
    If iAnchor > 0 Then
      sCurPathSegment = left$(psPathToMake, iOffset - 1)
      ' Determine if this directory already exists
      On Error Resume Next
      ChDir sCurPathSegment
      If Err.Number <> 0 Then
        ' We must create this directory
        On Error GoTo CreatePath_Err
        MkDir sCurPathSegment
      End If
    End If
  Loop Until iAnchor = 0

  CreatePath = True
CreatePath_Exit:
  ChDir sOldPath
  Exit Function

CreatePath_Err:
  'ShowError "CreatePath", Err.Number, Err.Description
  Resume CreatePath_Exit
End Function

