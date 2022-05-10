Attribute VB_Name = "MConsoul"
Option Compare Database
Option Explicit

'Just consts around QBColor() numerical indices
Public Const QBCOLOR_BLACK        As Integer = 0
Public Const QBCOLOR_BLUE         As Integer = 1
Public Const QBCOLOR_GREEN        As Integer = 2
Public Const QBCOLOR_CYAN         As Integer = 3
Public Const QBCOLOR_RED          As Integer = 4
Public Const QBCOLOR_MAGENTA      As Integer = 5
Public Const QBCOLOR_YELLOW       As Integer = 6
Public Const QBCOLOR_WHITE        As Integer = 7
Public Const QBCOLOR_GRAY         As Integer = 8
Public Const QBCOLOR_LIGHTBLUE    As Integer = 9
Public Const QBCOLOR_LIGHTGREEN   As Integer = 10
Public Const QBCOLOR_LIGHTCYAN    As Integer = 11
Public Const QBCOLOR_LIGHTRED     As Integer = 12
Public Const QBCOLOR_LIGHTMAGENTA As Integer = 13
Public Const QBCOLOR_LIGHTYELLOW  As Integer = 14
Public Const QBCOLOR_BRIGHTWHITE  As Integer = 15

'Consoul Window specific message consts
Public Const WM_USER = &H400
Public Const WM_USER_ROWCOLCHANGE = WM_USER + 500
Public Const WM_USER_ZONEENTER = WM_USER + 501
Public Const WM_USER_ZONELEAVE = WM_USER + 502
'Mouse button (and shift/ctrl) states for mouse events
Public Const MK_CONTROL   As Integer = &H8    'The CTRL key is down
Public Const MK_LBUTTON   As Integer = &H1    'The left mouse button is down
Public Const MK_MBUTTON   As Integer = &H10   'The middle mouse button is down
Public Const MK_RBUTTON   As Integer = &H2    'The right mouse button is down
Public Const MK_SHIFT     As Integer = &H4    'The SHIFT key is down
Public Const MK_XBUTTON1  As Integer = &H20   'The first X button is down
Public Const MK_XBUTTON2  As Integer = &H40   'The second X button is down

Public Const VT_EOM As String = "m" 'End of marker of a VT100 escape sequence

'We'll need point and GetCursorPos apis to handle mouse wheel
Public Type POINTAPI
  x As Long
  y As Long
End Type
#If Win64 Then
  Declare PtrSafe Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
  Declare PtrSafe Function apiWindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xpoint As Long, ByVal ypoint As Long) As LongPtr
  Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As LongPtr
  Declare PtrSafe Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
#Else
  Declare PtrSafe Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
  Declare PtrSafe Function apiWindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xpoint As Long, ByVal ypoint As Long) As Long
  Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As LongPtr
  Declare PtrSafe Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
#End If

'RECT is a classic API struct that we may use anywhere
Public Type RECT
  left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Function VT_ESC() As String
  VT_ESC = Chr$(27) & "["
End Function

Public Function FindConsoulLibrary() As Boolean
  'First look into current project path, or in \bin subdirectory
  Dim sLookupPath     As String
  Dim sFilename       As String
  Dim fOK             As Boolean
  Dim fInPath         As Boolean
  
  On Error GoTo FindConsoulLibrary_Err
  
#If Win64 Then
  Const CONSOUL_DLL_NAME As String = "consoul_010205_64.dll"
#Else
  Const CONSOUL_DLL_NAME As String = "consoul_010205_32.dll"
#End If
  'Look into project path
  sLookupPath = CurrentProject.Path
  sFilename = CombinePath(sLookupPath, CONSOUL_DLL_NAME)
  fOK = ExistFile(sFilename)
  If Not fOK Then
    'Look into \bin subdir
    sLookupPath = CombinePath(sLookupPath, "bin")
    sFilename = CombinePath(sLookupPath, CONSOUL_DLL_NAME)
    fOK = ExistFile(sFilename)
    If Not fOK Then
      Dim iPathCt     As Integer
      Dim asPath()    As String
      Dim i           As Integer
      iPathCt = SplitString(asPath(), Environ$("PATH"), ";")
      For i = 1 To iPathCt
        sLookupPath = asPath(i)
        sFilename = CombinePath(sLookupPath, CONSOUL_DLL_NAME)
        fOK = ExistFile(sFilename)
        If fOK Then
          fInPath = True
          Exit For
        End If
      Next i
    End If
  End If

  If fOK Then
    If Not fInPath Then
      Debug.Print "Found, chdir to : "; sLookupPath
    Else
      Debug.Print CONSOUL_DLL_NAME; " found in PATH : "; sLookupPath
    End If
    ChDir sLookupPath
  Else
    ShowUFError "Failed to load the Consoul Library", "Cannot find " & CONSOUL_DLL_NAME
  End If
  
FindConsoulLibrary_Exit:
  FindConsoulLibrary = fOK
  Exit Function

FindConsoulLibrary_Err:
  ShowUFError "Error when looking for Consoul library", Err.Description
  Resume FindConsoulLibrary_Exit
End Function

'Wrap supported VT100 escapes in global VT_ functions
'VT escape codes: https://en.wikipedia.org/wiki/ANSI_escape_code

Public Function VT_RESET() As String
  VT_RESET = VT_ESC() & "0" & VT_EOM
End Function

Public Function VT_INV_ON() As String
  VT_INV_ON = VT_ESC() & "7" & VT_EOM
End Function

Public Function VT_INV_OFF() As String
  VT_INV_OFF = VT_ESC() & "27" & VT_EOM
End Function

Public Function VT_BOLD_ON() As String
  VT_BOLD_ON = VT_ESC() & "1" & VT_EOM
End Function

Public Function VT_BOLD_OFF() As String
  VT_BOLD_OFF = VT_ESC() & "21" & VT_EOM
End Function

Public Function VT_ITAL_ON() As String
  VT_ITAL_ON = VT_ESC() & "3" & VT_EOM
End Function

Public Function VT_ITAL_OFF() As String
  VT_ITAL_OFF = VT_ESC() & "23" & VT_EOM
End Function

Public Function VT_UNDL_ON() As String
  VT_UNDL_ON = VT_ESC() & "4" & VT_EOM
End Function

Public Function VT_UNDL_OFF() As String
  VT_UNDL_OFF = VT_ESC() & "24" & VT_EOM
End Function

Public Function VT_STRIKE_ON() As String
  VT_STRIKE_ON = VT_ESC() & "9" & VT_EOM
End Function

Public Function VT_STRIKE_OFF() As String
  VT_STRIKE_OFF = VT_ESC() & "29" & VT_EOM
End Function

Function VT_FCOLOR(ByVal plColor As Long) As String
  VT_FCOLOR = VT_ESC() & "38;$" & Hex$(plColor) & VT_EOM
End Function

Function VT_BCOLOR(ByVal plColor As Long) As String
  VT_BCOLOR = VT_ESC() & "48;$" & Hex$(plColor) & VT_EOM
End Function

Function VTX_SETWIDTH(ByVal piNextWidth As Integer) As String
  VTX_SETWIDTH = VT_ESC() & "100;" & piNextWidth & VT_EOM
End Function

Function VTX_ADVANCE_ABS(ByVal piAdvancePixels As Integer) As String
  VTX_ADVANCE_ABS = VT_ESC() & "101;" & piAdvancePixels & VT_EOM
End Function

Function VTX_ADVANCE_REL(ByVal piAdvancePixels As Integer) As String
  'If piAdvancePixels <> 0 Then
    VTX_ADVANCE_REL = VT_ESC() & "102;" & piAdvancePixels & VT_EOM
  'End If
End Function

Function VTX_SAVEPOS() As String
  VTX_SAVEPOS = VT_ESC() & "103" & VT_EOM
End Function

Function VTX_RESTOREPOS() As String
  VTX_RESTOREPOS = VT_ESC() & "104" & VT_EOM
End Function

Public Function VTX_SPILL(ByVal pfOnOff As Boolean) As String
  VTX_SPILL = VT_ESC() & "111;" & Abs(CInt(pfOnOff)) & VT_EOM
End Function

Public Function VTX_DTFLAGS(ByVal plDTFlags As Long) As String
  VTX_DTFLAGS = VT_ESC() & "110;$" & Hex$(plDTFlags) & VT_EOM
End Function

Public Function VT_NOOP() As String
  VT_NOOP = VT_ESC() & VT_EOM
End Function

' VTX_ for escape codes that are not VT100 standard, but devinfo.net extensions
Function VTX_ZONE_BEGIN(ByVal piZoneID As Integer, Optional ByVal psZoneTag As String = "") As String
  Dim sRet      As String
  sRet = VT_ESC() & "98;" & piZoneID
  If Len(psZoneTag) = 0 Then
    sRet = sRet & VT_EOM
  Else
    Dim iFind   As Integer
    iFind = InStr(1, psZoneTag, VT_EOM, vbBinaryCompare)
    While iFind > 0
      Mid$(psZoneTag, iFind, 1) = Chr$(1)
      iFind = InStr(iFind + 1, psZoneTag, VT_EOM, vbBinaryCompare)
    Wend
    sRet = sRet & ":" & psZoneTag & VT_EOM
  End If
  VTX_ZONE_BEGIN = sRet
End Function

Function VTX_ZONE_END(ByVal piZoneID As Integer) As String
  VTX_ZONE_END = VT_ESC() & "99;" & piZoneID & VT_EOM
End Function

Public Function VT_RealLen(ByVal psText As String) As Long
  Dim iCSI      As Long
  Dim iEndCSI   As Long
  Dim iOrigLen  As Long
  Dim iLen      As Long
  iOrigLen = Len(psText)
  iCSI = InStr(1, psText, VT_ESC())
  Do While iCSI > 0
    iEndCSI = InStr(iCSI + 2, psText, VT_EOM)
    If iEndCSI > 0 Then
      iLen = iLen + iEndCSI - iCSI + 1
    End If
    iCSI = InStr(iCSI + 1, psText, VT_ESC())
  Loop
  VT_RealLen = iOrigLen - iLen
End Function

'VT100 version of splitting a string into an array.
'VT100 Escapes are preserved and we don't split inside them.
'VT100 closing escapes and reset are preserved at end of a word.
Public Function VT_SplitString(ByRef asRetItem() As String, ByVal psText As String, ByVal psSep As String) As Long
  Dim lRealLen  As Long
  Dim i         As Long
  
  Dim iEsc      As Long
  Dim iPos      As String
  Dim sChar     As String
  Dim sWord     As String
  Dim fInEscSeq As Boolean
  Dim iTextLen  As Long
  Dim iSrc      As Long
  Dim iDst      As Long
  Dim sEsc      As String
  Dim fKeep     As Boolean
  Dim iWordCt   As Long
  Dim iBack     As Long
  
  On Error Resume Next
  Erase asRetItem
  iTextLen = Len(psText)
  If iTextLen = 0 Then Exit Function
  
  'if there no vt100 escapes then do a normal split
  If InStr(1, psText, VT_ESC()) = 0 Then
    VT_SplitString = SplitString(asRetItem(), psText, psSep)
    Exit Function
  End If
  
  'we go char by char
  lRealLen = VT_RealLen(psText)
  
  sWord = Space$(iTextLen)  'we'll poke into that and trim if needed
  iDst = 1&: iSrc = 1&
  Do
    sChar = Mid$(psText, iSrc, 1) 'Note: VBA.Mid$ doesn't fail if we go off
    If fInEscSeq Then
      If sChar = VT_EOM Then fInEscSeq = False
    Else
      If sChar = Chr$(27) Then
        If Mid$(psText, iSrc, Len(VT_ESC())) = VT_ESC() Then
          fInEscSeq = True
        End If
      End If
    End If
    
    If sChar <> psSep Then
      Mid$(sWord, iDst, 1) = sChar
      iDst = iDst + 1&
    Else
      'save word and prepare for next
      iWordCt = iWordCt + 1&
      ReDim Preserve asRetItem(1 To iWordCt) As String
      asRetItem(iWordCt) = left$(sWord, iDst - 1&)
      sWord = Space$(iTextLen - iSrc)
      iDst = 1
      
      iBack = 0&
      If iSrc < iTextLen Then
        iSrc = iSrc + 1: iBack = 1
      Do While Mid$(psText, iSrc, Len(VT_ESC)) = VT_ESC
        i = InStr(iSrc + Len(VT_ESC), psText, VT_EOM)
        If i > 0& Then
          fKeep = False
          'the escapes sequences that close another one or the reset are codes:
          ' 99,29,24,23,31,27, and 0
          ' 99 (zone end) is followed by ";", the others by "m" (VT_EOM)
          sEsc = Mid$(psText, iSrc, i - iSrc + 1&)
          sEsc = Right$(sEsc, Len(sEsc) - Len(VT_ESC))
          Select Case left$(sEsc, 2)
          Case "0" & VT_EOM
            fKeep = True
          Case "99"
            If Mid$(sEsc, 3, 1) = ";" Then fKeep = True
          Case Else
            If (left$(sEsc, 1) = "2") And (Mid$(sEsc, 3, 1) = VT_EOM) Then
              fKeep = True
            End If
          End Select
          If fKeep Then
            sWord = sWord & Mid$(psText, iSrc, i - iSrc + 1&)
            iSrc = i + 1&
            iBack = 1&
          Else
            Exit Do
          End If
        Else
          Exit Do
        End If
      Loop
        iSrc = iSrc - iBack
      End If
      
    End If
  
    iSrc = iSrc + 1&
  Loop Until (iSrc > iTextLen)
  If iDst > 1& Then
    iWordCt = iWordCt + 1&
    ReDim Preserve asRetItem(1 To iWordCt) As String
    asRetItem(iWordCt) = left$(sWord, iDst - 1&)
  End If
  
  VT_SplitString = iWordCt
End Function

'Purge psText of all vt100 escapes.
Public Function VT_Purge(ByVal psText As String) As String
  Dim lRealLen  As Long
  Dim i         As Long
  
  Dim iEsc      As Integer
  Dim sChar     As String
  Dim sLeft     As String
  Dim fInEscSeq As Boolean
  Dim iTextLen  As Integer
  Dim iSrc      As Integer
  Dim iDst      As Integer
  
  If InStr(1, psText, VT_ESC()) = 0 Then
    VT_Purge = psText
    Exit Function
  End If
  
  'we go char by char
  iTextLen = Len(psText)
  iDst = 1: iSrc = 1
  Do
    sChar = Mid$(psText, iSrc, 1) 'Note: VBA.Mid$ doesn't fail if we go off
    If fInEscSeq Then
      If sChar = VT_EOM Then fInEscSeq = False
    Else
      If sChar <> Chr$(27) Then
        sLeft = sLeft & sChar
      Else
        If Mid$(psText, iSrc, Len(VT_ESC())) = VT_ESC() Then
          fInEscSeq = True
        Else
          sLeft = sLeft & sChar
        End If
      End If
    End If
    
    iSrc = iSrc + 1
    iDst = iDst + 1
  Loop Until iSrc > iTextLen
  
  VT_Purge = sLeft
End Function

'Take piChartCt on the left of psText, keeping the vt100 escapes.
'VT100 closing escapes and reset are preserved at end of a word.
Public Function VT_Left(ByVal psText As String, ByVal piCharCt As Integer) As String
  Dim lRealLen  As Long
  Dim i         As Long
  
  Dim iEsc      As Integer
  Dim iPos      As String
  Dim sChar     As String
  Dim sLeft     As String
  Dim fInEscSeq As Boolean
  Dim iTextLen  As Integer
  Dim iSrc      As Integer
  Dim iDst      As Integer
  Dim iDstCount As Integer
  Dim sEsc      As String
  Dim fKeep     As Boolean
  
  If piCharCt < 1 Then Exit Function
  
  If InStr(1, psText, VT_ESC()) = 0 Then
    VT_Left = left$(psText, piCharCt)
    Exit Function
  End If
  
  lRealLen = VT_RealLen(psText)
  If lRealLen <= piCharCt Then
    VT_Left = psText
    Exit Function
  End If
  
  'we go char by char
  iTextLen = Len(psText)
  sLeft = Space$(iTextLen)  'we'll poke into that and trim if needed
  iDst = 1: iSrc = 1: iDstCount = 0
  Do
    sChar = Mid$(psText, iSrc, 1) 'Note: VBA.Mid$ doesn't fail if we go off
    If fInEscSeq Then
      If sChar = VT_EOM Then fInEscSeq = False
    Else
      If sChar <> Chr$(27) Then
        iDstCount = iDstCount + 1
      Else
        If Mid$(psText, iSrc, Len(VT_ESC())) = VT_ESC() Then
          fInEscSeq = True
        Else
          iDstCount = iDstCount + 1
        End If
      End If
    End If
    
    Mid$(sLeft, iDst, 1) = sChar
    iSrc = iSrc + 1
    iDst = iDst + 1
  Loop Until (iSrc > iTextLen) Or (iDstCount = piCharCt)
  
  sLeft = left$(sLeft, iDst - 1)
  
  'advance over any closing trailing VT escapes, and/or cut after a VT reset.
  'the smallest escape is reset, with 4 chars (reset).
  If (iTextLen - iSrc) > (Len(VT_ESC) + 1) Then
    Do While Mid$(psText, iSrc, Len(VT_ESC)) = VT_ESC
      i = InStr(iSrc + Len(VT_ESC), psText, VT_EOM)
      If i > 0 Then
        fKeep = False
        'the escapes sequences that close another one or the reset are codes:
        ' 99,29,24,23,31,27, and 0
        ' 99 (zone end) is followed by ";", the others by "m"
        sEsc = Mid$(psText, iSrc, i - iSrc + 1)
        sEsc = Right$(sEsc, Len(sEsc) - Len(VT_ESC))
        Select Case left$(sEsc, 2)
        Case "0" & VT_EOM
          fKeep = True
        Case "99"
          If Mid$(sEsc, 3, 1) = ";" Then fKeep = True
        Case Else
          If (left$(sEsc, 1) = "2") And (Mid$(sEsc, 3, 1) = VT_EOM) Then
            fKeep = True
          End If
        End Select
        If fKeep Then
          sLeft = sLeft & Mid$(psText, iSrc, i - iSrc + 1)
          iSrc = i + 1
        Else
          Exit Do
        End If
      Else
        Exit Do
      End If
    Loop
  End If
  
  VT_Left = sLeft
End Function

Public Function VT_RPad(ByVal s As String, ByVal PadChar As String, ByVal iLen As Integer) As String
  Dim lRealLen As Long
  lRealLen = VT_RealLen(s)
  If lRealLen < iLen Then
    VT_RPad = s & String$(iLen - lRealLen, Asc(PadChar))
  Else
    VT_RPad = s
  End If
End Function

Public Function VT_WrapToString(ByVal psText As String, Optional ByVal piCols As Integer = 80, Optional ByVal psLineSep As String = vbCrLf) As String
  Dim asWord()      As String
  Dim lWordCt       As Long
  Dim sRow          As String
  Dim sLeft         As String
  Dim i             As Long
  Dim sOutput       As String
  
  lWordCt = VT_SplitString(asWord(), psText, " ")
  sRow = ""
  For i = 1 To lWordCt
    If Len(sRow) = 0 Then
      sRow = asWord(i)
    ElseIf VT_RealLen(sRow & " " & asWord(i)) <= piCols Then
      sRow = sRow & " " & asWord(i)
    Else
      If Len(sOutput) > 0 Then sOutput = sOutput & psLineSep
      sOutput = sOutput & sRow
      sRow = asWord(i)
    End If
  Next i
  
  If Len(sRow) > 0 Then
    If Len(sOutput) > 0 Then sOutput = sOutput & psLineSep
    sOutput = sOutput & sRow
  End If
  
  VT_WrapToString = sOutput
End Function

Public Function VT_WrapToArray( _
  ByRef pasRetLines() As String, _
  ByVal psText As String, _
  Optional ByVal piCols As Integer = 80, _
  Optional ByVal psLineSep As String = vbCrLf) As Long
  
  Dim asWord()      As String
  Dim lWordCt       As Long
  Dim sRow          As String
  Dim sLeft         As String
  Dim i             As Long
  Dim iPart         As Long
  Dim iPartCt       As Long
  Dim asPart()      As String
  Dim iFinalLineCt  As Long
  
  If Len(psText) = 0 Then Exit Function
  
  iPartCt = VT_SplitString(asPart(), psText, psLineSep)
  If iPartCt > 0& Then
    For iPart = 1 To iPartCt
    
      lWordCt = VT_SplitString(asWord(), asPart(iPart), " ")
      sRow = ""
      For i = 1 To lWordCt
        If Len(sRow) = 0 Then
          sRow = asWord(i)
        ElseIf VT_RealLen(sRow & " " & asWord(i)) <= piCols Then
          sRow = sRow & " " & asWord(i)
        Else
          iFinalLineCt = iFinalLineCt + 1&
          ReDim Preserve pasRetLines(1 To iFinalLineCt) As String
          pasRetLines(iFinalLineCt) = sRow
          sRow = asWord(i)
        End If
      Next i
      
      If Len(sRow) > 0 Then
        iFinalLineCt = iFinalLineCt + 1&
        ReDim Preserve pasRetLines(1 To iFinalLineCt) As String
        pasRetLines(iFinalLineCt) = sRow
      End If
      
    Next iPart
  End If
  
  VT_WrapToArray = iFinalLineCt
End Function

'Fold() will wrap (but not word wrap) the text on piCols, cutting it into lines of piCols characters.
Public Function VT_FoldToString(ByVal psText As String, Optional ByVal piCols As Integer = 80, Optional ByVal psLineSep As String = vbCrLf) As String
  Dim asWord()      As String
  Dim lWordCt       As Long
  Dim sRow          As String
  Dim i             As Long
  Dim sOutput       As String
  Dim sLeft         As String
  Dim sRight        As String
  Dim iDiff         As Integer
  
  lWordCt = VT_SplitString(asWord(), psText, " ")
  
  sRow = ""
  For i = 1 To lWordCt
    If Len(sRow) = 0 Then
      If VT_RealLen(asWord(i)) <= piCols Then
        sRow = asWord(i)
      Else
        sLeft = VT_Left$(asWord(i), piCols)
        sRow = sLeft
        asWord(i) = Right$(asWord(i), Len(asWord(i)) - Len(sLeft))
        i = i - 1 'do it again, with the cut word
      End If
    ElseIf VT_RealLen(sRow & " " & asWord(i)) <= piCols Then
      sRow = sRow & " " & asWord(i)
    Else
      iDiff = VT_RealLen(sRow & " " & asWord(i)) - piCols
      sLeft = VT_Left$(asWord(i), VT_RealLen(asWord(i)) - iDiff)
      sRight = Right$(asWord(i), Len(asWord(i)) - Len(sLeft))
      If Len(sOutput) > 0 Then sOutput = sOutput & psLineSep
      sOutput = sOutput & sRow & " " & sLeft
      sRow = ""
      asWord(i) = sRight
      i = i - 1 'do it again, with the cut word
    End If
  Next i
  
  If Len(sRow) > 0 Then
    If Len(sOutput) > 0 Then sOutput = sOutput & psLineSep
    sOutput = sOutput & sRow
  End If
  
  VT_FoldToString = sOutput
End Function

Public Function VT_FoldToArray( _
  ByRef pasRetLines() As String, _
  ByVal psText As String, _
  Optional ByVal piCols As Integer = 80, _
  Optional ByVal psLineSep As String = vbCrLf) As Long
  
  Dim asWord()      As String
  Dim lWordCt       As Long
  Dim sRow          As String
  Dim sLeft         As String
  Dim sRight        As String
  Dim iDiff         As Integer
  Dim i             As Long
  Dim iPart         As Long
  Dim iPartCt       As Long
  Dim asPart()      As String
  Dim iFinalLineCt  As Long
  
  If Len(psText) = 0 Then Exit Function
  
  iPartCt = VT_SplitString(asPart(), psText, psLineSep)
  If iPartCt > 0& Then
    For iPart = 1 To iPartCt
    
      lWordCt = VT_SplitString(asWord(), asPart(iPart), " ")
        
      sRow = ""
      For i = 1 To lWordCt
        If Len(sRow) = 0 Then
          If VT_RealLen(asWord(i)) <= piCols Then
            sRow = asWord(i)
          Else
            sLeft = VT_Left$(asWord(i), piCols)
            sRow = sLeft
            asWord(i) = Right$(asWord(i), Len(asWord(i)) - Len(sLeft))
            i = i - 1 'do it again, with the cut word
          End If
        ElseIf VT_RealLen(sRow & " " & asWord(i)) <= piCols Then
          sRow = sRow & " " & asWord(i)
        Else
          iDiff = VT_RealLen(sRow & " " & asWord(i)) - piCols
          sLeft = VT_Left$(asWord(i), VT_RealLen(asWord(i)) - iDiff)
          sRight = Right$(asWord(i), Len(asWord(i)) - Len(sLeft))
          
          iFinalLineCt = iFinalLineCt + 1&
          ReDim Preserve pasRetLines(1 To iFinalLineCt) As String
          pasRetLines(iFinalLineCt) = sRow & sLeft '" " & sLeft
          sRow = ""
          asWord(i) = sRight
          i = i - 1 'do it again, with the cut word
        End If
      
      Next i
        
      If Len(sRow) > 0 Then
        iFinalLineCt = iFinalLineCt + 1&
        ReDim Preserve pasRetLines(1 To iFinalLineCt) As String
        pasRetLines(iFinalLineCt) = sRow
      End If
      
    Next iPart
  End If
  
  VT_FoldToArray = iFinalLineCt
End Function


