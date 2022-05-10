VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGenerateCodeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum eCodeTargetLanguage
  eVisualBasic
  eAutoIT
End Enum

Public Enum eRowGenerationMethod
  eFunctionCall
  eVariableConcatenation
End Enum

Private moConGrid                 As CConsoleGrid
Private meTargetLanguage          As eCodeTargetLanguage
Private mfPreserveRightSpaces     As Boolean
Private msWrapperMethodName       As String
Private meRowGenerationMethod     As eRowGenerationMethod
Private msCallMethodName          As String
Private msVariableName            As String
Private mfHexEscape               As Boolean
Private mlHexExclRangeStart       As Long
Private mlHexExclRangeEnd         As Long

'IDialog implmentation
Implements IDialog
Private msDialogID    As String
Private mfCancelled   As Boolean
Private mfModal       As Boolean

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

'
' IDialog implementation
'

Private Sub Class_Initialize()
  msDialogID = RegDialogClass(Me)
  mfPreserveRightSpaces = False
  mfHexEscape = True
  mlHexExclRangeStart = 32&
  mlHexExclRangeEnd = 128&
  msWrapperMethodName = "DisplayAsciiPaintImage"
End Sub

Private Sub Class_Terminate()
  UnregDialogClass msDialogID
End Sub

Private Property Let IDialog_Cancelled(ByVal pfCancelled As Boolean)
  mfCancelled = pfCancelled
End Property

Private Property Get IDialog_Cancelled() As Boolean
  IDialog_Cancelled = mfCancelled
End Property

Private Property Get IDialog_DialogID() As String
  IDialog_DialogID = msDialogID
End Property

Private Property Get IDialog_IsModal() As Boolean
  IDialog_IsModal = mfModal
End Property

Private Function IDialog_ShowDialog(ByVal pfShowModal As Boolean) As Boolean
  ClearErr
  On Error GoTo ShowDialog_Err
  
  mfCancelled = False
  
  'Create the dialog
  Dim sFormName As String
  mfModal = pfShowModal
  sFormName = GetGenCodeFormName()
  If pfShowModal Then
    DoCmd.OpenForm sFormName, acNormal, WindowMode:=acDialog, OpenArgs:=msDialogID
  Else
    DoCmd.OpenForm sFormName, acNormal, WindowMode:=acWindowNormal, OpenArgs:=msDialogID
  End If
  
  IDialog_ShowDialog = True
  Exit Function
ShowDialog_Err:
  SetErr "ShowDialog", Err.Number, Err.Description
End Function

'
' Public methods
'

Public Property Get IIDialog() As IDialog
  Set IIDialog = Me
End Property

Public Function DialogForm() As Form
  On Error Resume Next
  Set DialogForm = Forms(GetGenCodeFormName())
End Function

Public Property Get ConsoleGrid() As CConsoleGrid
  Set ConsoleGrid = moConGrid
End Property

Public Property Set ConsoleGrid(ByRef poConGrid As CConsoleGrid)
  Set moConGrid = poConGrid
End Property

Public Property Get TargetLanguage() As eCodeTargetLanguage
  TargetLanguage = meTargetLanguage
End Property

Public Property Let TargetLanguage(ByVal peTargetLanguage As eCodeTargetLanguage)
  meTargetLanguage = peTargetLanguage
End Property

Public Property Get WrapperMethodName() As String
  WrapperMethodName = msWrapperMethodName
End Property

Public Property Let WrapperMethodName(ByVal psMethodName As String)
  msWrapperMethodName = psMethodName
End Property

Public Property Get RowGenerationMethod() As eRowGenerationMethod
  RowGenerationMethod = meRowGenerationMethod
End Property

Public Property Let RowGenerationMethod(ByVal peMethod As eRowGenerationMethod)
  meRowGenerationMethod = peMethod
End Property

Public Property Get CallMethodName() As String
  CallMethodName = msCallMethodName
End Property

Public Property Let CallMethodName(ByVal psMethodName As String)
  msCallMethodName = psMethodName
End Property

Public Property Get VariableName() As String
  VariableName = msVariableName
End Property

Public Property Let VariableName(ByVal psVariableName As String)
  msVariableName = psVariableName
End Property

Public Property Get PreserveRightSpaces() As Boolean
  PreserveRightSpaces = mfPreserveRightSpaces
End Property

Public Property Let PreserveRightSpaces(ByVal pfPreserve As Boolean)
  mfPreserveRightSpaces = pfPreserve
End Property

Public Property Get HexEscape() As Boolean
  HexEscape = mfHexEscape
End Property

Public Property Let HexEscape(ByVal pfPreserve As Boolean)
  mfHexEscape = pfPreserve
End Property

Public Property Get HexExclRangeStart() As Long
  HexExclRangeStart = mlHexExclRangeStart
End Property

Public Property Let HexExclRangeStart(ByVal piStart As Long)
  If (piStart >= 0) And (piStart <= 65535) Then
    mlHexExclRangeStart = piStart
  End If
End Property

Public Property Get HexExclRangeEnd() As Long
  HexExclRangeEnd = mlHexExclRangeEnd
End Property

Public Property Let HexExclRangeEnd(ByVal piEnd As Long)
  If (piEnd >= 0) And (piEnd <= 65535) Then
    mlHexExclRangeEnd = piEnd
  End If
End Property

Public Function SaveProfile(ByVal psFilename As String) As Boolean
  Const LOCAL_ERR_CTX As String = "SaveProfile"
  On Error GoTo SaveProfile_Err
  ClearErr
  
  Dim fh          As Integer
  Dim fIsOpen     As Boolean
  
  fh = FreeFile
  Open psFilename For Output As #fh
  fIsOpen = True
  
  Print #fh, Trim$(CStr(meTargetLanguage))
  Print #fh, Trim$(CStr(CInt(mfPreserveRightSpaces)))
  Print #fh, msWrapperMethodName
  Print #fh, Trim$(CStr(meRowGenerationMethod))
  Print #fh, msCallMethodName
  Print #fh, msVariableName
  Print #fh, Trim$(CStr(CInt(mfHexEscape)))
  Print #fh, mlHexExclRangeStart
  Print #fh, mlHexExclRangeEnd
  
  SaveProfile = True
  
SaveProfile_Exit:
  If fIsOpen Then
    Close fh
    fIsOpen = False
  End If
  Exit Function
  
SaveProfile_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume SaveProfile_Exit
  Resume
End Function

Public Function LoadProfile(ByVal psFilename As String) As Boolean
  Const LOCAL_ERR_CTX As String = "LoadProfile"
  On Error GoTo LoadProfile_Err
  ClearErr
  
  Dim fh          As Integer
  Dim fIsOpen     As Boolean
  Dim sBuf        As String
  
  fh = FreeFile
  Open psFilename For Input Access Read As #fh
  fIsOpen = True
  
  Line Input #fh, sBuf
  meTargetLanguage = Val(sBuf)
  
  Line Input #fh, sBuf
  mfPreserveRightSpaces = CBool(Val(sBuf))
  
  Line Input #fh, sBuf
  msWrapperMethodName = sBuf
  
  Line Input #fh, sBuf
  meRowGenerationMethod = Val(sBuf)
  
  Line Input #fh, sBuf
  msCallMethodName = sBuf
  
  Line Input #fh, sBuf
  msVariableName = sBuf
  
  If Not EOF(fh) Then
    Line Input #fh, sBuf
    mfHexEscape = CBool(Val(sBuf))
  
    Line Input #fh, sBuf
    mlHexExclRangeStart = Val(sBuf)
  
    Line Input #fh, sBuf
    mlHexExclRangeEnd = Val(sBuf)
  Else
    mfHexEscape = False
    mlHexExclRangeStart = 32&
    mlHexExclRangeEnd = 128&
  End If
  
  LoadProfile = True
  
LoadProfile_Exit:
  If fIsOpen Then
    Close fh
    fIsOpen = False
  End If
  Exit Function
  
LoadProfile_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LoadProfile_Exit
  Resume
End Function

Private Function HexEscapeCharCode(ByVal piCode As Integer) As String
  Dim sHex    As String
  sHex = Hex$(piCode)
  If Len(sHex) < 4 Then
    sHex = String$(4 - Len(sHex), "0") & sHex
  End If
  HexEscapeCharCode = sHex
End Function

Private Function IsInDisplayRange(ByVal psChar As String) As Boolean
  Dim iAscW As Integer
  iAscW = AscW(psChar)
  IsInDisplayRange = CBool((iAscW >= Me.HexExclRangeStart) And (iAscW <= Me.HexExclRangeEnd))
End Function

Private Function EncodeConcatChar(ByVal psChar As String, ByRef pfRetIsFunc As Boolean) As String
  Dim iAscW As Integer
  
  pfRetIsFunc = False
  If Not Me.HexEscape Then
    EncodeConcatChar = psChar
    Exit Function
  End If
  
  If Not IsInDisplayRange(psChar) Then
    iAscW = AscW(psChar)
    pfRetIsFunc = True
    If meTargetLanguage = eAutoIT Then
      EncodeConcatChar = "ChrW(" & iAscW & ")"
    Else
      EncodeConcatChar = "ChrW$(&H" & Hex$(iAscW) & ")"
    End If
  Else
    EncodeConcatChar = psChar
  End If
End Function

Private Sub HexEncodeChars(ByRef psText As String)
  Dim i         As Integer
  Dim iLen      As Integer
  Dim sBuf      As String
  Dim sChar     As String
  Dim iAscW     As Integer
  Dim fIsFunc   As Boolean
  Dim sQuote    As String
  
  iLen = Len(psText)
  If iLen = 0 Then Exit Sub
  sQuote = Chr$(34)
  
  For i = 1 To iLen
    sChar = Mid$(psText, i, 1)
    iAscW = AscW(sChar)
    If IsInDisplayRange(sChar) Then
      sBuf = sBuf & sChar
    Else
      sChar = EncodeConcatChar(sChar, fIsFunc)
      If Not fIsFunc Then
        sBuf = sBuf & sChar
      Else
        If i > 1 Then
          sBuf = sBuf & sQuote & " & "
        End If
        sBuf = sBuf & sChar
        If i < iLen Then
          sBuf = sBuf & " & " & sQuote
        End If
      End If
    End If
  Next i
  psText = Replace(sBuf, sQuote & sQuote & " & ", "")
End Sub

Private Function QuoteString(ByVal psString As String) As String
  QuoteString = Chr$(34) & Replace(psString, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function MakeCodeRow_AutoIT(ByVal psVT100 As String) As String
  Dim sRet        As String
  Dim iTextLen    As Long
  Dim asChar()    As String
  Dim aiCount()   As Integer
  Dim iGroupCt    As Integer
  Dim i           As Integer
  Dim fOK         As Boolean
  Dim sLine       As String
  Dim fIsChrW     As Boolean
  Dim fInQuotes   As Boolean
  Dim sQuote      As String
  
  Const FUNCTION_CHRW  As String = "ChrW"  'AutoIT wide char function
  sQuote = Chr$(34)
  
  iTextLen = Len(psVT100)
  If iTextLen > 0 Then
    fOK = BreakTextToCharGroups(psVT100, asChar(), aiCount(), iGroupCt)
    If Not fOK Then GoTo MakeCodeRow_AutoIT_Exit
    
    For i = 1 To iGroupCt
      If aiCount(i) = 1 Then
        If Not fInQuotes Then
          If i > 1 Then
            sLine = sLine & " & "
          End If
          sLine = sLine & sQuote
        End If
        sLine = sLine & asChar(i)
        fInQuotes = True
      Else
        If fInQuotes Or (i > 1) Then
          If fInQuotes Then
            sLine = sLine & sQuote
          End If
          sLine = sLine & " & "
          fInQuotes = False
        End If
        sLine = sLine & "StringW(" & aiCount(i) & ", 0x" & Hex$(AscW(asChar(i))) & ")"
      End If
    Next i
    If Len(sLine) > 0 Then
      If fInQuotes Then
        sLine = sLine & sQuote
      End If
      If aiCount(1) = 1 Then
        If left$(sLine, 1) <> sQuote Then
          sLine = sQuote & sLine
        End If
      End If
      'Now we can replace char not in range if asked so
      HexEncodeChars sLine
      sRet = sRet & "$sTemp = " & sLine & vbCrLf
    End If
    
    If meRowGenerationMethod = eFunctionCall Then
      If Len(msCallMethodName) = 0 Then
        msCallMethodName = "ConOutLn"
      End If
      sRet = sRet & msCallMethodName & "($sTemp)" & vbCrLf
      sRet = sRet & "$sTemp = " & sQuote & sQuote & vbCrLf
    Else
      sRet = sRet & msVariableName & " = " & msVariableName & " & $sTemp & $EOL" & vbCrLf
    End If
    
  Else
    'blank line
    If meRowGenerationMethod = eFunctionCall Then
      If Len(msCallMethodName) = 0 Then
        msCallMethodName = "ConOutLn"
      End If
      sRet = sRet & msCallMethodName & "(" & sQuote & sQuote & ")" & vbCrLf
    Else
      sRet = sRet & msVariableName & " = " & msVariableName & " & $EOL" & vbCrLf
    End If
  End If

MakeCodeRow_AutoIT_Exit:
  MakeCodeRow_AutoIT = sRet
End Function

Private Function MakeCodeRow_VBA(ByVal psVT100 As String) As String
  Dim sRet        As String
  Dim iTextLen    As Long
  Dim asChar()    As String
  Dim aiCount()   As Integer
  Dim iGroupCt    As Integer
  Dim i           As Integer
  Dim fOK         As Boolean
  Dim sLine       As String
  Dim fIsChrW     As Boolean
  Dim fInQuotes   As Boolean
  Dim sQuote      As String
  
  Const FUNCTION_CHRW  As String = "ChrW"  'AutoIT wide char function
  sQuote = Chr$(34)
  
  iTextLen = Len(psVT100)
  If iTextLen > 0 Then
    fOK = BreakTextToCharGroups(psVT100, asChar(), aiCount(), iGroupCt)
    If Not fOK Then GoTo MakeCodeRow_VBA_Exit
    
    For i = 1 To iGroupCt
      If aiCount(i) = 1 Then
        If Not fInQuotes Then
          If i > 1 Then
            sLine = sLine & " & "
          End If
          sLine = sLine & sQuote
        End If
        sLine = sLine & asChar(i)
        fInQuotes = True
      Else
        If fInQuotes Or (i > 1) Then
          If fInQuotes Then
            sLine = sLine & sQuote
          End If
          sLine = sLine & " & "
          fInQuotes = False
        End If
        sLine = sLine & "String$(" & aiCount(i) & ", &H" & Hex$(AscW(asChar(i))) & ")"
      End If
    Next i
    If Len(sLine) > 0 Then
      If fInQuotes Then
        sLine = sLine & sQuote
      End If
      If aiCount(1) = 1 Then
        If left$(sLine, 1) <> sQuote Then
          sLine = sQuote & sLine
        End If
      End If
      'Now we can replace char not in range if asked so
      HexEncodeChars sLine
      sRet = sRet & "sTemp = " & sLine & vbCrLf
    End If
    
    If meRowGenerationMethod = eFunctionCall Then
      If Len(msCallMethodName) = 0 Then
        msCallMethodName = "ConOutLn"
      End If
      sRet = sRet & msCallMethodName & " sTemp" & vbCrLf
      sRet = sRet & "sTemp = " & sQuote & sQuote & vbCrLf
    Else
      sRet = sRet & msVariableName & " = " & msVariableName & " & sTemp & sEOL" & vbCrLf
    End If
  Else
    'blank line
    If meRowGenerationMethod = eFunctionCall Then
      If Len(msCallMethodName) = 0 Then
        msCallMethodName = "ConOutLn"
      End If
      sRet = sRet & msCallMethodName & " " & sQuote & sQuote & vbCrLf
    Else
      sRet = sRet & msVariableName & " = " & msVariableName & " & sEOL" & vbCrLf
    End If
  End If

MakeCodeRow_VBA_Exit:
  MakeCodeRow_VBA = sRet
End Function

Private Function MakeCodeRow(ByVal psVT100 As String) As String
  Dim iAscW       As Integer
  Dim sRet        As String
  
  Select Case meTargetLanguage
  Case eCodeTargetLanguage.eAutoIT
    sRet = MakeCodeRow_AutoIT(psVT100)
  Case eCodeTargetLanguage.eVisualBasic
    sRet = MakeCodeRow_VBA(psVT100)
  End Select
  
  MakeCodeRow = sRet
End Function

Private Function MakeCodeProcBegin() As String
  Dim sRet        As String
  
  Select Case meTargetLanguage
  Case eCodeTargetLanguage.eAutoIT
    sRet = "Func " & msWrapperMethodName & "()" & vbCrLf
    If meRowGenerationMethod = eVariableConcatenation Then
      sRet = sRet & "Local $iRepeat, $sTemp, $EOL, " & msVariableName & vbCrLf
      sRet = sRet & "$EOL = ""<TODO: put your EOL char(s) here>""" & vbCrLf
      sRet = sRet & msVariableName & " = """"" & vbCrLf
    Else
      sRet = sRet & "Local $iRepeat, $sTemp" & vbCrLf
    End If
    sRet = sRet & "$sTemp = """"" & vbCrLf
  Case eCodeTargetLanguage.eVisualBasic
    If meRowGenerationMethod = eVariableConcatenation Then
      sRet = "Function " & msWrapperMethodName & "() As String" & vbCrLf
      sRet = sRet & "Dim sTemp  As String" & vbCrLf
      sRet = sRet & "Dim sRes   As String" & vbCrLf
      sRet = sRet & "Dim sEOL   As String" & vbCrLf
      sRet = sRet & "sEOL = vbCrLf" & vbCrLf
      sRet = sRet & msVariableName & " = """"" & vbCrLf
    Else
      If msCallMethodName <> "" Then
        sRet = "Sub " & Me.WrapperMethodName & "()" & vbCrLf
      Else
        sRet = "Sub " & Me.WrapperMethodName & "()" & vbCrLf
      End If
    End If
  End Select

  MakeCodeProcBegin = sRet
End Function

Private Function MakeCodeProcEnd() As String
  Dim sRet        As String
  
  Select Case meTargetLanguage
  Case eCodeTargetLanguage.eAutoIT
    If meRowGenerationMethod = eVariableConcatenation Then
      sRet = "Return " & msVariableName & vbCrLf
    End If
    sRet = sRet & "EndFunc" & vbCrLf
  Case eCodeTargetLanguage.eVisualBasic
    If meRowGenerationMethod = eVariableConcatenation Then
      sRet = msWrapperMethodName & " = " & msVariableName & vbCrLf
      sRet = sRet & "End Function" & vbCrLf
    Else
      sRet = sRet & "End Sub" & vbCrLf
    End If
  End Select

  MakeCodeProcEnd = sRet
End Function

Public Function GenerateCode() As String
  Const LOCAL_ERR_CTX As String = "GenerateCode"
  On Error GoTo GenerateCode_Err
  ClearErr
  
  Dim sRes        As String 'The generated code that we'll return
  Dim iRow        As Integer
  Dim sLine       As String
  Dim sCodeRow    As String
  
  If moConGrid Is Nothing Then
    SetErr LOCAL_ERR_CTX, -1&, "The grid is empty"
    Exit Function
  End If
  
  'We'll break each row in groups of similar characters.
  'Then we'll cal a language specific generator on those groups.
  sRes = MakeCodeProcBegin()
  If moConGrid.Rows > 0 Then
    For iRow = 1 To moConGrid.Rows
      sLine = moConGrid.GetLineVT100(iRow, False)
      If Not mfPreserveRightSpaces Then
        sLine = RTrim(sLine)
      End If
      sCodeRow = MakeCodeRow(sLine)
      sRes = sRes & sCodeRow
    Next iRow
  End If
  sRes = sRes & MakeCodeProcEnd()
  
  GenerateCode = sRes
GenerateCode_Exit:
  Exit Function
  
GenerateCode_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume GenerateCode_Exit
  Resume
End Function

Public Function BreakTextToCharGroups( _
    ByVal psText As String, _
    ByRef pasRetCharGroup() As String, _
    ByRef paiRetCharGroupCount() As Integer, _
    ByRef piRetGroupCount As Integer _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "BreakTextToCharGroups"
  
  On Error Resume Next
  Erase pasRetCharGroup
  Erase paiRetCharGroupCount
  piRetGroupCount = 0
  
  On Error GoTo BreakTextToCharGroups_Err
  ClearErr
  
  Dim iTextLen        As Integer
  Dim i               As Integer
  Dim sChar           As String
  Dim fCreateGroup    As Boolean
  Dim fInEscSeq       As Boolean
  
  iTextLen = Len(psText)
  If iTextLen = 0 Then
    BreakTextToCharGroups = True
    GoTo BreakTextToCharGroups_Exit
  End If
  'Dim to max that can happen
  ReDim pasRetCharGroup(1 To iTextLen) As String
  ReDim paiRetCharGroupCount(1 To iTextLen) As Integer
  
  'We have to keep consoul escape sequences together
  piRetGroupCount = 0
  For i = 1 To iTextLen
    sChar = Mid$(psText, i, 1)
    If Mid$(psText, i, 2) = VT_ESC() Then
      fInEscSeq = True
    End If
    If Not fInEscSeq Then
      If i > 1 Then
        If sChar = pasRetCharGroup(piRetGroupCount) Then
          paiRetCharGroupCount(piRetGroupCount) = paiRetCharGroupCount(piRetGroupCount) + 1
          fCreateGroup = False
        Else
          fCreateGroup = True
        End If
      Else
        fCreateGroup = True
      End If
      If fCreateGroup Then
        piRetGroupCount = piRetGroupCount + 1
        pasRetCharGroup(piRetGroupCount) = sChar
        paiRetCharGroupCount(piRetGroupCount) = 1
      End If
    Else
      'advance until end of seq marker and create groups all along
      piRetGroupCount = piRetGroupCount + 1
      pasRetCharGroup(piRetGroupCount) = sChar
      paiRetCharGroupCount(piRetGroupCount) = 1
      If sChar = VT_EOM Then
        fInEscSeq = False
      End If
    End If
  Next i
  
  'Adjust returned arrays sizes
  ReDim Preserve pasRetCharGroup(1 To piRetGroupCount) As String
  ReDim Preserve paiRetCharGroupCount(1 To piRetGroupCount) As Integer
  
  BreakTextToCharGroups = True
  
BreakTextToCharGroups_Exit:
  Exit Function

BreakTextToCharGroups_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume BreakTextToCharGroups_Exit
  Resume
  
End Function
