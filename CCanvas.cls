VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mvForeColor     As Variant
Private mvBackColor     As Variant
Private mvPenForeColor  As Variant
Private mvPenBackColor  As Variant

Private msFontName      As String
Private miFontSize      As Integer

Private msFilename      As String

Private miSelStartRow     As Integer
Private miSelStartCol     As Integer
Private miSelEndRow       As Integer
Private miSelEndCol       As Integer

Public Enum eTransparencyMode
  eColorTransparency
  eAlphaTransparency
End Enum
Private meTransMode           As eTransparencyMode
Private mvTransColor          As Variant 'Long or null
Private miTransAlphaPct       As Integer

Private miLineSpacingTop      As Integer
Private miLineSpacingBottom   As Integer
Private miLinePaddingTop      As Integer
Private miLinePaddingBottom   As Integer
Private mfAutoAdjustWidth     As Boolean

Private moGrid            As CConsoleGrid

'Class error context
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

Private Sub Class_Initialize()
  Set moGrid = New CConsoleGrid
End Sub

Private Sub Class_Terminate()
  Set moGrid = Nothing
End Sub

Public Property Get ConsoleGrid() As CConsoleGrid
  Set ConsoleGrid = moGrid
End Property

Private Function IsValidColorParam(ByRef pvColorParam As Variant) As Boolean
  If Not IsNull(pvColorParam) Then
    If VarType(pvColorParam) <> vbLong Then
      Exit Function
    End If
  End If
  IsValidColorParam = True
End Function

Public Property Get BackColor() As Variant
  BackColor = mvBackColor
End Property

Public Property Let BackColor(ByVal pvBackColor As Variant)
  If IsValidColorParam(pvBackColor) Then
    mvBackColor = pvBackColor
  End If
End Property

Public Property Get ForeColor() As Variant
  ForeColor = mvForeColor
End Property

Public Property Let ForeColor(ByVal pvForeColor As Variant)
  If IsValidColorParam(pvForeColor) Then
    mvForeColor = pvForeColor
  End If
End Property

Public Property Get PenBackColor() As Variant
  PenBackColor = mvPenBackColor
End Property

Public Property Let PenBackColor(ByVal pvPenBackColor As Variant)
  If IsValidColorParam(pvPenBackColor) Then
    mvPenBackColor = pvPenBackColor
  End If
End Property

Public Property Get PenForeColor() As Variant
  PenForeColor = mvPenForeColor
End Property

Public Property Let PenForeColor(ByVal pvPenForeColor As Variant)
  If IsValidColorParam(pvPenForeColor) Then
    mvPenForeColor = pvPenForeColor
  End If
End Property

Public Property Get FontName() As String
  FontName = msFontName
End Property

Public Property Let FontName(ByVal psFontName As String)
  msFontName = psFontName
End Property

Public Property Get FontSize() As Integer
  FontSize = miFontSize
End Property

Public Property Let FontSize(ByVal piFontSize As Integer)
  miFontSize = piFontSize
End Property

Public Property Get Filename() As String
  Filename = msFilename
End Property

Public Property Let Filename(ByVal psFilename As String)
  msFilename = psFilename
End Property

Public Property Get SelStartRow() As Integer
  SelStartRow = miSelStartRow
End Property

Public Property Let SelStartRow(ByVal piSelStartRow As Integer)
  miSelStartRow = piSelStartRow
End Property

Public Property Get SelStartCol() As Integer
  SelStartCol = miSelStartCol
End Property

Public Property Let SelStartCol(ByVal piSelStartCol As Integer)
  miSelStartCol = piSelStartCol
End Property

Public Property Get SelEndRow() As Integer
  SelEndRow = miSelEndRow
End Property

Public Property Let SelEndRow(ByVal piSelEndRow As Integer)
  miSelEndRow = piSelEndRow
End Property

Public Property Get SelEndCol() As Integer
  SelEndCol = miSelEndCol
End Property

Public Property Let SelEndCol(ByVal piSelEndCol As Integer)
  miSelEndCol = piSelEndCol
End Property

Public Property Get TransparentAlphaPct() As Integer
  TransparentAlphaPct = miTransAlphaPct
End Property

Public Property Let TransparentAlphaPct(ByVal piPct As Integer)
  If (piPct >= 0) And (piPct <= 100) Then
    miTransAlphaPct = piPct
  End If
End Property

Public Property Get TransparentColor() As Variant
  TransparentColor = mvTransColor
End Property

Public Property Let TransparentColor(ByVal pvTransColor As Variant)
  If IsValidColorParam(pvTransColor) Then
    mvTransColor = pvTransColor
  End If
End Property

Public Property Get TransparencyMode() As eTransparencyMode
  TransparencyMode = meTransMode
End Property

Public Property Let TransparencyMode(ByVal peTransMode As eTransparencyMode)
  meTransMode = peTransMode
End Property

Public Property Get AutoAdjustOnCharWidth() As Boolean
  AutoAdjustOnCharWidth = mfAutoAdjustWidth
End Property

Public Property Let AutoAdjustOnCharWidth(ByVal pfAutoAjust As Boolean)
  mfAutoAdjustWidth = pfAutoAjust
End Property

Public Property Get LineSpacingTop() As Integer
  LineSpacingTop = miLineSpacingTop
End Property

Public Property Let LineSpacingTop(ByVal piPixels As Integer)
  miLineSpacingTop = piPixels
End Property

Public Property Get LineSpacingBottom() As Integer
  LineSpacingBottom = miLineSpacingBottom
End Property

Public Property Let LineSpacingBottom(ByVal piPixels As Integer)
  miLineSpacingBottom = piPixels
End Property

Public Property Get LinePaddingTop() As Integer
  LinePaddingTop = miLinePaddingTop
End Property

Public Property Let LinePaddingTop(ByVal piPixels As Integer)
  miLinePaddingTop = piPixels
End Property

Public Property Get LinePaddingBottom() As Integer
  LinePaddingBottom = miLinePaddingBottom
End Property

Public Property Let LinePaddingBottom(ByVal piPixels As Integer)
  miLinePaddingBottom = piPixels
End Property

Public Function LoadFromNativeStream( _
    ByVal phStream As Integer, _
    ByRef piiProgress As IProgressIndicator, _
    Optional ByVal pfAutoResize As Boolean = True _
  ) As Boolean
  Dim fOK         As Boolean
  Dim iTemp       As Integer
  
  Dim fLostBool   As Boolean
  Dim lLostLong   As Long
  Dim iLostInt    As Integer
  Dim sLostString As String
  Dim lValue      As Long
  
  On Error GoTo LoadFromNativeStream_Err
  Const LOCAL_ERR_CTX As String = "LoadFromNativeStream"
  
  piiProgress.SetMax 2
  piiProgress.SetValue 1
  
  fOK = moGrid.LoadFromNativeStream(phStream, pfAutoResize:=pfAutoResize)
  If Not fOK Then
    SetErr LOCAL_ERR_CTX, -1&, moGrid.IIClassError.LastErrDesc
    GoTo LoadFromNativeStream_Exit
  End If
  'load additional info
  If Not EOF(phStream) Then
    Get #phStream, , lValue
    mvForeColor = lValue
    Get #phStream, , lValue
    mvBackColor = lValue
    Get #phStream, , miFontSize
    msFontName = FileGetUnicodeString(phStream)
    'V02.00.00 load canvas transparency and background picture filename (if any)
    piiProgress.SetValue 2
    'We test for EOF so we can still import old versions of the file format
    If Not EOF(phStream) Then
      'V002.00.06:
      'Transparency mode is a UI setting that is no more saved.
      'Setting was either 0 or 1, corresponding to the chkTransparency value in "Canvas" Form.
      'This code was previously in the Form_Canvas code, it's now incorporated in this class.
      'To keep compatibility with old versions, we have to read it for nothing.
      Get #phStream, , fLostBool 'We do nothing with that
      
      Get #phStream, , iTemp
      meTransMode = iTemp
      If meTransMode > eAlphaTransparency Then
        If Not piiProgress.Console Is Nothing Then
          piiProgress.Console.OutputLn LOCAL_ERR_CTX, "invalid value [" & meTransMode & "] at file offset [" & Seek(phStream) & "]"
          piiProgress.Console.OutputLn "Transparency mode can be either " & eAlphaTransparency & " (alpha) or " & eColorTransparency & " (color), set to alpha"
        End If
        meTransMode = eAlphaTransparency
      End If
      
      Get #phStream, , lValue
      mvTransColor = lValue
      Get #phStream, , miTransAlphaPct
      If miTransAlphaPct > 100 Then
        If Not piiProgress.Console Is Nothing Then
          piiProgress.Console.OutputLn LOCAL_ERR_CTX, "invalid value [" & miTransAlphaPct & "] at file offset [" & Seek(phStream) & "]"
          piiProgress.Console.OutputLn "Transparency percent must be in [0..100], set to 0 (zero)"
        End If
        miTransAlphaPct = 100
      End If
      
      'Previous version wrongfully saved the background image filename
      'msTransImageFilename = FileGetUnicodeString(phStream)
      'Read and lose
      sLostString = FileGetUnicodeString(phStream)
    End If
    If Not EOF(phStream) Then
      Get #phStream, , iTemp 'line padding top
      miLinePaddingTop = iTemp
      Get #phStream, , iTemp 'line padding top
      miLinePaddingBottom = iTemp
      Get #phStream, , iTemp 'line padding top
      miLineSpacingTop = iTemp
      Get #phStream, , iTemp 'line padding top
      miLineSpacingBottom = iTemp
      Get #phStream, , fLostBool
      mfAutoAdjustWidth = fLostBool
    End If
  End If
  
  LoadFromNativeStream = fOK
    
LoadFromNativeStream_Exit:
  Exit Function

LoadFromNativeStream_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LoadFromNativeStream_Exit
  Resume
End Function

Public Function LoadNative( _
    ByVal psFilename As String, _
    ByRef piiProgress As IProgressIndicator, _
    Optional ByVal pfAutoResize As Boolean = True _
  ) As Boolean
  Dim hStream     As Integer
  Dim fIsOpen     As Boolean
  Dim fOK         As Boolean
  
  On Error GoTo LoadNative_Err
  Const LOCAL_ERR_CTX As String = "LoadNative"
  
  'Open file for output (overwrite)
  hStream = FreeFile
  Open psFilename For Binary Access Read Lock Write As #hStream
  fIsOpen = True
  
  piiProgress.SetMax 2
  piiProgress.SetValue 1
  
  fOK = LoadFromNativeStream(hStream, piiProgress, pfAutoResize)
  
  'Close the file
  Close hStream
  fIsOpen = False
  
  LoadNative = fOK
    
LoadNative_Exit:
  If fIsOpen Then
    Close hStream
  End If
  Exit Function

LoadNative_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LoadNative_Exit
  Resume
End Function

Public Function SaveToNativeStream(ByVal phStream As Integer, ByRef piiProgress As IProgressIndicator) As Boolean
  On Error GoTo SaveToNativeStream_Err
  Const LOCAL_ERR_CTX As String = "SaveToNativeStream"
  
  Dim lValue    As Long
  Dim fLostBool As Boolean
  Dim fOK       As Boolean
  
  fOK = moGrid.SaveToNativeStream(phStream)
  If Not fOK Then
    SetErr LOCAL_ERR_CTX, -1&, moGrid.IIClassError.LastErrDesc
    GoTo SaveToNativeStream_Exit
  End If
  
  'append canvas colors
  lValue = CLng(Nz(mvForeColor, 0&))
  Put #phStream, , lValue
  lValue = CLng(Nz(mvBackColor, 0&))
  Put #phStream, , lValue
  'append font info
  Put #phStream, , miFontSize
  FilePutUnicodeString phStream, msFontName
  
  'V02.00.06 This is osbolete; was used to store the app transparency switch
  'We have to keep writing it for compatibility.
  Put #phStream, , fLostBool
  
  Put #phStream, , (CInt(meTransMode))
  lValue = CLng(Nz(mvTransColor, 0&))
  Put #phStream, , lValue
  Put #phStream, , miTransAlphaPct
  'V02.00.06 This is osbolete; was used to store the app bkgnd img filename
  'FilePutUnicodeString phStream, msTransImageFilename
  FilePutUnicodeString phStream, ""
  
  Put #phStream, , miLinePaddingTop
  Put #phStream, , miLinePaddingBottom
  Put #phStream, , miLineSpacingTop
  Put #phStream, , miLineSpacingBottom
  Put #phStream, , mfAutoAdjustWidth

  SaveToNativeStream = fOK
    
SaveToNativeStream_Exit:
  Exit Function

SaveToNativeStream_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume SaveToNativeStream_Exit
  Resume
End Function

Public Function SaveNative(ByVal psFilename As String, ByRef piiProgress As IProgressIndicator) As Boolean
  Dim hStream      As Integer
  Dim fIsOpen As Boolean
  Dim fOK         As Boolean
  Dim fBoolValue  As Boolean
  Dim iTemp       As Integer
  
  On Error GoTo SaveNative_Err
  
  'Open file for output (overwrite)
  hStream = FreeFile
  Open psFilename For Binary Access Write Lock Read Write As #hStream
  fIsOpen = True
  
  fOK = SaveToNativeStream(hStream, piiProgress)
  
  'Close the file
  Close hStream
  fIsOpen = False
  
  SaveNative = fOK
  
SaveNative_Exit:
  If fIsOpen Then
    Close hStream
  End If
  Exit Function

SaveNative_Err:
  ShowUFError "Failed to save file " & psFilename, Err.Description
  Resume SaveNative_Exit
End Function

